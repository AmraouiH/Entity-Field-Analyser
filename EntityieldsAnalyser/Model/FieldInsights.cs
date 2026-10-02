using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EntityieldsAnalyser
{
    /// <summary>A labelled value of a chart, with the Fields tab filter to apply when the user drills into it.</summary>
    public class InsightPoint
    {
        public InsightPoint(string label, double value, string filter = null, string tooltip = null)
        {
            Label = label;
            Value = value;
            Filter = filter;
            Tooltip = tooltip;
        }

        public string Label { get; }
        public double Value { get; }

        /// <summary>DataTable.Select expression on the fields table (null = not drillable).</summary>
        public string Filter { get; }
        public string Tooltip { get; }

        /// <summary>Optional secondary value (e.g. standard fields in a stacked column).</summary>
        public double Secondary { get; set; }
    }

    /// <summary>Weight of a field type in the physical table (same rule as the fields calculator).</summary>
    public enum ColumnWeight { Other = 1, OptionLike = 2, Lookup = 3 }

    /// <summary>
    /// Everything the charts need, computed once from the analysed fields.
    /// Pure computation: no UI, no service call.
    /// </summary>
    public class FieldInsights
    {
        public const int ColumnLimit = 1024;

        public FieldInsights(Dictionary<AttributeTypeCode, List<entityParam>> data, bool hasDataUsage)
        {
            Fields = data.Values.SelectMany(l => l).ToList();
            HasDataUsage = hasDataUsage;

            TypeCounts = data
                .Select(kv => new InsightPoint(kv.Key.ToString(), kv.Value.Count, "Type = '" + kv.Key + "'"))
                .OrderByDescending(p => p.Value).ToList();

            ColumnsByWeight = new Dictionary<ColumnWeight, int>
            {
                { ColumnWeight.Lookup, 0 }, { ColumnWeight.OptionLike, 0 }, { ColumnWeight.Other, 0 }
            };
            ColumnCostByType = new List<InsightPoint>();
            foreach (var kv in data)
            {
                var weight = WeightOf(kv.Key);
                var cost = kv.Value.Count * (int)weight;
                ColumnsByWeight[weight] += cost;
                ColumnCostByType.Add(new InsightPoint(kv.Key.ToString(), cost, "Type = '" + kv.Key + "'",
                    $"{kv.Key}: {kv.Value.Count} field(s) × {(int)weight} = {cost} columns"));
            }
            ColumnCostByType = ColumnCostByType.OrderByDescending(p => p.Value).ToList();
            ColumnsUsed = ColumnsByWeight.Values.Sum();

            BuildOrigin();
            BuildSettings();
            BuildPublishers();
            BuildTimeline();
            if (hasDataUsage)
                BuildUsage(data);
        }

        public List<entityParam> Fields { get; }
        public bool HasDataUsage { get; }
        public int Total => Fields.Count;
        public int CustomCount => Fields.Count(f => f.isCustom == true);
        public int UnmanagedCount => Fields.Count(f => f.isManaged == "Unmanaged");

        public List<InsightPoint> TypeCounts { get; }
        public List<InsightPoint> Origin { get; private set; }
        public List<InsightPoint> Settings { get; private set; }
        public List<InsightPoint> Publishers { get; private set; }
        public List<InsightPoint> Timeline { get; private set; }
        public string TimelineGranularity { get; private set; }

        public Dictionary<ColumnWeight, int> ColumnsByWeight { get; }
        public List<InsightPoint> ColumnCostByType { get; }
        public int ColumnsUsed { get; }
        public int ColumnsAvailable => Math.Max(0, ColumnLimit - ColumnsUsed);

        // Data usage
        public List<InsightPoint> FillDistribution { get; private set; } = new List<InsightPoint>();
        public List<InsightPoint> FillByType { get; private set; } = new List<InsightPoint>();
        public List<InsightPoint> LeastUsedCustom { get; private set; } = new List<InsightPoint>();
        public int UnusedCount { get; private set; }
        public int LowUsageCount { get; private set; }
        public double AverageFillRate { get; private set; }

        public static ColumnWeight WeightOf(AttributeTypeCode type)
        {
            switch (type)
            {
                case AttributeTypeCode.Owner:
                case AttributeTypeCode.Lookup:
                    return ColumnWeight.Lookup;
                case AttributeTypeCode.Status:
                case AttributeTypeCode.Boolean:
                case AttributeTypeCode.Picklist:
                case AttributeTypeCode.Money:
                    return ColumnWeight.OptionLike;
                default:
                    return ColumnWeight.Other;
            }
        }

        public static string Escape(string value) => (value ?? string.Empty).Replace("'", "''");

        private void BuildOrigin()
        {
            var standard = Fields.Count(f => f.isCustom != true && f.isManaged == "Managed");
            var standardUnmanaged = Fields.Count(f => f.isCustom != true && f.isManaged == "Unmanaged");
            var customManaged = Fields.Count(f => f.isCustom == true && f.isManaged == "Managed");
            var customUnmanaged = Fields.Count(f => f.isCustom == true && f.isManaged == "Unmanaged");

            Origin = new List<InsightPoint>
            {
                new InsightPoint("Standard", standard, "IsCustom = false AND [Managed/Unmanaged] = 'Managed'", "Out-of-the-box fields"),
                new InsightPoint("Custom · managed", customManaged, "IsCustom = true AND [Managed/Unmanaged] = 'Managed'", "Custom fields deployed through a managed solution"),
                new InsightPoint("Custom · unmanaged", customUnmanaged, "IsCustom = true AND [Managed/Unmanaged] = 'Unmanaged'", "Custom fields living in the default / unmanaged layer"),
            };
            if (standardUnmanaged > 0)
                Origin.Add(new InsightPoint("Standard · unmanaged", standardUnmanaged, "IsCustom = false AND [Managed/Unmanaged] = 'Unmanaged'"));
        }

        private void BuildSettings()
        {
            Func<Func<entityParam, bool>, double> share = predicate => Total == 0 ? 0 : Fields.Count(predicate) * 100.0 / Total;
            Settings = new List<InsightPoint>
            {
                new InsightPoint("Auditing enabled", share(f => f.isAuditable == true), "IsAuditable = true", "Fields whose changes are audited"),
                new InsightPoint("In Advanced Find", share(f => f.isSearchable == true), "IsSearchable = true", "Fields available in Advanced Find / views"),
                new InsightPoint("Required", share(f => f.requiredLevel != null && f.requiredLevel != "None"), "[Required Level] <> 'None'", "Business or system required fields"),
                new InsightPoint("Field security", share(f => f.IsSecured == true), "IsSecured = true", "Fields protected by column (field) level security"),
            };
        }

        private void BuildPublishers()
        {
            var groups = Fields
                .Where(f => f.isCustom == true && !string.IsNullOrEmpty(f.fieldName) && f.fieldName.IndexOf('_') > 0)
                .GroupBy(f => f.fieldName.Substring(0, f.fieldName.IndexOf('_')))
                .Select(g => new InsightPoint(g.Key + "_", g.Count(), "IsCustom = true AND SchemaName LIKE '" + Escape(g.Key) + "[_]*'", "Custom fields with the \"" + g.Key + "_\" publisher prefix"))
                .OrderByDescending(p => p.Value).ToList();

            const int max = 7;
            if (groups.Count > max)
            {
                var others = groups.Skip(max).Sum(p => p.Value);
                groups = groups.Take(max).ToList();
                groups.Add(new InsightPoint("Other prefixes", others));
            }
            Publishers = groups;
        }

        private void BuildTimeline()
        {
            var dated = Fields.Where(f => f.dateOfCreation > new DateTime(2000, 1, 1)).ToList();
            Timeline = new List<InsightPoint>();
            if (dated.Count == 0)
                return;

            var min = dated.Min(f => f.dateOfCreation);
            var max = dated.Max(f => f.dateOfCreation);
            var byMonth = (max - min).TotalDays <= 730;
            TimelineGranularity = byMonth ? "per month" : "per year";

            Func<DateTime, DateTime> bucket = d => byMonth ? new DateTime(d.Year, d.Month, 1) : new DateTime(d.Year, 1, 1);
            var start = bucket(min);
            var end = bucket(max);
            for (var b = start; b <= end; b = byMonth ? b.AddMonths(1) : b.AddYears(1))
            {
                var next = byMonth ? b.AddMonths(1) : b.AddYears(1);
                var inBucket = dated.Where(f => f.dateOfCreation >= b && f.dateOfCreation < next).ToList();
                var label = byMonth ? b.ToString("MMM yy", CultureInfo.CurrentCulture) : b.Year.ToString(CultureInfo.InvariantCulture);
                var filter = string.Format(CultureInfo.InvariantCulture, "CreatedOn >= #{0:MM/dd/yyyy}# AND CreatedOn < #{1:MM/dd/yyyy}#", b, next);
                Timeline.Add(new InsightPoint(label, inBucket.Count(f => f.isCustom == true), filter)
                {
                    Secondary = inBucket.Count(f => f.isCustom != true)
                });
            }
        }

        private void BuildUsage(Dictionary<AttributeTypeCode, List<entityParam>> data)
        {
            UnusedCount = Fields.Count(f => f.percentageOfUse <= 0);
            LowUsageCount = Fields.Count(f => f.percentageOfUse > 0 && f.percentageOfUse < 5);
            AverageFillRate = Fields.Count == 0 ? 0 : Fields.Average(f => f.percentageOfUse);

            Func<double, double, string, string, InsightPoint> bucket = (from, to, label, filter) =>
                new InsightPoint(label, Fields.Count(f => f.percentageOfUse > from && f.percentageOfUse <= to), filter);
            FillDistribution = new List<InsightPoint>
            {
                new InsightPoint("0%", UnusedCount, "[Percentage Of Use] = 0", "Never filled in any record"),
                bucket(0, 1, "< 1%", "[Percentage Of Use] > 0 AND [Percentage Of Use] <= 1"),
                bucket(1, 10, "1–10%", "[Percentage Of Use] > 1 AND [Percentage Of Use] <= 10"),
                bucket(10, 50, "10–50%", "[Percentage Of Use] > 10 AND [Percentage Of Use] <= 50"),
                bucket(50, 90, "50–90%", "[Percentage Of Use] > 50 AND [Percentage Of Use] <= 90"),
                bucket(90, double.MaxValue, "> 90%", "[Percentage Of Use] > 90"),
            };

            FillByType = data
                .Select(kv => new InsightPoint(kv.Key.ToString(), kv.Value.Average(f => f.percentageOfUse), "Type = '" + kv.Key + "'",
                    $"{kv.Key}: average fill rate {kv.Value.Average(f => f.percentageOfUse):0.#}% over {kv.Value.Count} field(s)"))
                .OrderByDescending(p => p.Value).ToList();

            LeastUsedCustom = Fields
                .Where(f => f.isCustom == true && !f.IsPrimaryId.GetValueOrDefault() && !f.IsPrimaryName.GetValueOrDefault())
                .OrderBy(f => f.percentageOfUse).ThenBy(f => f.fieldName)
                .Take(12)
                .Select(f => new InsightPoint(f.fieldName, f.percentageOfUse, "SchemaName = '" + Escape(f.fieldName) + "'",
                    $"{f.displayName} ({f.fieldName}): {f.percentageOfUse:0.##}%"))
                .ToList();
        }
    }
}
