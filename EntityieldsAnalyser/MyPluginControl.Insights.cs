using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using EntityieldsAnalyser.UI;

namespace EntityieldsAnalyser
{
    /// <summary>
    /// Insights pages (Overview, Fields, Data usage, Capacity), navigation between them and chart drill-down.
    /// </summary>
    public partial class MyPluginControl
    {
        private const string DataUsageMode = "Metadata + Data usage";

        private enum Page { Overview = 0, Fields = 1, Usage = 2, Capacity = 3 }

        /// <summary>Filters behind the chips of the Fields page (same order as the chips).</summary>
        private static readonly string[] ChipFilters =
        {
            null,
            "IsCustom = true",
            "[Managed/Unmanaged] = 'Unmanaged'",
            "[Required Level] <> 'None'",
            "[Percentage Of Use] = 0",
        };

        private static readonly Color NeutralSeries = Color.FromArgb(196, 204, 214);

        private FieldInsights _insights;
        private string _analysedType;
        private string _drillFilter;
        private bool _suppressFieldFilter;

        private Chart chartTypes;
        private Chart chartOrigin;
        private Chart chartSettings;
        private Chart chartPublishers;
        private Chart chartTimeline;
        private Chart chartFillDistribution;
        private Chart chartFillByType;
        private Chart chartLeastUsed;
        private Chart chartBudget;
        private Chart chartColumnCost;

        #region Build
        private void BuildInsightsUi()
        {
            pivotTabs.AddRange("Overview", "Fields", "Data usage", "Capacity");

            #region Overview
            chartTypes = ChartStyler.Create(cardTypes, "No field.");
            ChartStyler.StyleBars(chartTypes, true);
            ChartStyler.AddSeries(chartTypes, "types", SeriesChartType.Bar, Theme.Brand);
            ChartStyler.EnableDrillDown(chartTypes, p => NavigateToFields("Type: " + p.Label, p.Filter));

            chartOrigin = ChartStyler.Create(cardOrigin, "No field.");
            chartOrigin.Legends.Add(new Legend("Legend1"));
            chartOrigin.Series.Add(new Series("origin") { ChartArea = "main", Legend = "Legend1" });
            ChartStyler.StyleDoughnut(chartOrigin, new[] { Theme.Brand },
                () => _insights == null ? null : Tuple.Create(_insights.Total.ToString("N0"), "fields"));
            ChartStyler.EnableDrillDown(chartOrigin, p => NavigateToFields("Origin: " + p.Label, p.Filter));

            chartSettings = ChartStyler.Create(cardSettings, "No field.");
            ChartStyler.StyleBars(chartSettings, true);
            ChartStyler.AddSeries(chartSettings, "settings", SeriesChartType.Bar, Theme.ChartPalette[2]);
            ChartStyler.EnableDrillDown(chartSettings, p => NavigateToFields(p.Label, p.Filter));

            chartPublishers = ChartStyler.Create(cardPublishers, "No custom field on this entity.");
            ChartStyler.StyleBars(chartPublishers, true);
            ChartStyler.AddSeries(chartPublishers, "publishers", SeriesChartType.Bar, Theme.ChartPalette[1]);
            ChartStyler.EnableDrillDown(chartPublishers, p => NavigateToFields("Prefix: " + p.Label, p.Filter));

            chartTimeline = ChartStyler.Create(cardTimeline, "No creation date available for these fields.");
            ChartStyler.StyleBars(chartTimeline, false);
            ChartStyler.AddTopLegend(chartTimeline);
            ChartStyler.AddSeries(chartTimeline, "Custom", SeriesChartType.StackedColumn, Theme.Brand);
            ChartStyler.AddSeries(chartTimeline, "Standard", SeriesChartType.StackedColumn, NeutralSeries);
            ChartStyler.EnableDrillDown(chartTimeline, p => NavigateToFields("Created: " + p.Label, p.Filter));
            #endregion

            #region Data usage
            chartFillDistribution = ChartStyler.Create(cardFillDistribution, "No field.");
            ChartStyler.StyleBars(chartFillDistribution, false);
            ChartStyler.AddSeries(chartFillDistribution, "distribution", SeriesChartType.Column, Theme.Brand)["BarLabelStyle"] = "Outside";
            chartFillDistribution.Series[0].IsValueShownAsLabel = true;
            ChartStyler.EnableDrillDown(chartFillDistribution, p => NavigateToFields("Fill rate: " + p.Label, p.Filter));

            chartFillByType = ChartStyler.Create(cardFillByType, "No field.");
            ChartStyler.StyleBars(chartFillByType, true);
            ChartStyler.AddSeries(chartFillByType, "fillByType", SeriesChartType.Bar, Theme.ChartPalette[1]);
            ChartStyler.EnableDrillDown(chartFillByType, p => NavigateToFields("Type: " + p.Label, p.Filter));

            chartLeastUsed = ChartStyler.Create(cardLeastUsed, "No custom field on this entity.");
            ChartStyler.StyleBars(chartLeastUsed, true);
            ChartStyler.AddSeries(chartLeastUsed, "leastUsed", SeriesChartType.Bar, Theme.Warning);
            ChartStyler.EnableDrillDown(chartLeastUsed, p => NavigateToFields("Field: " + p.Label, p.Filter));
            #endregion

            #region Capacity
            chartBudget = ChartStyler.Create(budgetLayout, "No field.");
            budgetLayout.SetCellPosition(chartBudget, new TableLayoutPanelCellPosition(0, 1));
            ChartStyler.StyleBars(chartBudget, true);
            chartBudget.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
            chartBudget.ChartAreas[0].AxisX.LineWidth = 0;
            chartBudget.ChartAreas[0].AxisY.LabelStyle.Enabled = true;
            chartBudget.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
            var budgetLegend = ChartStyler.AddTopLegend(chartBudget);
            budgetLegend.Docking = Docking.Bottom;
            budgetLegend.Alignment = StringAlignment.Near;
            budgetLegend.Font = Theme.Body;
            AddBudgetSeries("Lookups (×3)", Theme.Brand);
            AddBudgetSeries("Option set, money, boolean, status (×2)", Theme.ChartPalette[1]);
            AddBudgetSeries("Other types (×1)", Theme.ChartPalette[2]);
            var planned = AddBudgetSeries("Planned", Theme.BrandSelected);
            planned.BackHatchStyle = ChartHatchStyle.WideUpwardDiagonal;
            planned.BackSecondaryColor = Theme.Brand;
            AddBudgetSeries("Available", Theme.StrokeSubtle);

            chartColumnCost = ChartStyler.Create(cardColumnCost, "No field.");
            ChartStyler.StyleBars(chartColumnCost, true);
            ChartStyler.AddSeries(chartColumnCost, "cost", SeriesChartType.Bar, Theme.Brand);
            ChartStyler.EnableDrillDown(chartColumnCost, p => NavigateToFields("Type: " + p.Label, p.Filter));
            #endregion
        }

        private Series AddBudgetSeries(string name, Color color)
        {
            var series = ChartStyler.AddSeries(chartBudget, name, SeriesChartType.StackedBar, color);
            series["PointWidth"] = "0.55";
            series.BorderColor = Theme.Surface;
            series.BorderWidth = 1;
            return series;
        }
        #endregion

        #region Render
        /// <summary>Fills every page from <see cref="_insights"/>.</summary>
        private void RenderInsights()
        {
            var i = _insights;
            if (i == null)
                return;

            #region KPIs
            kpiFields.Value = i.Total.ToString("N0");
            kpiFields.Caption = "open the list";
            kpiCustom.Value = i.CustomCount.ToString("N0");
            kpiCustom.Caption = Percent(i.CustomCount, i.Total) + " of fields";
            kpiUnmanaged.Value = i.UnmanagedCount.ToString("N0");
            kpiUnmanaged.Caption = (i.Total - i.UnmanagedCount).ToString("N0") + " managed";
            var ratio = (float)i.ColumnsUsed / FieldInsights.ColumnLimit;
            kpiCapacity.Value = Percent(i.ColumnsUsed, FieldInsights.ColumnLimit);
            kpiCapacity.Caption = $"{i.ColumnsAvailable:N0} columns left";
            kpiCapacity.Progress = ratio;
            kpiCapacity.Accent = CapacityColor(ratio);
            #endregion

            #region Overview charts
            ChartStyler.Bind(chartTypes.Series[0], i.TypeCounts, p => p.Value.ToString("N0"));
            ChartStyler.FitValueAxis(chartTypes, i.TypeCounts.Select(p => p.Value).DefaultIfEmpty(0).Max());
            ChartStyler.UpdateEmptyState(chartTypes);
            cardTypes.Subtitle = $"{i.TypeCounts.Count} types · click a bar to list them";

            var originColors = new Dictionary<string, Color>
            {
                { "Standard", NeutralSeries },
                { "Custom · managed", Theme.Brand },
                { "Custom · unmanaged", Theme.ChartPalette[4] },
                { "Standard · unmanaged", Theme.ChartPalette[2] },
            };
            ChartStyler.Bind(chartOrigin.Series[0], i.Origin.Where(p => p.Value > 0), null, p => originColors[p.Label]);
            ChartStyler.SetValueMode(chartOrigin, false);
            cardOrigin.Subtitle = "standard / custom, managed / unmanaged";

            ChartStyler.Bind(chartSettings.Series[0], i.Settings, p => $"{p.Value:0}%");
            ChartStyler.FitValueAxis(chartSettings, 100);
            ChartStyler.UpdateEmptyState(chartSettings);

            ChartStyler.Bind(chartPublishers.Series[0], i.Publishers, p => p.Value.ToString("N0"));
            ChartStyler.FitValueAxis(chartPublishers, i.Publishers.Select(p => p.Value).DefaultIfEmpty(0).Max());
            ChartStyler.UpdateEmptyState(chartPublishers);

            ChartStyler.Bind(chartTimeline.Series[0], i.Timeline);
            ChartStyler.Bind(chartTimeline.Series[1], i.Timeline, secondary: true);
            foreach (var point in chartTimeline.Series[0].Points)
                point.ToolTip = $"{point.AxisLabel}: {point.YValues[0]:N0} custom field(s)\nClick to see the fields created in this period";
            foreach (var point in chartTimeline.Series[1].Points)
                point.ToolTip = $"{point.AxisLabel}: {point.YValues[0]:N0} standard field(s)\nClick to see the fields created in this period";
            chartTimeline.ChartAreas[0].AxisX.Interval = Math.Max(1, (int)Math.Ceiling(i.Timeline.Count / 14.0));
            chartTimeline.ChartAreas[0].AxisY.Maximum = double.NaN;
            ChartStyler.UpdateEmptyState(chartTimeline);
            cardTimeline.Subtitle = i.Timeline.Count == 0 ? string.Empty : $"{i.TimelineGranularity} · click a column to list the fields";
            #endregion

            #region Data usage
            usageEmpty.Visible = !i.HasDataUsage;
            usageLayout.Visible = i.HasDataUsage;
            if (i.HasDataUsage)
            {
                var info = EntityFieldAnalyserManager.entityInfo;
                kpiRecords.Value = info.entityRecordsCount.ToString("N0");
                kpiRecords.Caption = "records read";
                kpiUnused.Value = i.UnusedCount.ToString("N0");
                kpiUnused.Caption = Percent(i.UnusedCount, i.Total) + " of fields";
                kpiUnused.Accent = i.UnusedCount > 0 ? Theme.Danger : Theme.Success;
                kpiLowUsage.Value = i.LowUsageCount.ToString("N0");
                kpiLowUsage.Caption = "filled in < 5% of records";
                kpiLowUsage.Accent = Theme.Warning;
                kpiAverage.Value = (i.AverageFillRate / 100).ToString("0%");
                kpiAverage.Caption = "per field";
                kpiAverage.Progress = (float)(i.AverageFillRate / 100);

                var ramp = new[] { Theme.Danger, Theme.Warning, Color.FromArgb(234, 163, 60), Color.FromArgb(120, 170, 220), Theme.Brand, Theme.BrandPressed };
                var index = 0;
                ChartStyler.Bind(chartFillDistribution.Series[0], i.FillDistribution, p => p.Value.ToString("N0"), p => ramp[index++ % ramp.Length]);
                chartFillDistribution.ChartAreas[0].AxisY.Maximum = Math.Ceiling(i.FillDistribution.Select(p => p.Value).DefaultIfEmpty(1).Max() * 1.2);
                ChartStyler.UpdateEmptyState(chartFillDistribution);

                ChartStyler.Bind(chartFillByType.Series[0], i.FillByType, p => $"{p.Value:0.#}%");
                ChartStyler.FitValueAxis(chartFillByType, 100);
                ChartStyler.UpdateEmptyState(chartFillByType);

                ChartStyler.Bind(chartLeastUsed.Series[0], i.LeastUsedCustom, p => $"{p.Value:0.##}%", p => p.Value <= 0 ? Theme.Danger : Theme.Warning);
                ChartStyler.FitValueAxis(chartLeastUsed, Math.Max(1, i.LeastUsedCustom.Select(p => p.Value).DefaultIfEmpty(0).Max()), 1.3);
                chartLeastUsed.ChartAreas[0].Visible = i.LeastUsedCustom.Count > 0;
                chartLeastUsed.Invalidate();
            }
            #endregion

            #region Capacity
            RenderBudget();
            ChartStyler.Bind(chartColumnCost.Series[0], i.ColumnCostByType, p => p.Value.ToString("N0"));
            ChartStyler.FitValueAxis(chartColumnCost, i.ColumnCostByType.Select(p => p.Value).DefaultIfEmpty(0).Max());
            ChartStyler.UpdateEmptyState(chartColumnCost);
            #endregion

            pivotTabs.SetBadge((int)Page.Fields, i.Total.ToString("N0"));
            pivotTabs.SetBadge((int)Page.Usage, i.HasDataUsage ? i.UnusedCount.ToString("N0") + " unused" : "not run");
            pivotTabs.SetBadge((int)Page.Capacity, Percent(i.ColumnsUsed, FieldInsights.ColumnLimit));
        }

        /// <summary>Column budget stacked bar + calculator result; called on every calculator keystroke.</summary>
        private void RenderBudget()
        {
            var i = _insights;
            if (i == null)
                return;

            var planned = ParseCount(textBox1.Text) * 3 + ParseCount(textBox2.Text) * 2 + ParseCount(textBox3.Text);
            var limit = FieldInsights.ColumnLimit;
            var available = limit - i.ColumnsUsed;
            var fits = available > planned; // same rule as EntityFieldAnalyserManager.CanICreateThisNumberOfFields

            var values = new[]
            {
                i.ColumnsByWeight[ColumnWeight.Lookup],
                i.ColumnsByWeight[ColumnWeight.OptionLike],
                i.ColumnsByWeight[ColumnWeight.Other],
                planned,
                Math.Max(0, limit - i.ColumnsUsed - planned),
            };
            for (var s = 0; s < chartBudget.Series.Count; s++)
            {
                var series = chartBudget.Series[s];
                series.Points.Clear();
                var dp = series.Points[series.Points.AddXY("Columns", values[s])];
                dp.ToolTip = $"{series.Name}: {values[s]:N0} columns";
                series.LegendText = $"{series.Name}  {values[s]:N0}";
                series.IsVisibleInLegend = s != 3 || planned > 0;
            }
            var plannedSeries = chartBudget.Series[3];
            plannedSeries.Color = fits ? Theme.BrandSelected : Theme.DangerSubtle;
            plannedSeries.BackSecondaryColor = fits ? Theme.Brand : Theme.Danger;

            var axis = chartBudget.ChartAreas[0].AxisY;
            axis.Minimum = 0;
            axis.Maximum = Math.Max(limit, i.ColumnsUsed + planned);
            axis.Interval = 128;
            axis.StripLines.Clear();
            if (i.ColumnsUsed + planned > limit)
                axis.StripLines.Add(new StripLine { IntervalOffset = limit, StripWidth = 0, BorderColor = Theme.Danger, BorderWidth = 2, Text = "limit", ForeColor = Theme.Danger, Font = Theme.Caption });

            budgetHeadline.Text = $"{i.ColumnsUsed:N0} of {limit:N0} columns used ({Percent(i.ColumnsUsed, limit)})  ·  {available:N0} available";
            budgetHeadline.ForeColor = CapacityColor((float)i.ColumnsUsed / limit) == Theme.Brand ? Theme.TextPrimary : CapacityColor((float)i.ColumnsUsed / limit);

            if (planned == 0)
            {
                calculatorResult.ForeColor = Theme.TextTertiary;
                calculatorResult.Text = "Enter the number of fields you plan to create.";
            }
            else if (fits)
            {
                calculatorResult.ForeColor = Theme.Success;
                calculatorResult.Text = $"✓  It fits: {planned:N0} columns needed, {available - planned:N0} left afterwards.";
            }
            else
            {
                calculatorResult.ForeColor = Theme.Danger;
                calculatorResult.Text = $"✕  It doesn't fit: {planned:N0} columns needed, only {available:N0} available.";
            }
        }

        private static int ParseCount(string text)
        {
            int value;
            return int.TryParse(text, out value) ? value : 0;
        }

        private static Color CapacityColor(float ratio) => ratio >= 0.9f ? Theme.Danger : ratio >= 0.75f ? Theme.Warning : Theme.Brand;

        private void calculatorInput_TextChanged(object sender, EventArgs e) => RenderBudget();
        #endregion

        #region Header & guide
        private string ConnectionName
        {
            get
            {
                try { return ConnectionDetail?.ConnectionName; }
                catch (Exception) { return null; }
            }
        }

        private void UpdateHeader()
        {
            if (_insights == null)
            {
                contextHeader.Title = "Entity Fields Analyser";
                contextHeader.Subtitle = "Understand the fields of a Dataverse entity: types, origin, data usage and remaining column capacity.";
                contextHeader.SetBadges(ConnectionName);
                contextHeader.Hint = string.Empty;
                return;
            }

            var info = EntityFieldAnalyserManager.entityInfo;
            contextHeader.Title = entitySelectedName;
            contextHeader.Subtitle = $"{entitySelectedSchemaName}  ·  analysed at {DateTime.Now:t}";
            contextHeader.SetBadges(
                ConnectionName,
                _insights.HasDataUsage ? "Metadata + data usage" : "Metadata only",
                _insights.HasDataUsage ? $"{info.entityRecordsCount:N0} records" : null);
            UpdateGuide();
        }

        /// <summary>Getting started steps + header hint when another entity than the analysed one is selected.</summary>
        private void UpdateGuide()
        {
            var loaded = dtEntities != null;
            var selection = EntityGridView.DataSource != null ? EntityFieldAnalyserManager.SelectedEntity(EntityGridView.Rows) : new string[0];
            var picked = selection.Length > 1;

            stepLoad.State = loaded ? StepState.Done : StepState.Current;
            stepPick.State = !loaded ? StepState.Pending : picked ? StepState.Done : StepState.Current;
            stepAnalyse.State = loaded && picked ? StepState.Current : StepState.Pending;

            if (!loaded)
            {
                guideAction.Visible = true;
                guideAction.Text = "Get Entities";
                guideAction.Glyph = Theme.Glyph.List;
                guideHint.Text = "Lists the entities of the connected environment, using the filter selected in the toolbar.";
            }
            else if (!picked)
            {
                guideAction.Visible = false;
                guideHint.Text = "←  Click an entity in the list on the left (or double-click it to analyse it right away).";
            }
            else
            {
                guideAction.Visible = true;
                guideAction.Text = "Analyse " + selection[0];
                guideAction.Glyph = Theme.Glyph.Play;
                guideAction.Width = Math.Max(Theme.Scale(170), TextRenderer.MeasureText(guideAction.Text, guideAction.Font).Width + Theme.Scale(60));
                guideHint.Text = "Mode: " + AnalyseType.SelectedItem + " (change it in the toolbar).";
            }
            guideAction.Invalidate();

            contextHeader.Hint = _insights != null && picked && selection[1] != entitySelectedSchemaName
                ? $"{selection[0]} selected · click Analyse to switch"
                : string.Empty;
        }

        private void guideAction_Click(object sender, EventArgs e)
        {
            if (dtEntities == null)
                getEntitiesButton_Click(sender, e);
            else if (analyseButton.Enabled)
                analyseButton_Click(sender, e);
        }
        #endregion

        #region Navigation
        private void ShowPage(Page page) => pivotTabs.SelectedIndex = (int)page;

        /// <summary>Welcome page (no analysis yet): tabs are visible but disabled, so the user sees what is coming.</summary>
        private void ShowWelcome()
        {
            for (var p = 0; p < pivotTabs.Items.Count; p++)
            {
                pivotTabs.SetEnabled(p, false);
                pivotTabs.SetBadge(p, null);
            }
            pivotTabs.SelectedIndex = -1;
            pivotTabs_SelectedIndexChanged(pivotTabs, EventArgs.Empty);
        }

        private void pivotTabs_SelectedIndexChanged(object sender, EventArgs e)
        {
            var index = pivotTabs.SelectedIndex;
            pagesHost.SuspendLayout();
            pageWelcome.Visible = index < 0;
            pageOverview.Visible = index == (int)Page.Overview;
            pageFields.Visible = index == (int)Page.Fields;
            pageUsage.Visible = index == (int)Page.Usage;
            pageCapacity.Visible = index == (int)Page.Capacity;
            pagesHost.ResumeLayout();
        }

        /// <summary>Opens the Fields tab with a drill-down filter (shown as a dismissible chip).</summary>
        private void NavigateToFields(string label, string filter, int chip = 0)
        {
            if (_insights == null)
                return;

            _suppressFieldFilter = true;
            try
            {
                if (fieldTypeCombobox.Items.Count > 0)
                    fieldTypeCombobox.SelectedIndex = 0;
                searchField.Text = string.Empty;
                fieldsChips.SelectedIndex = chip;
            }
            finally
            {
                _suppressFieldFilter = false;
            }

            _drillFilter = filter;
            fieldsChips.DismissibleText = filter == null ? null : label;
            ApplyFieldFilters();
            ShowPage(Page.Fields);
        }

        private void kpiFields_Click(object sender, EventArgs e) => NavigateToFields(null, null);
        private void kpiCustom_Click(object sender, EventArgs e) => NavigateToFields(null, null, 1);
        private void kpiUnmanaged_Click(object sender, EventArgs e) => NavigateToFields(null, null, 2);
        private void kpiCapacity_Click(object sender, EventArgs e) { if (_insights != null) ShowPage(Page.Capacity); }
        private void kpiUnused_Click(object sender, EventArgs e) => NavigateToFields(null, null, 4);
        private void kpiLowUsage_Click(object sender, EventArgs e) =>
            NavigateToFields("Fill rate: < 5%", "[Percentage Of Use] > 0 AND [Percentage Of Use] < 5");

        private void rerunWithUsage_Click(object sender, EventArgs e)
        {
            AnalyseType.SelectedIndex = 1;
            UpdateGuide();
            if (analyseButton.Enabled)
                analyseButton_Click(sender, e);
        }
        #endregion

        #region Fields filtering
        private void fieldsChips_SelectedIndexChanged(object sender, EventArgs e) => ApplyFieldFilters();

        private void fieldsChips_Dismissed(object sender, EventArgs e)
        {
            _drillFilter = null;
            ApplyFieldFilters();
        }

        /// <summary>Type + search + chip + drill-down filters, combined, applied to the fields grid.</summary>
        private void ApplyFieldFilters()
        {
            if (_suppressFieldFilter || entityFields == null || dtFields == null || fieldTypeCombobox.SelectedItem == null)
                return;

            var selectedType = fieldTypeCombobox.SelectedItem.ToString();
            var fields = selectedType == "ALL"
                ? entityFields.Values.SelectMany(list => list).ToList()
                : entityFields[(Microsoft.Xrm.Sdk.Metadata.AttributeTypeCode)fieldTypeCombobox.SelectedItem];

            dtFields.Clear();
            EntityFieldAnalyserManager.SetFieldDataGridViewContent(fields, dtFields, _analysedType);

            var filters = new List<string>();
            if (!string.IsNullOrWhiteSpace(searchField.Text))
            {
                var value = EscapeLike(searchField.Text.Trim());
                filters.Add($"DisplayName LIKE '*{value}*' OR SchemaName LIKE '*{value}*'");
            }
            var chip = fieldsChips.SelectedIndex;
            if (chip > 0 && chip < ChipFilters.Length && ChipFilters[chip] != null)
                filters.Add(ChipFilters[chip]);
            if (_drillFilter != null)
                filters.Add(_drillFilter);

            var view = dtFields;
            if (filters.Count > 0)
            {
                var rows = dtFields.Select(string.Join(" AND ", filters.Select(f => "(" + f + ")")));
                view = rows.Length > 0 ? rows.CopyToDataTable() : dtFields.Clone();
            }

            EntityFieldAnalyserManager.SetFieldDataGridViewHeaders(view, fieldPropretiesView, _analysedType, displayAllColumns.Checked, selectedType);
            fieldsResultLabel.Text = view.Rows.Count == _insights.Total
                ? $"{_insights.Total:N0} fields"
                : $"{view.Rows.Count:N0} of {_insights.Total:N0} fields";
        }

        /// <summary>Escapes a user value for a DataTable LIKE expression.</summary>
        private static string EscapeLike(string value)
        {
            var sb = new StringBuilder();
            foreach (var c in value)
            {
                if (c == '\'') sb.Append("''");
                else if (c == '[' || c == ']' || c == '*' || c == '%') sb.Append('[').Append(c).Append(']');
                else sb.Append(c);
            }
            return sb.ToString();
        }
        #endregion
    }
}
