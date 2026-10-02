using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace EntityieldsAnalyser.UI
{
    internal static class ChartStyler
    {
        private const string ValueColumn = "value";

        #region Factory
        /// <summary>Creates a chart docked in <paramref name="host"/>, with one transparent area and an empty state message.</summary>
        public static Chart Create(Control host, string emptyMessage)
        {
            var chart = new Chart
            {
                Dock = DockStyle.Fill,
                BackColor = Theme.Surface,
                AntiAliasing = AntiAliasingStyles.All,
                TextAntiAliasingQuality = TextAntiAliasingQuality.High,
                Tag = emptyMessage,
            };
            chart.ChartAreas.Add(new ChartArea("main") { BackColor = Color.Transparent });
            chart.PostPaint += (sender, e) =>
            {
                if (!(e.ChartElement is Chart) || chart.ChartAreas[0].Visible)
                    return;
                var r = new Rectangle(0, 0, chart.Width, chart.Height);
                Theme.DrawGlyph(e.ChartGraphics.Graphics, Theme.Glyph.Info, Theme.TextDisabled, new RectangleF(0, r.Height / 2f - Theme.Scale(30), r.Width, Theme.Scale(24)), Theme.Scale(20));
                TextRenderer.DrawText(e.ChartGraphics.Graphics, chart.Tag as string, Theme.Body, new Rectangle(Theme.Scale(12), r.Height / 2, r.Width - Theme.Scale(24), Theme.Scale(40)),
                    Theme.TextTertiary, TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak);
            };
            host.Controls.Add(chart);
            return chart;
        }

        /// <summary>Minimal axes: no tick marks, hairline grid on the value axis only for columns, values printed on bars.</summary>
        public static void StyleBars(Chart chart, bool horizontal)
        {
            var area = chart.ChartAreas[0];
            var x = area.AxisX;
            x.Interval = 1;
            x.IsLabelAutoFit = false;
            x.LabelStyle.Font = Theme.Body;
            x.LabelStyle.ForeColor = Theme.TextSecondary;
            x.LineColor = Theme.Stroke;
            x.MajorGrid.Enabled = false;
            x.MajorTickMark.Enabled = false;
            x.IsReversed = horizontal; // first item on top

            var y = area.AxisY;
            y.LineWidth = 0;
            y.MajorTickMark.Enabled = false;
            y.LabelStyle.Font = Theme.Caption;
            y.LabelStyle.ForeColor = Theme.TextTertiary;
            y.MajorGrid.LineColor = Theme.StrokeSubtle;
            y.MajorGrid.Enabled = !horizontal;
            y.LabelStyle.Enabled = !horizontal;
        }

        public static Series AddSeries(Chart chart, string name, SeriesChartType type, Color color)
        {
            var series = new Series(name)
            {
                ChartType = type,
                ChartArea = chart.ChartAreas[0].Name,
                Color = color,
                Font = Theme.Caption,
                LabelForeColor = Theme.TextSecondary,
                IsVisibleInLegend = true,
            };
            series["PointWidth"] = "0.62";
            series["BarLabelStyle"] = "Outside";
            chart.Series.Add(series);
            return series;
        }

        public static Legend AddTopLegend(Chart chart)
        {
            var legend = new Legend("main")
            {
                Docking = Docking.Top,
                Alignment = StringAlignment.Far,
                BackColor = Color.Transparent,
                Font = Theme.Caption,
                ForeColor = Theme.TextSecondary,
                IsTextAutoFit = false,
            };
            chart.Legends.Add(legend);
            return legend;
        }

        /// <summary>Fills a series; each DataPoint keeps its <see cref="InsightPoint"/> in Tag for drill-down.</summary>
        public static void Bind(Series series, IEnumerable<InsightPoint> points, Func<InsightPoint, string> label = null, Func<InsightPoint, Color?> color = null, bool secondary = false)
        {
            series.Points.Clear();
            foreach (var p in points)
            {
                var value = secondary ? p.Secondary : p.Value;
                var dp = series.Points[series.Points.AddXY(p.Label, value)];
                dp.Tag = p;
                dp.Label = label != null ? label(p) : string.Empty;
                dp.ToolTip = (p.Tooltip ?? $"{p.Label}: {value:N0}") + (p.Filter != null ? "\nClick to see these fields" : string.Empty);
                var c = color?.Invoke(p);
                if (c.HasValue)
                    dp.Color = c.Value;
            }
        }

        /// <summary>Leaves room after the longest bar for its value label.</summary>
        public static void FitValueAxis(Chart chart, double max, double headroom = 1.22)
        {
            var y = chart.ChartAreas[0].AxisY;
            y.Minimum = 0;
            y.Maximum = max <= 0 ? 1 : Math.Ceiling(max * headroom);
        }

        /// <summary>Hides the plot and shows the chart's empty message when no series has a non zero point.</summary>
        public static void UpdateEmptyState(Chart chart)
        {
            var hasData = false;
            foreach (var s in chart.Series)
                foreach (var p in s.Points)
                    if (p.YValues.Length > 0 && p.YValues[0] > 0) { hasData = true; break; }

            chart.ChartAreas[0].Visible = hasData;
            foreach (var legend in chart.Legends)
                legend.Enabled = hasData;
            chart.Invalidate();
        }

        /// <summary>Hand cursor over drillable points/legend items, and a callback when one is clicked.</summary>
        public static void EnableDrillDown(Chart chart, Action<InsightPoint> onDrill)
        {
            chart.MouseMove += (sender, e) =>
            {
                var point = PointAt(chart, e.Location);
                chart.Cursor = point != null && point.Filter != null ? Cursors.Hand : Cursors.Default;
            };
            chart.MouseClick += (sender, e) =>
            {
                var point = PointAt(chart, e.Location);
                if (point != null && point.Filter != null)
                    onDrill(point);
            };
        }

        private static InsightPoint PointAt(Chart chart, Point location)
        {
            HitTestResult hit;
            try
            {
                hit = chart.HitTest(location.X, location.Y);
            }
            catch (Exception)
            {
                return null;
            }

            if (hit.Series == null || hit.PointIndex < 0 || hit.PointIndex >= hit.Series.Points.Count)
                return null;
            if (hit.ChartElementType != ChartElementType.DataPoint && hit.ChartElementType != ChartElementType.DataPointLabel && hit.ChartElementType != ChartElementType.LegendItem)
                return null;
            return hit.Series.Points[hit.PointIndex].Tag as InsightPoint;
        }
        #endregion

        /// <summary>
        /// Turns an MS Chart into a thin Fluent doughnut with a table legend and a metric drawn in its center.
        /// </summary>
        public static void StyleDoughnut(Chart chart, Color[] palette, Func<Tuple<string, string>> centerText)
        {
            chart.BackColor = Theme.Surface;
            chart.AntiAliasing = AntiAliasingStyles.All;
            chart.TextAntiAliasingQuality = TextAntiAliasingQuality.High;
            chart.Palette = ChartColorPalette.None;
            chart.PaletteCustomColors = palette;

            foreach (var area in chart.ChartAreas)
            {
                area.BackColor = Color.Transparent;
                area.Area3DStyle.Enable3D = false;
            }

            foreach (var legend in chart.Legends)
            {
                legend.Docking = Docking.Right;
                legend.Alignment = StringAlignment.Center;
                legend.BackColor = Color.Transparent;
                legend.Font = Theme.Body;
                legend.ForeColor = Theme.TextSecondary;
                legend.IsTextAutoFit = false;
                legend.MaximumAutoSize = 70;
                legend.TableStyle = LegendTableStyle.Wide;
                legend.ItemColumnSpacing = 30;
                legend.CellColumns.Clear();
                legend.CellColumns.Add(new LegendCellColumn { ColumnType = LegendCellColumnType.SeriesSymbol, Margins = new Margins(0, 0, 15, 15) });
                legend.CellColumns.Add(new LegendCellColumn("", LegendCellColumnType.Text, "#VALX", ContentAlignment.MiddleLeft) { ForeColor = Theme.TextSecondary, Font = Theme.Body });
                legend.CellColumns.Add(new LegendCellColumn("", LegendCellColumnType.Text, "#VAL", ContentAlignment.MiddleRight) { Name = ValueColumn, ForeColor = Theme.TextPrimary, Font = Theme.BodyStrong, MinimumWidth = 250 });
            }

            foreach (var series in chart.Series)
            {
                series.ChartType = SeriesChartType.Doughnut;
                series["DoughnutRadius"] = "28";
                series["PieStartAngle"] = "270";
                series["PieLabelStyle"] = "Disabled";
                series.BorderColor = Theme.Surface;
                series.BorderWidth = 2;
                series.IsValueShownAsLabel = false;
                series.Font = Theme.Body;
            }

            chart.PostPaint += (sender, e) =>
            {
                var area = e.ChartElement as ChartArea;
                if (area == null || chart.Series.Count == 0 || chart.Series[0].Points.Count == 0)
                    return;

                var text = centerText?.Invoke();
                if (text == null)
                    return;

                var abs = e.ChartGraphics.GetAbsoluteRectangle(area.Position.ToRectangleF());
                var inner = area.InnerPlotPosition;
                var plot = inner.Width > 0 && inner.Height > 0
                    ? new RectangleF(abs.X + abs.Width * inner.X / 100f, abs.Y + abs.Height * inner.Y / 100f, abs.Width * inner.Width / 100f, abs.Height * inner.Height / 100f)
                    : abs;

                var g = e.ChartGraphics.Graphics;
                var oldHint = g.TextRenderingHint;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                using (var valueBrush = new SolidBrush(Theme.TextPrimary))
                using (var captionBrush = new SolidBrush(Theme.TextTertiary))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    var valueSize = g.MeasureString(text.Item1, Theme.MetricSmall);
                    var captionSize = g.MeasureString(text.Item2, Theme.Caption);
                    var cx = plot.X + plot.Width / 2f;
                    var top = plot.Y + (plot.Height - valueSize.Height - captionSize.Height) / 2f;
                    g.DrawString(text.Item1, Theme.MetricSmall, valueBrush, new RectangleF(cx - valueSize.Width, top, valueSize.Width * 2, valueSize.Height), format);
                    g.DrawString(text.Item2, Theme.Caption, captionBrush, new RectangleF(cx - captionSize.Width, top + valueSize.Height - 2, captionSize.Width * 2, captionSize.Height), format);
                }
                g.TextRenderingHint = oldHint;
            };
        }

        /// <summary>Switches legend values between raw counts and percentages, and keeps slice labels off.</summary>
        public static void SetValueMode(Chart chart, bool percentage)
        {
            foreach (var legend in chart.Legends)
            {
                if (legend.CellColumns.IndexOf(ValueColumn) >= 0)
                    legend.CellColumns[ValueColumn].Text = percentage ? "#PERCENT{P1}" : "#VAL{N0}";
            }

            foreach (var series in chart.Series)
            {
                series.IsValueShownAsLabel = false;
                series["PieLabelStyle"] = "Disabled";
            }

            chart.Invalidate();
        }
    }

    internal static class GridStyler
    {
        public const string PercentageColumn = "Percentage Of Use";

        /// <summary>
        /// Fluent table look: roomy rows, horizontal hairlines, quiet header, check glyphs for booleans,
        /// in-cell bars for the percentage of use and an empty state message.
        /// </summary>
        public static void Apply(DataGridView grid, Func<string> emptyText, Func<DataGridViewCellPaintingEventArgs, bool> customPainter = null)
        {
            grid.BackgroundColor = Theme.Surface;
            grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.GridColor = Theme.StrokeSubtle;
            grid.RowHeadersVisible = false;
            grid.AllowUserToResizeRows = false;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = Theme.Scale(36);
            grid.RowTemplate.Height = Theme.Scale(34);

            var header = grid.ColumnHeadersDefaultCellStyle;
            header.BackColor = Theme.Surface;
            header.ForeColor = Theme.TextSecondary;
            header.SelectionBackColor = Theme.Surface;
            header.SelectionForeColor = Theme.TextSecondary;
            header.Font = Theme.BodyStrong;
            header.Padding = new Padding(Theme.Scale(6), 0, Theme.Scale(6), 0);
            header.WrapMode = DataGridViewTriState.False;

            var cell = grid.DefaultCellStyle;
            cell.BackColor = Theme.Surface;
            cell.ForeColor = Theme.TextPrimary;
            cell.SelectionBackColor = Theme.BrandSubtle;
            cell.SelectionForeColor = Theme.TextPrimary;
            cell.Font = Theme.Body;
            cell.Padding = new Padding(Theme.Scale(6), 0, Theme.Scale(6), 0);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Theme.Surface;

            grid.CellPainting += (sender, e) =>
            {
                if (e.ColumnIndex < 0)
                    return;

                if (e.RowIndex == -1)
                {
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);
                    using (var pen = new Pen(Theme.Stroke))
                        e.Graphics.DrawLine(pen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                    e.Handled = true;
                    return;
                }

                if (customPainter != null && customPainter(e))
                    return;

                var column = grid.Columns[e.ColumnIndex];
                if (column.DataPropertyName == PercentageColumn && e.Value is double)
                {
                    PaintPercentage(e, (double)e.Value);
                }
                else if (column is DataGridViewCheckBoxColumn)
                {
                    PaintBoolean(e);
                }
            };

            grid.Paint += (sender, e) =>
            {
                var message = emptyText?.Invoke();
                if (grid.Rows.Count > 0 || string.IsNullOrEmpty(message))
                    return;

                var top = grid.ColumnHeadersVisible && grid.Columns.Count > 0 ? grid.ColumnHeadersHeight : 0;
                var area = new Rectangle(0, top, grid.Width, grid.Height - top);
                var iconRect = new RectangleF(0, area.Y + area.Height / 2f - Theme.Scale(34), area.Width, Theme.Scale(28));
                Theme.DrawGlyph(e.Graphics, Theme.Glyph.Info, Theme.TextDisabled, iconRect, Theme.Scale(22));
                TextRenderer.DrawText(e.Graphics, message, Theme.Body,
                    new Rectangle(Theme.Scale(16), area.Y + area.Height / 2 - Theme.Scale(2), area.Width - Theme.Scale(32), Theme.Scale(40)),
                    Theme.TextTertiary, TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
            };
            grid.Resize += (sender, e) => { if (grid.Rows.Count == 0) grid.Invalidate(); };
        }

        private static bool IsSelected(DataGridViewCellPaintingEventArgs e) => (e.State & DataGridViewElementStates.Selected) != 0;

        private static void PaintBoolean(DataGridViewCellPaintingEventArgs e)
        {
            e.PaintBackground(e.CellBounds, IsSelected(e));
            if (e.Value is bool && (bool)e.Value)
            {
                Theme.DrawGlyph(e.Graphics, Theme.Glyph.CheckMark, Theme.Brand, e.CellBounds, Theme.Scale(12));
            }
            else
            {
                var w = Theme.Scale(8);
                using (var pen = new Pen(Theme.TextDisabled))
                    e.Graphics.DrawLine(pen, e.CellBounds.X + (e.CellBounds.Width - w) / 2, e.CellBounds.Y + e.CellBounds.Height / 2, e.CellBounds.X + (e.CellBounds.Width + w) / 2, e.CellBounds.Y + e.CellBounds.Height / 2);
            }
            e.Handled = true;
        }

        private static void PaintPercentage(DataGridViewCellPaintingEventArgs e, double value)
        {
            e.PaintBackground(e.CellBounds, IsSelected(e));
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var bounds = e.CellBounds;
            var pad = Theme.Scale(10);
            var text = value.ToString("0.##") + " %";
            var textSize = TextRenderer.MeasureText(g, text, Theme.Body, Size.Empty, TextFormatFlags.NoPadding);
            var barWidth = Math.Max(Theme.Scale(30), Math.Min(Theme.Scale(96), bounds.Width - textSize.Width - pad * 3));
            var bar = new RectangleF(bounds.X + pad, bounds.Y + (bounds.Height - Theme.Scale(6)) / 2f, barWidth, Theme.Scale(6));

            Theme.FillRounded(g, Theme.StrokeSubtle, bar, bar.Height / 2);
            var fillWidth = (float)(bar.Width * Math.Min(100, Math.Max(0, value)) / 100.0);
            if (fillWidth >= 1)
                Theme.FillRounded(g, value < 5 ? Theme.Warning : Theme.Brand, new RectangleF(bar.X, bar.Y, Math.Max(bar.Height, fillWidth), bar.Height), bar.Height / 2);

            TextRenderer.DrawText(g, text, Theme.Body,
                new Rectangle((int)bar.Right + pad, bounds.Y, bounds.Right - (int)bar.Right - pad, bounds.Height),
                value <= 0 ? Theme.TextTertiary : Theme.TextPrimary, Draw.LeftMiddle);
            g.SmoothingMode = SmoothingMode.Default;
            e.Handled = true;
        }
    }
}
