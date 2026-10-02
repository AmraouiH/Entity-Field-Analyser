using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace EntityieldsAnalyser.UI
{
    #region PivotTabs
    public class PivotItem
    {
        public PivotItem(string text) { Text = text; }
        public string Text { get; set; }
        public string Badge { get; set; }
        public bool Enabled { get; set; } = true;
        internal Rectangle Bounds;
    }

    /// <summary>Fluent "pivot": text tabs with an accent underline and an optional count badge.</summary>
    public class PivotTabs : Control
    {
        private readonly List<PivotItem> _items = new List<PivotItem>();
        private int _selected = -1;
        private int _hover = -1;

        public PivotTabs()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
        }

        public event EventHandler SelectedIndexChanged;

        public IList<PivotItem> Items => _items;

        public int SelectedIndex
        {
            get { return _selected; }
            set
            {
                if (value == _selected || value < -1 || value >= _items.Count)
                    return;
                _selected = value;
                Invalidate();
                SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public void AddRange(params string[] texts)
        {
            foreach (var t in texts)
                _items.Add(new PivotItem(t));
            Invalidate();
        }

        public void SetBadge(int index, string badge) { _items[index].Badge = badge; Invalidate(); }
        public void SetEnabled(int index, bool enabled) { _items[index].Enabled = enabled; Invalidate(); }

        private void LayoutItems(Graphics g)
        {
            var x = 0;
            foreach (var item in _items)
            {
                var textWidth = TextRenderer.MeasureText(g, item.Text, Theme.Subtitle, Size.Empty, TextFormatFlags.NoPadding).Width;
                var badgeWidth = string.IsNullOrEmpty(item.Badge) ? 0 : TextRenderer.MeasureText(g, item.Badge, Theme.Caption, Size.Empty, TextFormatFlags.NoPadding).Width + Theme.Scale(14);
                var width = textWidth + (badgeWidth > 0 ? badgeWidth + Theme.Scale(6) : 0) + Theme.Scale(24);
                item.Bounds = new Rectangle(x, 0, width, Height);
                x += width;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            LayoutItems(g);

            using (var pen = new Pen(Theme.Stroke))
                g.DrawLine(pen, 0, Height - 1, Width, Height - 1);

            for (var i = 0; i < _items.Count; i++)
            {
                var item = _items[i];
                var selected = i == _selected;
                var color = !item.Enabled ? Theme.TextDisabled : selected ? Theme.TextPrimary : i == _hover ? Theme.TextPrimary : Theme.TextSecondary;
                var font = selected ? Theme.Subtitle : new Font(Theme.FontFamily, 10.5F);

                var textWidth = TextRenderer.MeasureText(g, item.Text, Theme.Subtitle, Size.Empty, TextFormatFlags.NoPadding).Width;
                var textRect = new Rectangle(item.Bounds.X + Theme.Scale(12), 0, textWidth + 2, Height - Theme.Scale(4));
                TextRenderer.DrawText(g, item.Text, font, textRect, color, Draw.LeftMiddle);

                var contentRight = textRect.Right;
                if (!string.IsNullOrEmpty(item.Badge))
                {
                    var badgeSize = TextRenderer.MeasureText(g, item.Badge, Theme.Caption, Size.Empty, TextFormatFlags.NoPadding);
                    var badge = new Rectangle(textRect.Right + Theme.Scale(6), (Height - Theme.Scale(4) - Theme.Scale(20)) / 2, badgeSize.Width + Theme.Scale(14), Theme.Scale(20));
                    Theme.FillRounded(g, selected ? Theme.BrandSubtle : Theme.StrokeSubtle, badge, badge.Height / 2f);
                    TextRenderer.DrawText(g, item.Badge, Theme.Caption, badge, selected ? Theme.Brand : Theme.TextSecondary, Draw.Centered);
                    contentRight = badge.Right;
                }

                if (selected || (i == _hover && item.Enabled))
                {
                    var line = new RectangleF(item.Bounds.X + Theme.Scale(10), Height - Theme.Scale(4), contentRight - item.Bounds.X - Theme.Scale(8), Theme.Scale(3));
                    Theme.FillRounded(g, selected ? Theme.Brand : Theme.StrokeStrong, line, line.Height / 2f);
                }
            }

            if (Focused && ShowFocusCues && _selected >= 0)
            {
                var r = _items[_selected].Bounds;
                Theme.DrawRounded(g, Theme.TextPrimary, new RectangleF(r.X + 2, 3, r.Width - 4, Height - 10), Theme.Scale(4));
            }
        }

        private int HitTest(Point p)
        {
            for (var i = 0; i < _items.Count; i++)
                if (_items[i].Bounds.Contains(p))
                    return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var hit = HitTest(e.Location);
            Cursor = hit >= 0 && _items[hit].Enabled ? Cursors.Hand : Cursors.Default;
            if (hit != _hover) { _hover = hit; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); _hover = -1; Invalidate(); }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            var hit = HitTest(e.Location);
            if (hit >= 0 && _items[hit].Enabled)
            {
                Focus();
                SelectedIndex = hit;
            }
        }

        protected override bool IsInputKey(Keys keyData) => keyData == Keys.Left || keyData == Keys.Right || base.IsInputKey(keyData);

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            var step = e.KeyCode == Keys.Right ? 1 : e.KeyCode == Keys.Left ? -1 : 0;
            if (step == 0 || _items.Count == 0)
                return;
            var i = _selected;
            do { i = (i + step + _items.Count) % _items.Count; } while (!_items[i].Enabled && i != _selected);
            SelectedIndex = i;
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
    }
    #endregion

    #region ChipBar
    /// <summary>
    /// Single choice filter chips, plus an optional dismissible chip that shows an active drill-down filter
    /// (so the user always knows why the list is filtered, and can clear it in one click).
    /// </summary>
    public class ChipBar : Control
    {
        private readonly List<string> _chips = new List<string>();
        private readonly List<Rectangle> _bounds = new List<Rectangle>();
        private Rectangle _dismissBounds;
        private string _dismissText;
        private int _selected;
        private int _hover = -2;

        public ChipBar()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        public event EventHandler SelectedIndexChanged;
        public event EventHandler Dismissed;

        public int SelectedIndex
        {
            get { return _selected; }
            set
            {
                if (value == _selected) return;
                _selected = value;
                Invalidate();
                SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        /// <summary>Text of the drill-down chip, null to hide it.</summary>
        public string DismissibleText
        {
            get { return _dismissText; }
            set { _dismissText = value; Invalidate(); }
        }

        public void SetChips(params string[] chips)
        {
            _chips.Clear();
            _chips.AddRange(chips);
            _selected = 0;
            Invalidate();
        }

        private int ChipHeight => Theme.Scale(28);

        private void LayoutChips(Graphics g)
        {
            _bounds.Clear();
            var x = 0;
            var y = (Height - ChipHeight) / 2;
            foreach (var chip in _chips)
            {
                var w = TextRenderer.MeasureText(g, chip, Theme.Body, Size.Empty, TextFormatFlags.NoPadding).Width + Theme.Scale(24);
                _bounds.Add(new Rectangle(x, y, w, ChipHeight));
                x += w + Theme.Scale(6);
            }

            if (!string.IsNullOrEmpty(_dismissText))
            {
                x += Theme.Scale(10);
                var w = TextRenderer.MeasureText(g, _dismissText, Theme.BodyStrong, Size.Empty, TextFormatFlags.NoPadding).Width + Theme.Scale(60);
                _dismissBounds = new Rectangle(x, y, Math.Min(w, Math.Max(0, Width - x - 1)), ChipHeight);
            }
            else
            {
                _dismissBounds = Rectangle.Empty;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            LayoutChips(g);

            for (var i = 0; i < _chips.Count; i++)
            {
                var r = new RectangleF(_bounds[i].X + 0.5f, _bounds[i].Y + 0.5f, _bounds[i].Width - 1, _bounds[i].Height - 1);
                var selected = i == _selected;
                var fill = selected ? Theme.BrandSubtle : i == _hover ? Theme.SurfaceHover : Theme.Surface;
                Theme.FillRounded(g, Enabled ? fill : Theme.SurfaceDisabled, r, r.Height / 2);
                Theme.DrawRounded(g, !Enabled ? Theme.Stroke : selected ? Theme.Brand : Theme.StrokeStrong, r, r.Height / 2);
                TextRenderer.DrawText(g, _chips[i], selected ? Theme.BodyStrong : Theme.Body, _bounds[i], !Enabled ? Theme.TextDisabled : selected ? Theme.Brand : Theme.TextPrimary, Draw.Centered);
            }

            if (!_dismissBounds.IsEmpty)
            {
                var r = new RectangleF(_dismissBounds.X + 0.5f, _dismissBounds.Y + 0.5f, _dismissBounds.Width - 1, _dismissBounds.Height - 1);
                Theme.FillRounded(g, _hover == -1 ? Theme.BrandHover : Theme.Brand, r, r.Height / 2);
                Theme.DrawGlyph(g, Theme.Glyph.Filter, Theme.TextOnBrand, new RectangleF(r.X + Theme.Scale(8), r.Y, Theme.Scale(14), r.Height), Theme.Scale(11));
                var textRect = new Rectangle(_dismissBounds.X + Theme.Scale(26), _dismissBounds.Y, _dismissBounds.Width - Theme.Scale(56), _dismissBounds.Height);
                TextRenderer.DrawText(g, _dismissText, Theme.BodyStrong, textRect, Theme.TextOnBrand, Draw.LeftMiddle);
                Theme.DrawGlyph(g, Theme.Glyph.Close, Theme.TextOnBrand, new RectangleF(r.Right - Theme.Scale(22), r.Y, Theme.Scale(14), r.Height), Theme.Scale(10));
            }
        }

        private int HitTest(Point p)
        {
            if (_dismissBounds.Contains(p)) return -1;
            for (var i = 0; i < _bounds.Count; i++)
                if (_bounds[i].Contains(p)) return i;
            return -2;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var hit = HitTest(e.Location);
            Cursor = hit > -2 ? Cursors.Hand : Cursors.Default;
            if (hit != _hover) { _hover = hit; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); _hover = -2; Invalidate(); }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (!Enabled) return;
            var hit = HitTest(e.Location);
            if (hit == -1)
            {
                _dismissText = null;
                Invalidate();
                Dismissed?.Invoke(this, EventArgs.Empty);
            }
            else if (hit >= 0)
            {
                SelectedIndex = hit;
            }
        }
    }
    #endregion

    #region ContextHeader
    /// <summary>Page header: what is being looked at (entity, mode, records), plus an optional call to action hint.</summary>
    public class ContextHeader : Control
    {
        private string _title = string.Empty;
        private string _subtitle = string.Empty;
        private string _hint = string.Empty;
        private readonly List<string> _badges = new List<string>();

        public ContextHeader()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        public string Title { get { return _title; } set { _title = value ?? string.Empty; Invalidate(); } }
        public string Subtitle { get { return _subtitle; } set { _subtitle = value ?? string.Empty; Invalidate(); } }
        public string Hint { get { return _hint; } set { _hint = value ?? string.Empty; Invalidate(); } }

        public void SetBadges(params string[] badges)
        {
            _badges.Clear();
            foreach (var b in badges)
                if (!string.IsNullOrEmpty(b)) _badges.Add(b);
            Invalidate();
        }

        private static readonly Font TitleFont = new Font("Segoe UI Semibold", 15F);

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var icon = Theme.Scale(40);
            var iconRect = new Rectangle(Theme.Scale(6), (Height - icon) / 2, icon, icon);
            Theme.FillRounded(g, Theme.Brand, iconRect, Theme.Scale(8));
            Theme.DrawGlyph(g, Theme.Glyph.Table, Theme.TextOnBrand, iconRect, Theme.Scale(18));

            var x = iconRect.Right + Theme.Scale(14);
            var titleSize = TextRenderer.MeasureText(g, _title, TitleFont, Size.Empty, TextFormatFlags.NoPadding);
            var titleRect = new Rectangle(x, Theme.Scale(6), titleSize.Width + 2, Height / 2 + Theme.Scale(2));
            TextRenderer.DrawText(g, _title, TitleFont, titleRect, Theme.TextPrimary, TextFormatFlags.Left | TextFormatFlags.Bottom | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);

            var bx = titleRect.Right + Theme.Scale(12);
            foreach (var badge in _badges)
            {
                var size = TextRenderer.MeasureText(g, badge, Theme.Caption, Size.Empty, TextFormatFlags.NoPadding);
                var r = new Rectangle(bx, titleRect.Bottom - Theme.Scale(22), size.Width + Theme.Scale(16), Theme.Scale(20));
                Theme.FillRounded(g, Theme.Surface, r, r.Height / 2f);
                Theme.DrawRounded(g, Theme.Stroke, r, r.Height / 2f);
                TextRenderer.DrawText(g, badge, Theme.Caption, r, Theme.TextSecondary, Draw.Centered);
                bx = r.Right + Theme.Scale(6);
            }

            var subRect = new Rectangle(x, Height / 2 + Theme.Scale(10), Width - x - Theme.Scale(8), Theme.Scale(20));
            TextRenderer.DrawText(g, _subtitle, Theme.Body, subRect, Theme.TextSecondary, Draw.LeftMiddle);

            if (!string.IsNullOrEmpty(_hint))
            {
                var size = TextRenderer.MeasureText(g, _hint, Theme.BodyStrong, Size.Empty, TextFormatFlags.NoPadding);
                var r = new Rectangle(Width - size.Width - Theme.Scale(44), (Height - Theme.Scale(30)) / 2, size.Width + Theme.Scale(36), Theme.Scale(30));
                if (r.X > bx + Theme.Scale(20))
                {
                    Theme.FillRounded(g, Theme.WarningSubtle, r, Theme.Scale(6));
                    Theme.DrawRounded(g, Theme.Blend(Theme.WarningSubtle, Theme.Warning, 0.35f), r, Theme.Scale(6));
                    Theme.DrawGlyph(g, Theme.Glyph.Info, Theme.Warning, new RectangleF(r.X + Theme.Scale(8), r.Y, Theme.Scale(16), r.Height), Theme.Scale(13));
                    TextRenderer.DrawText(g, _hint, Theme.BodyStrong, new Rectangle(r.X + Theme.Scale(28), r.Y, r.Width - Theme.Scale(30), r.Height), Theme.Warning, Draw.LeftMiddle);
                }
            }
        }
    }
    #endregion

    #region StepCard
    public enum StepState { Pending, Current, Done }

    /// <summary>One step of the getting started guide.</summary>
    public class StepCard : Control
    {
        private StepState _state;
        private string _description = string.Empty;

        public StepCard()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Margin = new Padding(6);
        }

        public int Number { get; set; }
        public string Description { get { return _description; } set { _description = value ?? string.Empty; Invalidate(); } }
        public StepState State { get { return _state; } set { _state = value; Invalidate(); } }

        protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var current = _state == StepState.Current;
            var bounds = new RectangleF(1, 1, Width - 2.5f, Height - 2.5f);
            Theme.FillRounded(g, Theme.Surface, bounds, Theme.Scale(8));
            Theme.DrawRounded(g, current ? Theme.Brand : Theme.Stroke, bounds, Theme.Scale(8), current ? 2f : 1f);

            var pad = Theme.Scale(18);
            var circle = Theme.Scale(32);
            var c = new RectangleF(pad, pad, circle, circle);
            var circleColor = _state == StepState.Done ? Theme.Success : current ? Theme.Brand : Theme.StrokeSubtle;
            using (var brush = new SolidBrush(circleColor))
                g.FillEllipse(brush, c);
            if (_state == StepState.Done)
                Theme.DrawGlyph(g, Theme.Glyph.CheckMark, Theme.TextOnBrand, c, Theme.Scale(14));
            else
                TextRenderer.DrawText(g, Number.ToString(), Theme.Subtitle, Rectangle.Round(c), current ? Theme.TextOnBrand : Theme.TextSecondary, Draw.Centered);

            var status = _state == StepState.Done ? "Done" : current ? "Next step" : "Later";
            var statusColor = _state == StepState.Done ? Theme.Success : current ? Theme.Brand : Theme.TextTertiary;
            var statusSize = TextRenderer.MeasureText(g, status, Theme.BodyStrong, Size.Empty, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, status, Theme.BodyStrong, new Rectangle(Width - pad - statusSize.Width - 2, pad, statusSize.Width + 2, circle), statusColor, Draw.LeftMiddle);

            TextRenderer.DrawText(g, Text, Theme.Subtitle, new Rectangle(pad, pad + circle + Theme.Scale(14), Width - pad * 2, Theme.Scale(24)),
                _state == StepState.Pending ? Theme.TextSecondary : Theme.TextPrimary, Draw.LeftMiddle);
            TextRenderer.DrawText(g, _description, Theme.Body, new Rectangle(pad, pad + circle + Theme.Scale(42), Width - pad * 2, Height - pad - circle - Theme.Scale(46)),
                Theme.TextSecondary, TextFormatFlags.Left | TextFormatFlags.Top | TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
        }
    }
    #endregion
}
