using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace EntityieldsAnalyser.UI
{
    internal static class Draw
    {
        public const TextFormatFlags LeftMiddle = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
        public const TextFormatFlags Centered = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;

        public static Color ParentBack(Control control) => control.Parent != null ? control.Parent.BackColor : Theme.Canvas;
    }

    #region CardBox
    /// <summary>
    /// A GroupBox rendered as a Fluent card: white rounded surface, title and optional subtitle.
    /// Keeps the GroupBox API so existing code (Text, Visible, Controls) keeps working.
    /// </summary>
    public class CardBox : GroupBox
    {
        private string _subtitle = string.Empty;

        public CardBox()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Theme.Surface;
            ForeColor = Theme.TextPrimary;
            Padding = new Padding(14, 0, 14, 12);
            Margin = new Padding(6);
        }

        [DefaultValue("")]
        public string Subtitle
        {
            get { return _subtitle; }
            set { _subtitle = value ?? string.Empty; Invalidate(); }
        }

        private int HeaderHeight => Theme.Scale(44);

        public override Rectangle DisplayRectangle
        {
            get
            {
                return new Rectangle(
                    Padding.Left,
                    HeaderHeight,
                    Math.Max(0, Width - Padding.Horizontal),
                    Math.Max(0, Height - HeaderHeight - Padding.Bottom));
            }
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var bounds = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            Theme.FillRounded(g, BackColor, bounds, Theme.Scale(8));
            Theme.DrawRounded(g, Theme.Stroke, bounds, Theme.Scale(8));

            var header = new Rectangle(Padding.Left, Theme.Scale(6), Width - Padding.Horizontal, HeaderHeight - Theme.Scale(6));
            TextRenderer.DrawText(g, Text, Theme.Subtitle, header, Theme.TextPrimary, Draw.LeftMiddle);

            if (!string.IsNullOrEmpty(_subtitle))
            {
                var titleWidth = TextRenderer.MeasureText(g, Text, Theme.Subtitle, header.Size, Draw.LeftMiddle).Width;
                var sub = new Rectangle(header.X + titleWidth + Theme.Scale(10), header.Y + Theme.Scale(1), Math.Max(0, header.Width - titleWidth - Theme.Scale(10)), header.Height);
                TextRenderer.DrawText(g, _subtitle, Theme.Body, sub, Theme.TextTertiary, Draw.LeftMiddle);
            }
        }
    }
    #endregion

    #region KpiTile
    /// <summary>Metric tile: icon, caption, big value, optional progress bar.</summary>
    public class KpiTile : Control
    {
        private string _title = string.Empty;
        private string _value = "—";
        private string _caption = string.Empty;
        private char _glyph = Theme.Glyph.Info;
        private Color _accent = Theme.Brand;
        private float _progress = -1f;

        public KpiTile()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Margin = new Padding(6);
        }

        public string Title { get { return _title; } set { _title = value; Invalidate(); } }
        public string Value { get { return _value; } set { _value = value; Invalidate(); } }
        public string Caption { get { return _caption; } set { _caption = value; Invalidate(); } }
        public char Glyph { get { return _glyph; } set { _glyph = value; Invalidate(); } }
        public Color Accent { get { return _accent; } set { _accent = value; Invalidate(); } }

        /// <summary>0..1 draws a progress bar, negative hides it.</summary>
        public float Progress { get { return _progress; } set { _progress = value; Invalidate(); } }

        private bool _hover;
        private bool _clickable;

        /// <summary>Shows a hover state and a chevron: the tile navigates somewhere when clicked.</summary>
        public bool Clickable
        {
            get { return _clickable; }
            set { _clickable = value; Cursor = value ? Cursors.Hand : Cursors.Default; Invalidate(); }
        }

        protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); _hover = true; Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); _hover = false; Invalidate(); }

        public void Reset()
        {
            _value = "—";
            _caption = string.Empty;
            _progress = _progress >= 0 ? 0 : -1;
            _accent = Theme.Brand;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var bounds = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            var hot = _clickable && _hover && _value != "—";
            Theme.FillRounded(g, Theme.Surface, bounds, Theme.Scale(8));
            Theme.DrawRounded(g, hot ? Theme.Brand : Theme.Stroke, bounds, Theme.Scale(8));
            if (_clickable && _value != "—")
                Theme.DrawGlyph(g, Theme.Glyph.ChevronRight, hot ? Theme.Brand : Theme.TextDisabled, new RectangleF(Width - Theme.Scale(28), Theme.Scale(8), Theme.Scale(20), Theme.Scale(20)), Theme.Scale(11));

            var pad = Theme.Scale(16);
            var iconSize = Theme.Scale(40);
            var iconRect = new Rectangle(pad, (Height - iconSize) / 2, iconSize, iconSize);
            Theme.FillRounded(g, Theme.Blend(Theme.Surface, _accent, 0.12f), iconRect, Theme.Scale(8));
            Theme.DrawGlyph(g, _glyph, _accent, iconRect, Theme.Scale(18));

            var x = iconRect.Right + Theme.Scale(14);
            var width = Math.Max(0, Width - x - pad);

            TextRenderer.DrawText(g, _title, Theme.Body, new Rectangle(x, Theme.Scale(10), width, Theme.Scale(20)), Theme.TextSecondary, Draw.LeftMiddle);

            var valueSize = TextRenderer.MeasureText(g, _value, Theme.Metric, Size.Empty, TextFormatFlags.NoPadding);
            var valueRect = new Rectangle(x, Theme.Scale(28), Math.Min(width, valueSize.Width), valueSize.Height);
            TextRenderer.DrawText(g, _value, Theme.Metric, valueRect, Theme.TextPrimary, Draw.LeftMiddle);

            if (!string.IsNullOrEmpty(_caption))
            {
                var captionRect = new Rectangle(valueRect.Right + Theme.Scale(8), valueRect.Y + Theme.Scale(4), Math.Max(0, width - valueRect.Width - Theme.Scale(8)), valueRect.Height);
                TextRenderer.DrawText(g, _caption, Theme.Body, captionRect, Theme.TextTertiary, Draw.LeftMiddle);
            }

            if (_progress >= 0)
            {
                var bar = new RectangleF(x, Height - Theme.Scale(16), width, Theme.Scale(4));
                Theme.FillRounded(g, Theme.StrokeSubtle, bar, bar.Height / 2);
                var fill = new RectangleF(bar.X, bar.Y, bar.Width * Math.Min(1f, _progress), bar.Height);
                if (fill.Width >= 1)
                    Theme.FillRounded(g, _accent, fill, bar.Height / 2);
            }
        }
    }
    #endregion

    #region ToggleSwitch
    /// <summary>A CheckBox drawn as a Fluent toggle switch.</summary>
    public class ToggleSwitch : CheckBox
    {
        public ToggleSwitch()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Cursor = Cursors.Hand;
            AutoSize = true;
        }

        private int TrackWidth => Theme.Scale(40);
        private int TrackHeight => Theme.Scale(20);

        public override Size GetPreferredSize(Size proposedSize)
        {
            var text = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPadding);
            return new Size(TrackWidth + Theme.Scale(10) + text.Width + 2, Math.Max(TrackHeight + Theme.Scale(4), text.Height + Theme.Scale(4)));
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(BackColor);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var track = new RectangleF(1, (Height - TrackHeight) / 2f, TrackWidth, TrackHeight);
            if (Checked)
            {
                Theme.FillRounded(g, Enabled ? Theme.Brand : Theme.TextDisabled, track, TrackHeight / 2f);
                var knob = Theme.Scale(12);
                using (var brush = new SolidBrush(Theme.Surface))
                    g.FillEllipse(brush, track.Right - knob - Theme.Scale(4), track.Y + (TrackHeight - knob) / 2f, knob, knob);
            }
            else
            {
                Theme.FillRounded(g, Theme.Surface, track, TrackHeight / 2f);
                Theme.DrawRounded(g, Enabled ? Theme.TextSecondary : Theme.TextDisabled, track, TrackHeight / 2f);
                var knob = Theme.Scale(12);
                using (var brush = new SolidBrush(Enabled ? Theme.TextSecondary : Theme.TextDisabled))
                    g.FillEllipse(brush, track.X + Theme.Scale(4), track.Y + (TrackHeight - knob) / 2f, knob, knob);
            }

            var textRect = new Rectangle((int)track.Right + Theme.Scale(10), 0, Width - (int)track.Right - Theme.Scale(10), Height);
            TextRenderer.DrawText(g, Text, Font, textRect, Enabled ? Theme.TextPrimary : Theme.TextDisabled, Draw.LeftMiddle);
        }
    }
    #endregion

    #region FluentTextBox
    /// <summary>Borderless TextBox hosted in a rounded Fluent shell, with optional icon and placeholder.</summary>
    [DefaultEvent("TextChanged")]
    public class FluentTextBox : UserControl
    {
        private const int EM_SETCUEBANNER = 0x1501;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        private readonly TextBox _box;
        private string _placeholder = string.Empty;
        private char _glyph = '\0';

        public FluentTextBox()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            _box = new TextBox
            {
                BorderStyle = BorderStyle.None,
                BackColor = Theme.Surface,
                ForeColor = Theme.TextPrimary,
            };
            Controls.Add(_box);

            _box.TextChanged += (s, e) => OnTextChanged(e);
            _box.Click += (s, e) => OnClick(e);
            _box.GotFocus += (s, e) => Invalidate();
            _box.LostFocus += (s, e) => Invalidate();
            _box.KeyPress += (s, e) =>
            {
                if (NumericOnly && !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
                    e.Handled = true;
            };
            _box.HandleCreated += (s, e) => ApplyPlaceholder();

            BackColor = Theme.Surface;
            Cursor = Cursors.IBeam;
            Size = new Size(200, 32);
        }

        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public override string Text
        {
            get { return _box.Text; }
            set { _box.Text = value; }
        }

        [DefaultValue("")]
        public string Placeholder
        {
            get { return _placeholder; }
            set { _placeholder = value ?? string.Empty; ApplyPlaceholder(); }
        }

        /// <summary>Icon glyph drawn on the left ('\0' for none).</summary>
        public char Glyph
        {
            get { return _glyph; }
            set { _glyph = value; PerformLayout(); Invalidate(); }
        }

        [DefaultValue(false)]
        public bool NumericOnly { get; set; }

        public int MaxLength
        {
            get { return _box.MaxLength; }
            set { _box.MaxLength = value; }
        }

        public HorizontalAlignment TextAlign
        {
            get { return _box.TextAlign; }
            set { _box.TextAlign = value; }
        }

        public void Clear() => _box.Clear();

        private void ApplyPlaceholder()
        {
            if (_box.IsHandleCreated)
                SendMessage(_box.Handle, EM_SETCUEBANNER, (IntPtr)1, _placeholder);
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            _box.Font = Font;
            PerformLayout();
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            var left = Theme.Scale(10) + (_glyph != '\0' ? Theme.Scale(22) : 0);
            var right = Theme.Scale(10);
            _box.SetBounds(left, (Height - _box.Height) / 2, Math.Max(0, Width - left - right), _box.Height);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            _box.BackColor = Enabled ? Theme.Surface : Theme.SurfaceDisabled;
            Invalidate();
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            if (!_box.Focused)
                _box.Focus();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var bounds = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            var radius = Theme.Scale(4);
            Theme.FillRounded(g, Enabled ? Theme.Surface : Theme.SurfaceDisabled, bounds, radius);
            Theme.DrawRounded(g, !Enabled ? Theme.Stroke : _box.Focused ? Theme.Brand : Theme.StrokeStrong, bounds, radius);

            if (_box.Focused)
            {
                // Fluent focus indicator: accent underline
                using (var brush = new SolidBrush(Theme.Brand))
                    g.FillRectangle(brush, radius, Height - Theme.Scale(2) - 1, Width - radius * 2, Theme.Scale(2));
            }

            if (_glyph != '\0')
                Theme.DrawGlyph(g, _glyph, Enabled ? Theme.TextSecondary : Theme.TextDisabled, new RectangleF(Theme.Scale(8), 0, Theme.Scale(20), Height), Theme.Scale(14));
        }
    }
    #endregion

    #region FluentButton
    public class FluentButton : Button
    {
        private bool _hover;
        private bool _pressed;

        public FluentButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            Font = Theme.BodyStrong;
        }

        [DefaultValue(true)]
        public bool Primary { get; set; } = true;

        public char Glyph { get; set; } = '\0';

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Draw.ParentBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var bounds = new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f);
            var radius = Theme.Scale(4);
            Color fill, text;
            if (!Enabled)
            {
                fill = Theme.StrokeSubtle;
                text = Theme.TextDisabled;
            }
            else if (Primary)
            {
                fill = _pressed ? Theme.BrandPressed : _hover ? Theme.BrandHover : Theme.Brand;
                text = Theme.TextOnBrand;
            }
            else
            {
                fill = _pressed ? Theme.SurfacePressed : _hover ? Theme.SurfaceHover : Theme.Surface;
                text = Theme.TextPrimary;
            }

            Theme.FillRounded(g, fill, bounds, radius);
            if (!Primary && Enabled)
                Theme.DrawRounded(g, Theme.StrokeStrong, bounds, radius);

            var textSize = TextRenderer.MeasureText(g, Text, Font, Size.Empty, TextFormatFlags.NoPadding);
            var iconSize = Glyph != '\0' ? Theme.Scale(16) : 0;
            var gap = Glyph != '\0' ? Theme.Scale(8) : 0;
            var x = (Width - (iconSize + gap + textSize.Width)) / 2;
            if (Glyph != '\0')
                Theme.DrawGlyph(g, Glyph, text, new RectangleF(x, 0, iconSize, Height), Theme.Scale(14));
            TextRenderer.DrawText(g, Text, Font, new Rectangle(x + iconSize + gap, 0, textSize.Width + 2, Height), text, Draw.LeftMiddle);

            if (Focused && ShowFocusCues)
                Theme.DrawRounded(g, Theme.TextPrimary, new RectangleF(1.5f, 1.5f, Width - 3.5f, Height - 3.5f), radius);
        }
    }
    #endregion

    #region FluentComboBox
    /// <summary>Owner drawn DropDownList with a flat Fluent look.</summary>
    public class FluentComboBox : ComboBox
    {
        private const int WM_PAINT = 0x000F;

        public FluentComboBox()
        {
            DrawMode = DrawMode.OwnerDrawFixed;
            DropDownStyle = ComboBoxStyle.DropDownList;
            FlatStyle = FlatStyle.Flat;
            ItemHeight = Theme.Scale(26);
            BackColor = Theme.Surface;
            ForeColor = Theme.TextPrimary;
            IntegralHeight = false;
            MaxDropDownItems = 15;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            var isEdit = (e.State & DrawItemState.ComboBoxEdit) != 0;
            var selected = (e.State & DrawItemState.Selected) != 0 && !isEdit;
            var back = isEdit ? (Enabled ? Theme.Surface : Theme.SurfaceDisabled) : selected ? Theme.BrandSubtle : Theme.Surface;

            using (var brush = new SolidBrush(back))
                e.Graphics.FillRectangle(brush, e.Bounds);

            if (selected)
            {
                using (var brush = new SolidBrush(Theme.Brand))
                    e.Graphics.FillRectangle(brush, e.Bounds.X, e.Bounds.Y + Theme.Scale(6), Theme.Scale(3), e.Bounds.Height - Theme.Scale(12));
            }

            if (e.Index >= 0)
            {
                var text = GetItemText(Items[e.Index]);
                var rect = new Rectangle(e.Bounds.X + Theme.Scale(isEdit ? 6 : 10), e.Bounds.Y, e.Bounds.Width - Theme.Scale(16), e.Bounds.Height);
                TextRenderer.DrawText(e.Graphics, text, Font, rect, Enabled ? Theme.TextPrimary : Theme.TextDisabled, Draw.LeftMiddle);
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg != WM_PAINT || !IsHandleCreated)
                return;

            using (var g = Graphics.FromHwnd(Handle))
            {
                var parentBack = Draw.ParentBack(this);
                using (var pen = new Pen(parentBack))
                    g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);

                var buttonWidth = Theme.Scale(26);
                var button = new Rectangle(Width - buttonWidth - 1, 1, buttonWidth, Height - 2);
                using (var brush = new SolidBrush(Enabled ? Theme.Surface : Theme.SurfaceDisabled))
                    g.FillRectangle(brush, button);
                Theme.DrawGlyph(g, Theme.Glyph.ChevronDown, Enabled ? Theme.TextSecondary : Theme.TextDisabled, button, Theme.Scale(10));

                g.SmoothingMode = SmoothingMode.AntiAlias;
                var border = !Enabled ? Theme.Stroke : (Focused || DroppedDown) ? Theme.Brand : Theme.StrokeStrong;
                Theme.DrawRounded(g, border, new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f), Theme.Scale(4));
            }
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void OnDropDownClosed(EventArgs e) { base.OnDropDownClosed(e); Invalidate(); }
        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); Invalidate(); }
    }

    /// <summary>Hosts a <see cref="FluentComboBox"/> in a ToolStrip, exposing the ToolStripComboBox members the tool uses.</summary>
    public class ToolStripFluentComboBox : ToolStripControlHost
    {
        public ToolStripFluentComboBox() : base(new FluentComboBox())
        {
            AutoSize = false;
        }

        public FluentComboBox ComboBox => (FluentComboBox)Control;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public ComboBox.ObjectCollection Items => ComboBox.Items;

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public object SelectedItem
        {
            get { return ComboBox.SelectedItem; }
            set { ComboBox.SelectedItem = value; }
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int SelectedIndex
        {
            get { return ComboBox.SelectedIndex; }
            set { ComboBox.SelectedIndex = value; }
        }

        public int DropDownWidth
        {
            get { return ComboBox.DropDownWidth; }
            set { ComboBox.DropDownWidth = value; }
        }

        public event EventHandler SelectedIndexChanged;

        protected override Size DefaultSize => new Size(160, 32);

        protected override void OnSubscribeControlEvents(Control control)
        {
            base.OnSubscribeControlEvents(control);
            ((ComboBox)control).SelectedIndexChanged += HandleSelectedIndexChanged;
        }

        protected override void OnUnsubscribeControlEvents(Control control)
        {
            base.OnUnsubscribeControlEvents(control);
            ((ComboBox)control).SelectedIndexChanged -= HandleSelectedIndexChanged;
        }

        private void HandleSelectedIndexChanged(object sender, EventArgs e) => SelectedIndexChanged?.Invoke(this, e);
    }
    #endregion

    #region ToolStrip renderer
    /// <summary>
    /// Fluent command bar: white surface, bottom hairline, rounded hover states.
    /// Items tagged "primary" are drawn as the accent button, items tagged "link" as quiet links.
    /// </summary>
    public class FluentToolStripRenderer : ToolStripProfessionalRenderer
    {
        public FluentToolStripRenderer()
        {
            RoundedEdges = false;
        }

        private static bool Is(ToolStripItem item, string tag) => tag.Equals(item.Tag as string);

        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            using (var brush = new SolidBrush(Theme.Surface))
                e.Graphics.FillRectangle(brush, e.AffectedBounds);
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            using (var pen = new Pen(Theme.Stroke))
                e.Graphics.DrawLine(pen, 0, e.ToolStrip.Height - 1, e.ToolStrip.Width, e.ToolStrip.Height - 1);
        }

        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
        {
            var item = e.Item;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = new RectangleF(0.5f, 1.5f, item.Width - 1.5f, item.Height - 3f);
            var radius = Theme.Scale(4);

            if (Is(item, "primary"))
            {
                var fill = !item.Enabled ? Theme.StrokeSubtle
                    : item.Pressed ? Theme.BrandPressed
                    : item.Selected ? Theme.BrandHover
                    : Theme.Brand;
                Theme.FillRounded(g, fill, bounds, radius);
                return;
            }

            if (!item.Enabled)
                return;

            if (item.Pressed)
                Theme.FillRounded(g, Theme.SurfacePressed, bounds, radius);
            else if (item.Selected)
                Theme.FillRounded(g, Theme.Blend(Theme.Surface, Theme.SurfacePressed, 0.45f), bounds, radius);
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            if (!e.Item.Enabled)
                e.TextColor = Theme.TextDisabled;
            else if (Is(e.Item, "primary"))
                e.TextColor = Theme.TextOnBrand;
            else if (Is(e.Item, "link"))
                e.TextColor = e.Item.Selected ? Theme.Brand : Theme.TextSecondary;
            else
                e.TextColor = Theme.TextPrimary;

            base.OnRenderItemText(e);
        }

        protected override void OnRenderItemImage(ToolStripItemImageRenderEventArgs e)
        {
            if (e.Item.Enabled || e.Image == null)
            {
                base.OnRenderItemImage(e);
                return;
            }

            // Disabled: tint the glyph with the disabled text color instead of the default grayscale
            var c = Theme.TextDisabled;
            var matrix = new ColorMatrix(new[]
            {
                new float[] { 0, 0, 0, 0, 0 },
                new float[] { 0, 0, 0, 0, 0 },
                new float[] { 0, 0, 0, 0, 0 },
                new float[] { 0, 0, 0, 1, 0 },
                new float[] { c.R / 255f, c.G / 255f, c.B / 255f, 0, 1 },
            });
            using (var attributes = new ImageAttributes())
            {
                attributes.SetColorMatrix(matrix);
                var r = e.ImageRectangle;
                e.Graphics.DrawImage(e.Image, r, 0, 0, e.Image.Width, e.Image.Height, GraphicsUnit.Pixel, attributes);
            }
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            var x = e.Item.Width / 2;
            using (var pen = new Pen(Theme.Stroke))
                e.Graphics.DrawLine(pen, x, Theme.Scale(8), x, e.Item.Height - Theme.Scale(8));
        }
    }
    #endregion
}
