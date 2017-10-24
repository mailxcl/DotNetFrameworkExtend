using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace System.Windows.Forms
{
    [ToolboxBitmap(typeof(System.Windows.Forms.ProgressBar))]
    public class ProgressBarExtend : System.Windows.Forms.ProgressBar
    {
        private int gap = 0;
        private SolidBrush stripBrush = null;
        private SolidBrush stringBrush = null;
        private SolidBrush backgroundBrush = null;
        private Font stringFont = new Font("微软雅黑", 12f, FontStyle.Regular);
        private bool isCustom { get; set; }

        public EnumCustomProgressBarType ProgressType { get; set; }

        public ProgressBarExtend()
            : this(EnumCustomProgressBarType.Bar)
        {
        }

        public ProgressBarExtend(EnumCustomProgressBarType type)
        {
            ProgressType = type;

            this.SetStyle(System.Windows.Forms.ControlStyles.UserPaint, true);
            this.SetStyle(System.Windows.Forms.ControlStyles.OptimizedDoubleBuffer, true);
            this.SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, true);
            this.Step = 1;
            this.ForeColor = Color.FromArgb(141, 223, 114);
            this.BackColor = Color.Beige;
            stripBrush = new SolidBrush(this.ForeColor);
            stringBrush = new SolidBrush(Color.Blue);
            backgroundBrush = new SolidBrush(SystemColors.Control);
        }

        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            Rectangle rect = new Rectangle(0, 0, this.Width, this.Height);
            if (ProgressType == EnumCustomProgressBarType.Bar)
            {
                if (System.Windows.Forms.ProgressBarRenderer.IsSupported)
                {
                    System.Windows.Forms.ProgressBarRenderer.DrawHorizontalBar(e.Graphics, rect);
                }

                rect.Height -= 2 * gap;
                rect.Width = (int)(rect.Width * ((double)Value / Maximum)) - 2 * gap;

                if (System.Windows.Forms.ProgressBarRenderer.IsSupported)
                {
                    System.Windows.Forms.ProgressBarRenderer.DrawHorizontalChunks(e.Graphics, rect);
                }
                //e.Graphics.FillRectangle(stripBrush, gap, gap, rect.Width, rect.Height);
                e.Graphics.DrawString(Value + " %", stringFont, stringBrush, rect.X + Width / 2, rect.Y);
            }
            else if (ProgressType == EnumCustomProgressBarType.Pie)
            {
                e.Graphics.FillRectangle(backgroundBrush, 0, 0, rect.Width, rect.Height);

                e.Graphics.FillPie(SystemBrushes.ControlDark, rect, 0, 360);
                e.Graphics.FillPie(stripBrush, rect, 0, 360 * Value / (Maximum - Minimum));
                int minus = rect.Width / 2;
                Rectangle rect2 = new Rectangle(minus / 2, minus / 2, this.Width - minus, this.Height - minus);
                e.Graphics.FillPie(backgroundBrush, rect2, 0, 360);

                e.Graphics.DrawString(
                    Value.ToString(),
                    stringFont, stringBrush,
                    rect.X + Width / 2 - stringFont.Size,
                    rect.Y + Height / 2 - stringFont.Size);
            }
        }

    }

    public enum EnumCustomProgressBarType
    {
        Bar,
        Pie
    }
}
