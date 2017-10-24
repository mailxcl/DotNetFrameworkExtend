using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace System.Windows.Forms
{
    [ToolboxBitmap(typeof(System.Windows.Forms.ProgressBar))]
    public class CustomProgressBar : System.Windows.Forms.ProgressBar
    {
        private int gap = 0;
        private SolidBrush stripBrush = null;
        private SolidBrush stringBrush = null;
        private Font stringFont = new Font("微软雅黑", 12f, FontStyle.Regular);

        public EnumCustomProgressBarType ProgressType { get; set; }

        public CustomProgressBar()
            : this(EnumCustomProgressBarType.Bar)
        {
        }

        public CustomProgressBar(EnumCustomProgressBarType type)
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
        }

        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            Rectangle rect = new Rectangle(0, 0, this.Width, this.Height);
            if (System.Windows.Forms.ProgressBarRenderer.IsSupported)
            {
                System.Windows.Forms.ProgressBarRenderer.DrawHorizontalBar(e.Graphics, rect);
            }

            if (ProgressType == EnumCustomProgressBarType.Bar)
            {
                rect.Height -= 2 * gap;
                rect.Width = (int)(rect.Width * ((double)Value / Maximum)) - 2 * gap;

                e.Graphics.FillRectangle(stripBrush, gap, gap, rect.Width, rect.Height);
                e.Graphics.DrawString(Value + " %", stringFont, stringBrush, rect.X + Width / 2, rect.Y);
            }
            else if (ProgressType == EnumCustomProgressBarType.Pie)
            {
                e.Graphics.FillPie(stripBrush, rect, 0, 360 * Value / (Maximum - Minimum));
            }
        }
    }

    public enum EnumCustomProgressBarType
    {
        Bar,
        Pie
    }
}
