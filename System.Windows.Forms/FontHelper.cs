using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace System.Windows.Forms
{
    public class FontHelper : AxHost
    {
        public FontHelper()
            : base("1EDBC8C7-21A0-4FD3-9BB6-F59AF78A6691")
        { }

        public static Font GetFontFrom(stdole.IFontDisp fntDisp)
        {
            return GetFontFromIFontDisp(fntDisp);
        }

        public static stdole.IFontDisp GetIFontDispFrom(Font fnt)
        {
            return GetIFontDispFromFont(fnt) as stdole.IFontDisp;
        }
    }
}
