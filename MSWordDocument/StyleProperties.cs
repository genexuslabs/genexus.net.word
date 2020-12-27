using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Genexus.Word
{
    internal class StyleProperties
    {
        // Indicates how we should consider font sizes. 
        private static bool? s_doubleByte;

        private static Dictionary<string, Func<string, OpenXmlElement>> m_Properties = new Dictionary<string, Func<string, OpenXmlElement>>()
        {
            {  "bold", (_) => new Bold() },
            {  "italic", (_) => new Italic() },
            {  "caps", (_) => new Caps() },
            {  "smallcaps", (_) => new SmallCaps() },
            {  "strike", (_) => new Strike() },
            {  "doublestrike", (_) => new DoubleStrike() },
            {  "outline", (_) => new DoubleStrike() },
            {  "shadow", (_) => new Shadow() },
            {  "emboss", (_) => new Emboss() },
            {  "snaptogrid", (_) => new SnapToGrid() },
            {  "highlight", (_) => new Highlight() },
            {  "underline", (_) => new Underline() { Val = UnderlineValues.Single }  },
            {  "fontsize",(size) =>
            {
                if (IsDoubleByte())
				{
                    float f = float.Parse(size);
                    size = (f * 2).ToString();
				}                    
                return new FontSize() { Val = size};
              } },
            {  "color",(color) => new Color() { Val = color.ToString() } },
            {  "fontfamily",(fname) => new RunFonts() { HighAnsi = fname, EastAsia = fname, Ascii = fname } }
        };

     
        public static bool IsDoubleByte()
		{
            if (!s_doubleByte.HasValue)
                return Thread.CurrentThread.CurrentCulture.TwoLetterISOLanguageName.Trim().Contains("ja");
            return s_doubleByte.Value;
		}
        public static void SetDoubleByte(bool v)
		{
            s_doubleByte = v;
		}
        internal static bool Exists(string name)
        {
            string[] parts = name.Split(':');
            string propName = parts[0].ToLower();
            return m_Properties.ContainsKey(propName.ToLower());
        }

        internal static OpenXmlElement RunFunctionProperty(string prop)
        {
            string[] parts = prop.Split(':');
            string propName = parts[0].ToLower();
            string parm1 = String.Empty;
            if (parts.Length > 1)
			{
                parm1 = parts[1].Trim();
			}
            if (Exists(propName))
                return m_Properties[propName](parm1);
            Debug.Assert(false);
            ExceptionManager.LogException($"Asking for a non existing style -{prop}-");
            return null;
        }
    }
}
