using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSWordDocument
{
	public class BasicStyle
	{
		public bool Bold { get; set; } = false;
		public bool Italic { get; set; } = false;
		public string FontFamily { get; set; }
		public int FontSize { get; set; } = 0;

		public string Color { get; set; }

		public List<string> GetProperties()
		{
			List<string> props = new List<string>();
			if (Bold)
				props.Add("bold");
			if (Italic)
				props.Add("italic");
			if (!string.IsNullOrEmpty(FontFamily))
				props.Add($"FontFamily:{FontFamily}");
			if (FontSize > 0)
				props.Add($"FontSize:{FontSize}");
			if (!string.IsNullOrEmpty(Color))
				props.Add($"Color:{Color}");
			return props;
		}
	}
}
