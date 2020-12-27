using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSWordDocument
{
	public class BasicStyle
	{
		public bool? Bold { get; set; }
		public bool? Italic { get; set; }
		public string FontFamily { get; set; }
		public int? FontSize { get; set; }

		public string Color { get; set; }

		public List<string> GetProperties()
		{
			List<string> props = new List<string>();
			if (Bold.HasValue && Bold.Value)
				props.Add("bold");
			if (Italic.HasValue && Italic.Value)
				props.Add("italic");
			if (!string.IsNullOrEmpty(FontFamily))
				props.Add($"FontFamily:{FontFamily}");
			if (FontSize.HasValue)
				props.Add($"FontSize:{FontSize.Value}");
			if (!string.IsNullOrEmpty(Color))
				props.Add($"Color:{Color}");
			return props;
		}
	}
}
