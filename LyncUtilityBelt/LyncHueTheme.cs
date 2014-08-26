using Q42.HueApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LyncUtilityBelt
{
	public class LyncHueTheme
	{
		public LightCommand Available { get; set; }
		public LightCommand Busy { get; set; }
		public LightCommand Away { get; set; }
		public LightCommand Off { get; set; }

		private static LightCommand OFF = new LightCommand
		{
			On = false
		};

		// ColorTemperature/ColorMode/ColorCoordinates/Hue/Brightness/Saturation

		// 287/xy/0.4077,0.5154/25718/37/254
		private static LightCommand GREEN = new LightCommand
		{
			On = true,
			ColorTemperature = 287,
			ColorCoordinates = new double[] { 0.4077, 0.5154 },
			Hue = 25718,
			Brightness = 37,
			Saturation = 254
		};

		// 500/xy/0.6736,0.3221/65527/86/253
		private static LightCommand RED = new LightCommand
		{
			On = true,
			ColorTemperature = 500,
			ColorCoordinates = new double[] { 0.6736, 0.3221 },
			Hue = 65527,
			Brightness = 86,
			Saturation = 253
		};

		// 500/xy/0.5494,0.4133/11960/166/252
		private static LightCommand ORANGE = new LightCommand
		{
			On = true,
			ColorTemperature = 500,
			ColorCoordinates = new double[] { 0.5494, 0.4133 },
			Hue = 11960,
			Brightness = 166,
			Saturation = 252
		};

		public static LyncHueTheme LYNC_PRESENCE_THEME = new LyncHueTheme
		{
			Available = GREEN,
			Busy = RED,
			Away = ORANGE,
			Off = OFF
		};

		public static LyncHueTheme TRAFFIC_LIGHT_THEME = new LyncHueTheme
		{
			Available = ORANGE,
			Busy = RED,
			Away = GREEN,
			Off = OFF
		};
	}
}
