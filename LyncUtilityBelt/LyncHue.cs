using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Q42.HueApi;

namespace LyncUtilityBelt
{
	// maybe use https://github.com/cDima/Hue instead of Q42.HueApi since it has fewer/no deps (inc .net 4.5)
	public class LyncHue
	{
		private NotifyIcon _icon;
		private MenuItem _bridgeMenu;
		private MenuItem _lightMenu;
		private ILyncHueConfig _config;

		private HueClient _hue;
		private AvailabilityMonitor _monitor;
		public LyncHue(NotifyIcon icon, MenuItem bridgeMenu, MenuItem lightMenu, ILyncHueConfig config)
		{
			_icon = icon;
			_bridgeMenu = bridgeMenu;
			_lightMenu = lightMenu;
			_config = config;

			_monitor = new AvailabilityMonitor();
			_monitor.AvailabilityChanged += monitor_AvailabilityChanged;
			_monitor.Initialize();
		}

		public void Start()
		{
			// Creating Bridge Locator
			var locator = new HttpBridgeLocator();
			var bridgeIPs = locator.LocateBridgesAsync(TimeSpan.FromSeconds(5));

			//var locator = new SSDPBridgeLocator();
			//var bridgeIPs = locator.LocateBridgesAsync(TimeSpan.FromSeconds(5));

			// Waiting for Bridge Locator
			bridgeIPs.Wait();

			var choices = bridgeIPs.Result.OrderBy(x => x).ToList();
			UpdateBridgeChoice(choices);
		}

		public void Stop()
		{
			_hue = null; 
			// this stop's UpdateLight from sending any more commands
			// but still continues to record the last/current color so it can be
			// sent when it's turned back on again
		}

		#region Hue interface
		const string APP_NAME = "LyncHue";
		const int RETRY_MAX = 60;
		const int RETRY_DELAY = 1000;

		public void SetBridge(string bridgeIP, string appKey, bool alreadyRegistered)
		{
			_hue = new HueClient(bridgeIP);

			if (alreadyRegistered)
			{
				_hue.Initialize(appKey);
			}
			else
			{
				var retryCount = 0;
				var regOk = false;
				while (!regOk || retryCount >= RETRY_MAX)
				{
					var reg = _hue.RegisterAsync(APP_NAME, appKey);
					reg.Wait();
					if (!reg.Result)
						ShowWarning("Please press the button on the bridge to register the application", 30);
					else
						regOk = true;

					retryCount++;
					if (!regOk && retryCount < RETRY_MAX)
						Thread.Sleep(RETRY_DELAY);
				}

				if (!regOk)
					ShowError("Failed to register application with bridge", 30);
			}

			var lights = _hue.GetLightsAsync();
			lights.Wait();

			var choices = lights.Result.Select(x => new Tuple<string, string>(x.Id, string.Format("{0} ({1})", x.Name, x.Type))).ToList();
			UpdateLightChoice(choices);
		}

		private string _lightID;
		public void SetLight(string lightID)
		{
			_lightID = lightID;

			var lastCmd = _lastCmd;

			// turn on (if off) & blink light to acknowledge/confirm
			UpdateLight(new LightCommand().TurnOn());
			UpdateLight(new LightCommand { Alert = Alert.Once });

			if (lastCmd != null)
				UpdateLight(lastCmd);
		}

		private LightCommand _lastCmd = null;
		public void UpdateLight(LightCommand cmd)
		{
			_lastCmd = cmd;
			if (_hue != null)
				_hue.SendCommandAsync(cmd, new string[] { _lightID });
		}
		#endregion

		#region icon/menu updaters
		void ShowError(string msg, int timeout = 10)
		{
			_icon.ShowBalloonTip(timeout * 1000, "Error", msg, System.Windows.Forms.ToolTipIcon.Error);
		}
		void ShowWarning(string msg, int timeout = 10)
		{
			_icon.ShowBalloonTip(timeout * 1000, "Warning", msg, System.Windows.Forms.ToolTipIcon.Warning);
		}
		void ShowInfo(string msg, int timeout = 10)
		{
			_icon.ShowBalloonTip(timeout * 1000, "Info", msg, System.Windows.Forms.ToolTipIcon.Info);
		}

		void UpdateBridgeChoice(ICollection<string> bridges)
		{
			var mi = _bridgeMenu;
			mi.MenuItems.Clear();
			var items = bridges.Select(x => new System.Windows.Forms.MenuItem(x, hue_ClickBridge) { Checked = string.Equals(_config.BridgeIP, x) }).ToArray();
			mi.MenuItems.AddRange(items);

			var curr = items.FirstOrDefault(x => x.Checked);
			if (curr != null)
				SetBridge(curr.Text, _config.AppKey, true);
			else if (items.Length == 0)
				ShowWarning("No Philips Hue bridges");
			else if (items.Length == 1)
			{
				_config.BridgeIP = items[0].Text;
				_config.Save();
				items[0].Checked = true;
				SetBridge(_config.BridgeIP, _config.AppKey, false);
			}
			else
				ShowInfo("Please select your Philips Hue bridge", 20);

		}

		void UpdateLightChoice(ICollection<Tuple<string, string>> lights)
		{
			var mi = _lightMenu;
			mi.MenuItems.Clear();
			var items = lights.Select(x => new System.Windows.Forms.MenuItem(x.Item2, hue_ClickLight) { Tag = x.Item1, Checked = string.Equals(_config.LightID, x.Item1) }).ToArray();
			mi.MenuItems.AddRange(items);

			var curr = items.FirstOrDefault(x => x.Checked);
			if (curr != null)
				SetLight(curr.Tag as string);
			else if (items.Length == 0)
				ShowWarning("No Philips Hue lights");
			else if (items.Length == 1)
			{
				_config.LightID = items[0].Tag as string;
				_config.Save();
				items[0].Checked = true;
				SetLight(_config.LightID);
			}
			else
				ShowInfo("Please select your Philips Hue light", 20);
		}
		#endregion

		#region Callback functions
		// TODO: make theme selectable via context menu/config
		private LyncHueTheme _lightTheme = LyncHueTheme.TRAFFIC_LIGHT_THEME;
		void monitor_AvailabilityChanged(Microsoft.Lync.Model.ContactAvailability availability, string activityId)
		{
			Q42.HueApi.LightCommand cmd = null;
			switch (availability)
			{
				case Microsoft.Lync.Model.ContactAvailability.Away:
				case Microsoft.Lync.Model.ContactAvailability.TemporarilyAway:
					cmd = _lightTheme.Away;
					break;
				case Microsoft.Lync.Model.ContactAvailability.Busy:
				case Microsoft.Lync.Model.ContactAvailability.BusyIdle:
				case Microsoft.Lync.Model.ContactAvailability.DoNotDisturb:
					cmd = _lightTheme.Busy;
					break;
				case Microsoft.Lync.Model.ContactAvailability.Free:
				case Microsoft.Lync.Model.ContactAvailability.FreeIdle:
					cmd = _lightTheme.Available;
					break;
				case Microsoft.Lync.Model.ContactAvailability.Invalid:
				case Microsoft.Lync.Model.ContactAvailability.None:
				case Microsoft.Lync.Model.ContactAvailability.Offline:
					cmd = _lightTheme.Off;
					break;
			}

			if (cmd != null)
				UpdateLight(cmd);
		}

		private void hue_ClickLight(object sender, EventArgs e)
		{
			var mi = (System.Windows.Forms.MenuItem)sender;
			foreach (System.Windows.Forms.MenuItem x in mi.Parent.MenuItems)
				if (x.Checked)
					x.Checked = false;
			mi.Checked = true;

			var choice = mi.Tag as string;
			_config.LightID = choice;
			_config.Save();
			SetLight(choice);
		}

		private void hue_ClickBridge(object sender, EventArgs e)
		{
			var mi = (System.Windows.Forms.MenuItem)sender;
			foreach (System.Windows.Forms.MenuItem x in mi.Parent.MenuItems)
				if (x.Checked)
					x.Checked = false;
			mi.Checked = true;

			var choice = mi.Text;
			_config.BridgeIP = choice;
			_config.Save();
			SetBridge(choice, _config.AppKey, false);
		}
		#endregion
	}
}
