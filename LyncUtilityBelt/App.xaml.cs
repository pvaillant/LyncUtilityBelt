using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Resources;
using System.Windows;
using NotifyIcon = System.Windows.Forms.NotifyIcon;

namespace LyncUtilityBelt
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{
		private WhoIsCallingMe _whoIsCallingMe;
		private LyncHue _lyncHue;

		private NotifyIcon _icon;
		private LyncUtilityBeltConfig _config;

		private const string ICO_NAME = "logo.ico";
		private static Icon SysTrayIcon
		{
			get
			{
				var self = typeof(App).Assembly;
				var resStream = self.GetManifestResourceStream("LyncUtilityBelt.g.resources");
				using (var resRdr = new ResourceReader(resStream))
				{
					foreach (DictionaryEntry resEntry in resRdr)
					{
						if (resEntry.Key.ToString().ToLower().Equals(ICO_NAME))
							return new System.Drawing.Icon(resEntry.Value as Stream);
					}
				}
				throw new Exception(ICO_NAME + " not found");
			}
		}

		protected override void OnStartup(StartupEventArgs e)
		{
			base.OnStartup(e);

			_icon = new NotifyIcon();
			_icon.Icon = SysTrayIcon;
			_icon.Visible = true;
			_icon.Text = "Lync Utility Belt";
			_icon.ContextMenu = new System.Windows.Forms.ContextMenu(new System.Windows.Forms.MenuItem[] {
				new System.Windows.Forms.MenuItem("Who Is Calling Me?", wicm_Toggle),
				new System.Windows.Forms.MenuItem("Philips Hue Lync", new System.Windows.Forms.MenuItem[] {
					new System.Windows.Forms.MenuItem("Bridges", new System.Windows.Forms.MenuItem[] {}),
					new System.Windows.Forms.MenuItem("Lights", new System.Windows.Forms.MenuItem[] {})
				}),
				new System.Windows.Forms.MenuItem("--"),
				new System.Windows.Forms.MenuItem("Exit", icon_Exit)
			});

			_config = LyncUtilityBeltConfig.Load();

			_whoIsCallingMe = new WhoIsCallingMe(_icon);
			if (_config.WhoIsCallingMeEnabled)
				wicm_Toggle(null, null);

			_lyncHue = new LyncHue(_icon, _icon.ContextMenu.MenuItems[1].MenuItems[0], _icon.ContextMenu.MenuItems[1].MenuItems[1], _config);
			_lyncHue.Start(); // always started unless we find a way to make the submenu 'unclickable'

			base.OnStartup(e);
		}

		private void wicm_Toggle(object sender, EventArgs e)
		{
			var mi = _icon.ContextMenu.MenuItems[0];
			mi.Checked = !mi.Checked;
			if (mi.Checked)
				_whoIsCallingMe.Start();
			else
				_whoIsCallingMe.Stop();
			_config.WhoIsCallingMeEnabled = mi.Checked;
			_config.Save();
		}

		private void icon_Exit(object sender, EventArgs e)
		{
			this.Shutdown();
		}

		protected override void OnExit(ExitEventArgs e)
		{
			if (this._icon != null)
				this._icon.Dispose();

			base.OnExit(e);
		}
	}
}
