using Microsoft.Win32;
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
		private OutlookWorkHours _outlookWorkHours;
		private WhoIsCallingMe _whoIsCallingMe;
		private LyncHue _lyncHue;

		private NotifyIcon _icon;
		private LyncUtilityBeltConfig _config;

		private RegistryKey _runKey;
		private const string RUN_KEY_VALUE = "LyncUtilityBelt";

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
			AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

			_runKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

			base.OnStartup(e);

			_icon = new NotifyIcon();
			_icon.Icon = SysTrayIcon;
			_icon.Visible = true;
			_icon.Text = "Lync Utility Belt";
			_icon.ContextMenu = new System.Windows.Forms.ContextMenu(new System.Windows.Forms.MenuItem[] {
				new System.Windows.Forms.MenuItem("Outlook Work Hours", owh_Toggle),
				new System.Windows.Forms.MenuItem("Who Is Calling Me?", wicm_Toggle),
				new System.Windows.Forms.MenuItem("Philips Hue Lync", new System.Windows.Forms.MenuItem[] {
					new System.Windows.Forms.MenuItem("Bridges", new System.Windows.Forms.MenuItem[] {}),
					new System.Windows.Forms.MenuItem("Lights", new System.Windows.Forms.MenuItem[] {})
				}),
				new System.Windows.Forms.MenuItem("--"),
				new System.Windows.Forms.MenuItem("Run at Startup", rasu_Toggle),
				new System.Windows.Forms.MenuItem("Exit", icon_Exit)
			});

			_icon.ContextMenu.MenuItems[4].Checked = (_runKey.GetValue(RUN_KEY_VALUE) != null);

			_config = LyncUtilityBeltConfig.Load();

			_outlookWorkHours = new OutlookWorkHours(_icon);
			if (_config.OutlookWorkHoursEnabled)
			{
				_icon.ContextMenu.MenuItems[0].Checked = true;
				_outlookWorkHours.Start();
			}

			_whoIsCallingMe = new WhoIsCallingMe(_icon);
			if (_config.WhoIsCallingMeEnabled)
			{
				_icon.ContextMenu.MenuItems[1].Checked = true;
				_whoIsCallingMe.Start();
			}

			_lyncHue = new LyncHue(_icon, _icon.ContextMenu.MenuItems[2].MenuItems[0], _icon.ContextMenu.MenuItems[2].MenuItems[1], _config);
			_lyncHue.Start(); // always started unless we find a way to make the submenu 'unclickable'

			base.OnStartup(e);
		}

		void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			var ex = (Exception)e.ExceptionObject;
			var msg = DumpException(ex);
			var file = string.Format("LyncUtilityBelt-{0:yyyyMMdd-HHmmss}.err.txt", DateTime.Now);
			var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), file);
			File.WriteAllText(path, msg);
		}

		string DumpException(Exception ex) 
		{
			var msg = ex.Message + Environment.NewLine + ex.StackTrace;
			if (ex.InnerException != null)
				msg = msg + Environment.NewLine + "----" + DumpException(ex.InnerException);
			return msg;
		}

		private void owh_Toggle(object sender, EventArgs e)
		{
			var mi = _icon.ContextMenu.MenuItems[0];
			mi.Checked = !mi.Checked;
			if (mi.Checked)
				_outlookWorkHours.Start();
			else
				_outlookWorkHours.Stop();
			_config.OutlookWorkHoursEnabled = mi.Checked;
			_config.Save();
		}

		private void wicm_Toggle(object sender, EventArgs e)
		{
			var mi = _icon.ContextMenu.MenuItems[1];
			mi.Checked = !mi.Checked;
			if (mi.Checked)
				_whoIsCallingMe.Start();
			else
				_whoIsCallingMe.Stop();
			_config.WhoIsCallingMeEnabled = mi.Checked;
			_config.Save();
		}

		private void rasu_Toggle(object sender, EventArgs e)
		{
			var mi = _icon.ContextMenu.MenuItems[4];
			mi.Checked = !mi.Checked;
			_runKey.DeleteValue(RUN_KEY_VALUE, false);
			if (mi.Checked)
				_runKey.SetValue(RUN_KEY_VALUE, System.Reflection.Assembly.GetExecutingAssembly().Location);
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
