using Microsoft.Lync.Model;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LyncUtilityBelt
{
	public class OutlookWorkHours
	{
		// read work hours from HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Calendar
		//	WorkDay, CalDefStart & CalDefEnd (minutes from midnight local time)
		[Flags]
		public enum WorkDay : int
		{
			Saturday = 2,
			Friday = 4,
			Thursday = 8,
			Wednesday = 16,
			Tuesday = 32,
			Monday = 64,
			Sunday = 128
		}

		private NotifyIcon _icon;

		private WorkDay _workDays;
		private TimeSpan _calDefStart;
		private TimeSpan _calDefEnd;

		private System.Threading.Timer _dispatcher;

		public OutlookWorkHours(NotifyIcon icon)
		{
			_icon = icon;

			ReadWorkHours();
		}

		public void Start() {
			_dispatcher = new System.Threading.Timer(dispatcher_Tick, null, 0, 60 * 1000);

		}

		public void Stop() {
			_dispatcher.Dispose();
			_dispatcher = null;
		}

		private void dispatcher_Tick(object state) //sender, EventArgs e)
		{
 			ReadWorkHours();
			ContactAvailability availability;
			string activityId;
			var ok = AvailabilityMonitor.TryGetSelfAvailability(out availability, out activityId);
			if(ok)
				if(InWorkHours()) 
				{
					if (availability == ContactAvailability.Away && string.Equals("off-work", activityId))
					{
						_icon.ShowBalloonTip(5000, "Available", "Updating your presence to free", ToolTipIcon.Info);
						AvailabilityMonitor.SetSelfAvailability(ContactAvailability.Free);
					}
				} 
				else
				{
					// make sure we aren't available
					if (availability != ContactAvailability.Away || !string.Equals("off-work", activityId))
					{
						_icon.ShowBalloonTip(5000, "Off Work", "Updating your presence to off-work", ToolTipIcon.Info);
						AvailabilityMonitor.SetSelfAvailability(ContactAvailability.Away, "off-work");
					}
				}
			// else we don't know what the availability is so it doesn't matter
		}

		private string GetOutlookVersion()
		{
			return "15.0";
		}

		private const string CALENDAR_OPTIONS_PATH = @"HKEY_CURRENT_USER\Software\Microsoft\Office\{0}\Outlook\Options\Calendar";
		private void ReadWorkHours()
		{
			var path = string.Format(CALENDAR_OPTIONS_PATH, GetOutlookVersion());
			_workDays = (WorkDay)(int)Registry.GetValue(path, "WorkDay", null);
			_calDefStart = new TimeSpan(TimeSpan.TicksPerMinute * (int)Registry.GetValue(path, "CalDefStart", null));
			_calDefEnd = new TimeSpan(TimeSpan.TicksPerMinute * (int)Registry.GetValue(path, "CalDefEnd", null));

			//_icon.ShowBalloonTip(10000, "Work Hours", string.Format(@"{0} {1:hh\:mm}-{2:hh\:mm}", _workDays, _calDefStart, _calDefEnd), ToolTipIcon.Info);
		}

		private static readonly Dictionary<DayOfWeek, WorkDay> DAY_OF_WEEK_MAP = new Dictionary<DayOfWeek, WorkDay> 
		{
			{DayOfWeek.Sunday, WorkDay.Sunday}, 
			{DayOfWeek.Monday, WorkDay.Monday}, 
			{DayOfWeek.Tuesday, WorkDay.Tuesday}, 
			{DayOfWeek.Wednesday, WorkDay.Wednesday}, 
			{DayOfWeek.Thursday, WorkDay.Thursday}, 
			{DayOfWeek.Friday, WorkDay.Friday}, 
			{DayOfWeek.Saturday, WorkDay.Saturday}, 
		};

		private bool InWorkHours() {
			var now = DateTime.Now;

			var isWorkDay = (_workDays & DAY_OF_WEEK_MAP[now.DayOfWeek]) == DAY_OF_WEEK_MAP[now.DayOfWeek];
			var isAfterStart = now.TimeOfDay >= _calDefStart;
			var isBeforeEnd = now.TimeOfDay < _calDefEnd;

			return isWorkDay && isAfterStart && isBeforeEnd;
		}
	}
}
