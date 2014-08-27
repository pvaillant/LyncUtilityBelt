using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LyncUtilityBelt
{
	public class WhoIsCallingMe
	{
		private NotifyIcon _icon;
		
		private LocalCallingGuide _localCallingGuide;

		private IncomingCallMonitor _incomingCallMonitor;
		
		public WhoIsCallingMe(NotifyIcon icon)
		{
			_icon = icon;
			_localCallingGuide = new LocalCallingGuide();
			_incomingCallMonitor = new IncomingCallMonitor();
		}

		public void Start()
		{
			_incomingCallMonitor.NewCall += IncomingCallNotifier_NewCall;
		}

		public void Stop()
		{
			_incomingCallMonitor.NewCall -= IncomingCallNotifier_NewCall;
		}

		void IncomingCallNotifier_NewCall(object sender, NewIncomingCallEventArgs e)
		{
			if (e.HasAudioVideo && e.RemoteParticipant.StartsWith("tel:"))
			{
				var ratecenter = _localCallingGuide.Lookup(e.RemoteParticipant);

				var fmt = string.Format("{0} {1}", e.RemoteParticipant, ratecenter);
				_icon.ShowBalloonTip(10000, "Incoming call", fmt, System.Windows.Forms.ToolTipIcon.Info);
			}
			//else
			//	_icon.ShowBalloonTip(20000, "UNKNOWN", e.RemoteParticipant, System.Windows.Forms.ToolTipIcon.Warning);
		}
	}
}
