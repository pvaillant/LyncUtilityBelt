using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LyncUtilityBelt
{
	public class NewIncomingCallEventArgs
	{
		public string RemoteParticipant;
		public bool HasSharingOnly;
		public bool HasInstantMessaging;
		public bool HasAudioVideo;
		public bool IsConference;

		public NewIncomingCallEventArgs(string remoteParticipant, bool hasSharingOnly, bool hasInstantMessaging, bool hasAudioVideo, bool isConference)
		{
			RemoteParticipant = remoteParticipant;
			HasSharingOnly = hasSharingOnly;
			HasInstantMessaging = hasInstantMessaging;
			HasAudioVideo = hasAudioVideo;
			IsConference = isConference;
		}
	}
}
