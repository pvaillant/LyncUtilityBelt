using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LyncUtilityBelt
{
	public class IncomingCallMonitor
	{
		//public Action<string, bool, bool, bool, bool> OnNewCall = delegate { };
		internal event EventHandler<NewIncomingCallEventArgs> NewCall = delegate { };
		void OnNewCall(string remoteParticipant, bool hasSharingOnly, bool hasInstantMessaging, bool hasAudioVideo, bool isConference)
		{
			NewCall(this, new NewIncomingCallEventArgs(remoteParticipant, hasSharingOnly, hasInstantMessaging, hasAudioVideo, isConference));
		}

		private LyncClient _lyncClient = null;

		internal IncomingCallMonitor()
		{
			// Get a reference to the running Lync client, register for the ConversationAdded event.
			// Note: This assumes that the Lync client is running
			_lyncClient = LyncClient.GetClient();
			_lyncClient.ConversationManager.ConversationAdded += ConversationManager_ConversationAdded;
		}

		void ConversationManager_ConversationAdded(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
		{
			var conversation = e.Conversation;

			// Test conversation state. If inactive, then the new conversation window was opened by the user, not a remote participant
			if (conversation.State == ConversationState.Inactive)
				return;

			// Get the URI of the "Inviter" contact
			var remoteParticipant = ((Contact)conversation.Properties[ConversationProperty.Inviter]).Uri;

			// Determine which modalities are available in the conversation
			bool hasSharingOnly = true;

			bool hasInstantMessaging = false;
			if (ModalityIsNotified(conversation, ModalityTypes.InstantMessage))
			{
				hasInstantMessaging = true;
				hasSharingOnly = false;
			}

			bool hasAudioVideo = false;
			if (ModalityIsNotified(conversation, ModalityTypes.AudioVideo))
			{
				hasAudioVideo = true;
				hasSharingOnly = false;
			}

			// Get whether this is a conference
			bool isConference = conversation.Properties[ConversationProperty.ConferencingUri] != null;

			// Raise the NewCall event
			OnNewCall(remoteParticipant, hasSharingOnly, hasInstantMessaging, hasAudioVideo, isConference);
		}

		private bool ModalityIsNotified(Conversation conversation, ModalityTypes modalityType)
		{
			return conversation.Modalities.ContainsKey(modalityType) &&
				   conversation.Modalities[modalityType].State == ModalityState.Notified;
		}
	}
}
