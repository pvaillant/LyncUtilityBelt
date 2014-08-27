using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LyncUtilityBelt
{
	// SEE: http://rcosic.wordpress.com/2011/11/17/availability-presence-in-lync-client/
	/*
		Availability
		Invalid (-1),
		None (0)				– Do not use this enumerator. This flag indicates that the cotact state is unspecified.,
		Free (3500)				– A flag indicating that the contact is available,
		FreeIdle (5000)			– Contact is free but inactive,
		Busy (6500)				– A flag indicating that the contact is busy and inactive,
		BusyIdle (7500)			– Contact is busy but inactive,
		DoNotDisturb (9500)		– A flag indicating that the contact does not want to be disturbed,
		TemporarilyAway (12500) – A flag indicating that the contact is temporarily away,
		Away (15500)			– A flag indicating that the contact is away,
		Offline (18500)			– A flag indicating that the contact is signed out.

		Combine with ActivityId to get some other stats like:
		in-a-conference
		in-a-meeting
		out-of-office
		off-work
	 */
	public class AvailabilityMonitor
	{
		public static bool TryGetSelfAvailability(out ContactAvailability availability, out string activityId)
		{
			try
			{
				var info = LyncClient.GetClient().Self.Contact.GetContactInformation(new ContactInformationType[] {
					ContactInformationType.ActivityId,
					ContactInformationType.Availability
				});
				availability = (ContactAvailability)(int)info[ContactInformationType.Availability];
				activityId = (string)info[ContactInformationType.ActivityId];
				return true;
			}
			catch
			{
				availability = ContactAvailability.None;
				activityId = null;
				return false;
			}
		}

		public static void SetSelfAvailability(ContactAvailability availability, string activityId = null)
		{
			if(string.IsNullOrEmpty(activityId))
				activityId = availability.ToString().ToLower();
			try
			{
				var info = new Dictionary<PublishableContactInformationType, object> {
					{PublishableContactInformationType.Availability, availability},
					{PublishableContactInformationType.ActivityId, activityId}
				};
				LyncClient.GetClient().Self.BeginPublishContactInformation(info, null, null);
			}
			catch { }
		}

		internal event Action<ContactAvailability, string> AvailabilityChanged = delegate { };

		private LyncClient _lyncClient;

		internal AvailabilityMonitor()
		{
			// Get a reference to the running Lync client, register for the ConversationAdded event.
			// Note: This assumes that the Lync client is running
			_lyncClient = LyncClient.GetClient();
			_lyncClient.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;
		}

		internal void Initialize()
		{
			OnAvailabilityChanged(_lyncClient.Self.Contact);
		}

		void Contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
		{
			if (e.ChangedContactInformation.Any(x => x == ContactInformationType.Availability))
				OnAvailabilityChanged((Contact)sender);
		}

		private void OnAvailabilityChanged(Contact contact)
		{
			if (_lyncClient.State != ClientState.SignedIn)
				return;

			var info = contact.GetContactInformation(new ContactInformationType[] {
				ContactInformationType.ActivityId,
				ContactInformationType.Availability
			});
			var availability = (ContactAvailability)(int)info[ContactInformationType.Availability];
			var activityId = (string)info[ContactInformationType.ActivityId];
			//Console.WriteLine("AVAILABILITY CHANGED {0} {1}", availability, activityId);
			AvailabilityChanged(availability, activityId);
		}
	}
}
