using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace LyncUtilityBelt
{
	public interface ILyncHueConfig
	{
		string AppKey { get; set; }
		string BridgeIP { get; set; }
		string LightID { get; set; }

		void Save();
	}

	[XmlRoot]
	public class LyncUtilityBeltConfig : ILyncHueConfig
	{
		private static readonly string FILE_NAME = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LyncUtilityBelt.xml");
		private const string KEY_CHARS = "abcdefghkmnprstuvwxyzABCDEFGHKLMNPRSTUVWXYZ123456789";
		private const int KEY_LENGTH = 20;

		public static LyncUtilityBeltConfig Load()
		{
			if (File.Exists(FILE_NAME))
			{
				var ser = new XmlSerializer(typeof(LyncUtilityBeltConfig));
				using (var file = new FileStream(FILE_NAME, FileMode.Open))
					return (LyncUtilityBeltConfig)ser.Deserialize(file);
			}
			else
			{
				var r = new Random();
				var sb = new System.Text.StringBuilder();
				for (var i = 0; i < KEY_LENGTH; i++)
					sb.Append(KEY_CHARS[r.Next(KEY_CHARS.Length)]);

				return new LyncUtilityBeltConfig { AppKey = sb.ToString() };
			}
		}

		[XmlIgnore]
		public bool Dirty { get; private set; }

		#region Outlook Work Hours config
		private bool _outlookWorkHoursEnabled;
		[XmlElement]
		public bool OutlookWorkHoursEnabled
		{
			get { return _outlookWorkHoursEnabled; }
			set { _outlookWorkHoursEnabled = value; Dirty = true; }
		}
		#endregion

		#region Who Is Calling Me config
		private bool _whoIsCallingMeEnabled;
		[XmlElement]
		public bool WhoIsCallingMeEnabled
		{
			get { return _whoIsCallingMeEnabled; }
			set { _whoIsCallingMeEnabled = value; Dirty = true; }
		}
		#endregion

		#region Lync Hue config
		private string _lyncHueAppKey;
		[XmlElement("LyncHueAppKey")]
		public string AppKey
		{
			get { return _lyncHueAppKey; }
			set { _lyncHueAppKey = value; Dirty = true; }
		}

		private string _lyncHueBridgeIP;
		[XmlElement("LyncHueBridgeIP")]
		public string BridgeIP
		{
			get { return _lyncHueBridgeIP; }
			set { _lyncHueBridgeIP = value; Dirty = true; }
		}

		private string _lyncHueLightID;
		[XmlElement("LyncHueLightID")]
		public string LightID
		{
			get { return _lyncHueLightID; }
			set { _lyncHueLightID = value; Dirty = true; }
		}
		#endregion

		public void Save()
		{
			if (Dirty)
			{
				var ser = new XmlSerializer(typeof(LyncUtilityBeltConfig));
				using (var file = File.Open(FILE_NAME, (File.Exists(FILE_NAME) ? FileMode.Truncate : FileMode.CreateNew)))
					ser.Serialize(file, this);
				Dirty = false;
			}
		}
	}
}
