using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace LyncUtilityBelt
{
	class LocalCallingGuide
	{
		public string Lookup(string number)
		{
			if (number.StartsWith("tel:"))
				number = number.Substring(4);

			if (number.StartsWith("+"))
				number = number.Substring(1);

			if (number.StartsWith("1") && number.Length == 11)
			{
				var npa = int.Parse(number.Substring(1, 3));
				var nxx = int.Parse(number.Substring(4, 3));
				return LookupNpaNxxRatecenter(npa, nxx);
			}
			else
			{
				return string.Empty;
			}
		}

		// technically this isn't "uControl", but it's the same source, and it really should be in uControl
		private Dictionary<int, string> _npaNxxCache = new Dictionary<int, string>();
		// SEE http://en.wikipedia.org/wiki/Toll-free_telephone_number#North_America and http://nanpa.com/pdf/PL_455.pdf
		private int[] TOLL_FREE_NPAS = new int[] { 800, 888, 877, 866, 855, 844, 833, 822 };
		public string LookupNpaNxxRatecenter(int npa, int nxx)
		{
			if (TOLL_FREE_NPAS.Contains(npa))
				return "Toll-free";

			var npanxx = npa * 1000 + nxx;
			if (!_npaNxxCache.ContainsKey(npanxx))
			{
				const string URL_TEMPLATE = "http://www.localcallingguide.com/xmlprefix.php?npa={0}&nxx={1}";
				var url = string.Format(URL_TEMPLATE, npa, nxx);
				var xml = new WebClient().DownloadString(url);

				string rc = null, region = null;
				Match m;

				m = Regex.Match(xml, "<rc>(.+)</rc>");
				if (m.Success)
					rc = m.Groups[1].Value;
				if (string.IsNullOrEmpty(rc))
					throw new Exception("Invalid ratecenter for " + npa + " " + nxx);

				m = Regex.Match(xml, "<region>(.+)</region>");
				if (m.Success)
					region = m.Groups[1].Value;
				if (string.IsNullOrEmpty(region))
					throw new Exception("Invalid region for " + npa + " " + nxx);

				_npaNxxCache[npanxx] = string.Format("{0}, {1}", rc, region).Replace('é', 'e');
			}
			return _npaNxxCache[npanxx];
		}
	}
}
