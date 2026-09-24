using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using WeldPropUtils.Settings;

namespace WeldPropUtils
{
    // Weld numbers. Welds of one JointType with the same OD, wall thickness and material on both sides form a group
    // and share one number. Butt welds start at numbering.buttweldStart (11), taps at tapStart (51), socket welds at
    // socketweldStart (71).
    public static class WeldNumbering
    {
        public const string WeldNumberProperty = "WeldNumber";

        private static readonly Regex LeadingNumber = new Regex(@"-?\d+(?:[.,]\d+)?", RegexOptions.Compiled);

        // Size value of a property such as "114.3", "114,3" or "114.3 mm"; 0 when there is no number.
        // Invariant culture, so it works on comma-decimal Windows too.
        public static double ParseSize(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            Match match = LeadingNumber.Match(value);
            if (!match.Success) return 0;
            double.TryParse(match.Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out double result);
            return result;
        }

        // Welds with the same key get the same number. Ports are already normalised (Port1 = larger side).
        public static string GroupKey(Weld weld)
        {
            return weld.WeldType + "|" + PortKey(weld.Port1) + "|" + PortKey(weld.Port2);
        }

        private static string PortKey(structPort port)
        {
            return ParseSize(port.OD).ToString(CultureInfo.InvariantCulture) + "|"
                + ParseSize(port.WallThickness).ToString(CultureInfo.InvariantCulture) + "|" + port.Material;
        }

        // Renumbers all given welds (SetWeldNumber): per type, groups sorted by size (largest first), numbered from the
        // start number. Existing numbers are overwritten.
        public static void NumberAll(List<Weld> welds, NumberingSettings numbering)
        {
            foreach (IGrouping<string, Weld> byType in welds.GroupBy(w => w.WeldType))
            {
                int? start = numbering.StartFor(byType.Key);
                if (start == null) continue;
                int number = start.Value - 1;
                string lastKey = null;
                IEnumerable<Weld> sorted = byType
                    .OrderByDescending(w => ParseSize(w.Port1.OD))
                    .ThenByDescending(w => ParseSize(w.Port1.WallThickness))
                    .ThenByDescending(w => w.Port1.Material, StringComparer.Ordinal)
                    .ThenByDescending(w => ParseSize(w.Port2.OD))
                    .ThenByDescending(w => ParseSize(w.Port2.WallThickness))
                    .ThenByDescending(w => w.Port2.Material, StringComparer.Ordinal);
                foreach (Weld weld in sorted)
                {
                    string key = GroupKey(weld);
                    if (key != lastKey)
                    {
                        number++;
                        lastKey = key;
                    }
                    Write(weld, number.ToString(CultureInfo.InvariantCulture));
                }
            }
        }

        public static void Write(Weld weld, string number)
        {
            weld.WeldNumber = number;
            Acad.dlm.SetProperties(weld.WeldId, new StringCollection { WeldNumberProperty }, new StringCollection { number });
        }
    }

    // Numbers for new welds without touching existing ones: a new weld gets the number of its group,
    // or the next number after the highest one of its type. Built once per drawing from the numbered welds.
    public sealed class WeldNumberIndex
    {
        private readonly Dictionary<string, int> _byKey = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly Dictionary<string, int> _maxByType = new Dictionary<string, int>(StringComparer.Ordinal);

        // 'welds' must have WeldNumber filled (welds without a number are ignored).
        public WeldNumberIndex(IEnumerable<Weld> welds)
        {
            foreach (Weld weld in welds)
            {
                if (!int.TryParse(weld.WeldNumber, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number)) continue;
                Remember(weld, number);
            }
        }

        // Null for weld types without a start number.
        public int? NumberFor(Weld weld, NumberingSettings numbering)
        {
            int? start = numbering.StartFor(weld.WeldType);
            if (start == null) return null;
            if (_byKey.TryGetValue(WeldNumbering.GroupKey(weld), out int existing)) return existing;
            int number = _maxByType.TryGetValue(weld.WeldType, out int max) ? Math.Max(max + 1, start.Value) : start.Value;
            Remember(weld, number);
            return number;
        }

        private void Remember(Weld weld, int number)
        {
            string key = WeldNumbering.GroupKey(weld);
            if (!_byKey.ContainsKey(key)) _byKey[key] = number;
            string type = weld.WeldType ?? "";
            if (!_maxByType.TryGetValue(type, out int max) || number > max) _maxByType[type] = number;
        }
    }
}
