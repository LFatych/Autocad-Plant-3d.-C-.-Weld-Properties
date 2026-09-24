using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace WeldPropUtils.Settings
{
    // Per-project settings, stored as JSON in <Plant project folder>\SPDS\WeldPropUtils.json.
    // DataContract serialization skips constructors, so defaults for missing members are applied in Normalize().
    [DataContract]
    public class WeldPropSettings
    {
        public const int CurrentSchemaVersion = 1;

        [DataMember(Name = "schemaVersion", Order = 0)]
        public int SchemaVersion { get; set; }

        [DataMember(Name = "activeProfile", Order = 1)]
        public string ActiveProfile { get; set; }

        [DataMember(Name = "autoUpdate", Order = 2)]
        public bool AutoUpdate { get; set; }

        [DataMember(Name = "numbering", Order = 3)]
        public NumberingSettings Numbering { get; set; }

        [DataMember(Name = "profiles", Order = 4)]
        public List<MappingProfile> Profiles { get; set; }

        public static WeldPropSettings CreateDefault()
        {
            return new WeldPropSettings
            {
                SchemaVersion = CurrentSchemaVersion,
                ActiveProfile = MappingProfile.DefaultName,
                AutoUpdate = false,
                Numbering = NumberingSettings.CreateDefault(),
                Profiles = new List<MappingProfile> { MappingProfile.CreateDefault() }
            };
        }

        public void Normalize()
        {
            if (SchemaVersion == 0) SchemaVersion = CurrentSchemaVersion;
            if (Numbering == null) Numbering = NumberingSettings.CreateDefault();
            if (Profiles == null) Profiles = new List<MappingProfile>();
            Profiles.RemoveAll(p => p == null);
            foreach (MappingProfile profile in Profiles) profile.Normalize();
            if (Profiles.Count == 0) Profiles.Add(MappingProfile.CreateDefault());
            if (string.IsNullOrEmpty(ActiveProfile)) ActiveProfile = Profiles[0].Name;
        }

        // Falls back to the first profile when the active name does not exist.
        public MappingProfile GetActiveProfile()
        {
            return Profiles.FirstOrDefault(p => p.Name == ActiveProfile) ?? Profiles.FirstOrDefault();
        }
    }

    [DataContract]
    public class MappingProfile
    {
        public const string DefaultName = "Default";

        [DataMember(Name = "name", Order = 0)]
        public string Name { get; set; }

        // UI hint: mapping one side also maps the other side's counterpart.
        [DataMember(Name = "mirrorSides", Order = 1)]
        public bool MirrorSides { get; set; }

        [DataMember(Name = "mappings", Order = 2)]
        public List<PropertyMapping> Mappings { get; set; }

        // Same result as the original hard-coded SetWeldProp.
        public static MappingProfile CreateDefault()
        {
            var mappings = new List<PropertyMapping>();
            foreach (int side in new[] { 1, 2 })
            {
                mappings.Add(new PropertyMapping("Material" + side, side, "Material"));
                mappings.Add(new PropertyMapping("OD" + side, side, "MatchingPipeOd"));
                mappings.Add(new PropertyMapping("WallThickness" + side, side, "WallThickness"));
                mappings.Add(new PropertyMapping("LDS" + side, side, "PartSizeLongDesc"));
                mappings.Add(new PropertyMapping("SPEC" + side, side, "Spec"));
            }
            return new MappingProfile { Name = DefaultName, MirrorSides = true, Mappings = mappings };
        }

        public PropertyMapping Find(string target)
        {
            return Mappings.FirstOrDefault(m => m.Target == target);
        }

        public string GetSource(string target)
        {
            PropertyMapping mapping = Find(target);
            return mapping == null ? null : mapping.Source;
        }

        // Maps weld property 'target' to part property 'source' of the given side (1 or 2).
        public void Assign(string target, int side, string source)
        {
            Mappings.RemoveAll(m => m.Target == target);
            Mappings.Add(new PropertyMapping(target, side == 2 ? 2 : 1, source));
        }

        public void Unassign(string target)
        {
            Mappings.RemoveAll(m => m.Target == target);
        }

        public MappingProfile Clone(string newName)
        {
            return new MappingProfile
            {
                Name = newName,
                MirrorSides = MirrorSides,
                Mappings = Mappings.Select(m => new PropertyMapping(m.Target, m.Side, m.Source)).ToList()
            };
        }

        public void Normalize()
        {
            if (string.IsNullOrEmpty(Name)) Name = DefaultName;
            if (Mappings == null) Mappings = new List<PropertyMapping>();
            Mappings.RemoveAll(m => m == null || !m.IsValid);
            // One source per target weld property: the last entry wins.
            Mappings = Mappings.GroupBy(m => m.Target).Select(g => g.Last()).ToList();
        }
    }

    // Pairs weld properties of side 1 and side 2 by name: the same name with one '1'/'2' digit swapped
    // ("Material1" <-> "Material2", "Port1_Material" <-> "Port2_Material").
    public static class WeldSides
    {
        // The counterpart that exists in 'known' (the weld properties from Project Setup); null when there is none.
        // Without 'known' only a trailing digit is swapped.
        public static string Counterpart(string target, ICollection<string> known)
        {
            if (string.IsNullOrEmpty(target)) return null;
            if (known == null || known.Count == 0)
            {
                char last = target[target.Length - 1];
                return last == '1' || last == '2' ? target.Substring(0, target.Length - 1) + Swap(last) : null;
            }
            for (int i = 0; i < target.Length; i++)
            {
                char ch = target[i];
                if (ch != '1' && ch != '2') continue;
                string other = target.Substring(0, i) + Swap(ch) + target.Substring(i + 1);
                if (known.Contains(other)) return other;
            }
            return null;
        }

        private static char Swap(char digit)
        {
            return digit == '1' ? '2' : '1';
        }
    }

    [DataContract]
    public class PropertyMapping
    {
        // Weld property written (e.g. "Material1", "Port1_Material"); any property of the weld class.
        [DataMember(Name = "target", Order = 0)]
        public string Target { get; set; }

        // 1 = larger connected part (Weld.Port1), 2 = the other part (Weld.Port2).
        [DataMember(Name = "side", Order = 1)]
        public int Side { get; set; }

        // Property read from the connected part (e.g. "Material").
        [DataMember(Name = "source", Order = 2)]
        public string Source { get; set; }

        public PropertyMapping() { }

        public PropertyMapping(string target, int side, string source)
        {
            Target = target;
            Side = side;
            Source = source;
        }

        public bool IsValid
        {
            get { return !string.IsNullOrEmpty(Target) && !string.IsNullOrEmpty(Source) && (Side == 1 || Side == 2); }
        }
    }

    [DataContract]
    public class NumberingSettings
    {
        [DataMember(Name = "buttweldStart", Order = 0)]
        public int ButtweldStart { get; set; }

        [DataMember(Name = "tapStart", Order = 1)]
        public int TapStart { get; set; }

        [DataMember(Name = "socketweldStart", Order = 2)]
        public int SocketweldStart { get; set; }

        // New welds get a number automatically (WeldAutoUpdate): the number of their group, or the next free one.
        [DataMember(Name = "auto", Order = 3)]
        public bool Auto { get; set; }

        // Start number for a weld JointType (Buttweld/Tap/Socketweld); null for other types.
        public int? StartFor(string weldType)
        {
            switch (weldType)
            {
                case "Buttweld": return ButtweldStart;
                case "Tap": return TapStart;
                case "Socketweld": return SocketweldStart;
                default: return null;
            }
        }

        public static NumberingSettings CreateDefault()
        {
            return new NumberingSettings { ButtweldStart = 11, TapStart = 51, SocketweldStart = 71 };
        }
    }
}
