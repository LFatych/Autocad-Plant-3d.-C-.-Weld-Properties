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

        public void Normalize()
        {
            if (string.IsNullOrEmpty(Name)) Name = DefaultName;
            if (Mappings == null) Mappings = new List<PropertyMapping>();
            Mappings.RemoveAll(m => m == null || !m.IsValid);
            // One source per target weld property: the last entry wins.
            Mappings = Mappings.GroupBy(m => m.Target).Select(g => g.Last()).ToList();
        }
    }

    [DataContract]
    public class PropertyMapping
    {
        // Weld property written (e.g. "Material1").
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

        public static NumberingSettings CreateDefault()
        {
            return new NumberingSettings { ButtweldStart = 11, TapStart = 51, SocketweldStart = 71 };
        }
    }
}
