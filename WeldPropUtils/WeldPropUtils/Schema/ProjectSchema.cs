using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WeldPropUtils.Schema
{
    // A property that connected parts can have, and the class groups (Pipe, Fittings, ...) that have it.
    public sealed class SourceProperty
    {
        public string Name { get; }
        public List<string> Groups { get; } = new List<string>();

        public SourceProperty(string name)
        {
            Name = name;
        }
    }

    // Reads the class definitions of the piping project (what Project Setup defines) from the project database.
    public static class ProjectSchema
    {
        // Class groups shown as tabs in the mapping window, and the database classes (tables) in each.
        public static readonly IReadOnlyList<KeyValuePair<string, string[]>> SourceGroups = new[]
        {
            new KeyValuePair<string, string[]>("Pipe", new[] { "Pipe" }),
            new KeyValuePair<string, string[]>("Fittings", new[] { "Elbow", "Tee", "Reducer", "Cap", "Coupling", "Cross", "Olet", "Swage", "Nipple", "Lateral", "Wye" }),
            new KeyValuePair<string, string[]>("Flanges", new[] { "Flange", "BlindFlange" }),
            new KeyValuePair<string, string[]>("Valves", new[] { "Valve" }),
            new KeyValuePair<string, string[]>("Nozzle", new[] { "Nozzle" }),
        };

        // Weld classes whose properties are the mapping targets.
        public static readonly string[] WeldTables = { "Buttweld", "Socketweld", "TapWeld", "FusionWeld" };

        // All non-system properties of the connected-part classes, with the groups that have them, sorted by name.
        public static List<SourceProperty> ReadSourceProperties(DataLinksManager dlm)
        {
            var byName = new Dictionary<string, SourceProperty>(StringComparer.Ordinal);
            PnPDatabase db = dlm.GetPnPDatabase();
            foreach (KeyValuePair<string, string[]> group in SourceGroups)
            {
                foreach (string column in ReadColumns(dlm, db, group.Value))
                {
                    if (!byName.TryGetValue(column, out SourceProperty prop))
                        byName[column] = prop = new SourceProperty(column);
                    if (!prop.Groups.Contains(group.Key)) prop.Groups.Add(group.Key);
                }
            }
            return byName.Values.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // All non-system properties of the weld classes (the possible mapping targets), sorted by name.
        public static List<string> ReadWeldProperties(DataLinksManager dlm)
        {
            PnPDatabase db = dlm.GetPnPDatabase();
            return ReadColumns(dlm, db, WeldTables).OrderBy(c => c, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static IEnumerable<string> ReadColumns(DataLinksManager dlm, PnPDatabase db, IEnumerable<string> tables)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            foreach (string table in tables)
            {
                if (!dlm.HasTable(table)) continue;
                foreach (PnPColumn column in db.Tables[table].AllColumns)
                {
                    if (string.IsNullOrEmpty(column.Name) || dlm.IsSystemColumn(table, column.Name)) continue;
                    result.Add(column.Name);
                }
            }
            return result;
        }
    }
}
