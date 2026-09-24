using Autodesk.AutoCAD.ApplicationServices; 
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.PnP3dObjects;

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;

using Autodesk.ProcessPower.PlantInstance;

using WeldPropUtils.Schema;
using WeldPropUtils.Settings;
using WeldPropUtils.UI;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace WeldPropUtils
{
    public class WeldPropertiesHandler
    {
        public static void LoopThroughWelds(
            Action<Connector, Weld> SetProp = null, Action<List<Weld>> WeldNumAssign = null)
        {
            SelectionFilter selFilter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "ACPPCONNECTOR"),
                new TypedValue((int)DxfCode.Visibility, 0)
            });
            PromptSelectionResult selRes = Acad.ed.SelectAll(selFilter);
            if (selRes.Status != PromptStatus.OK)
            {
                Acad.ed.WriteMessage("\nNo connectors found.");
                return;
            }
            ProcessWelds(selRes.Value.GetObjectIds(), SetProp, WeldNumAssign);
        }

        // Runs SetProp for every weld (Buttweld/Tap/Socketweld connector) among 'objIds' of the active drawing, then
        // WeldNumAssign with all of them. A connector that fails (e.g. not linked to the project) is skipped and counted.
        // The caller must hold the document lock (a Modal command has it; WeldAutoUpdate locks explicitly).
        // Returns the number of welds processed.
        internal static int ProcessWelds(
            IEnumerable<ObjectId> objIds, Action<Connector, Weld> SetProp = null, Action<List<Weld>> WeldNumAssign = null)
        {
            List<Weld> welds = new List<Weld>();
            int failed = 0;
            string firstError = null;

            using (Transaction tr = Acad.db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId objId in objIds)
                {
                    try
                    {
                        Connector connector = tr.GetObject(objId, OpenMode.ForRead) as Connector;
                        if (connector is null) continue;
                        Dictionary<string, string> connectorProps = Acad.dlm.FindAcPpRowId(connector.ObjectId).GetP3dProps();
                        if (!connectorProps.IsWeld()) continue;

                        Weld weld = connector.AsWeld(tr);
                        SetProp?.Invoke(connector, weld);
                        weld.WeldType = connectorProps["JointType"];
                        welds.Add(weld);
                    }
                    catch (System.Exception ex)
                    {
                        failed++;
                        if (firstError == null) firstError = ex.Message;
                    }
                }
                WeldNumAssign?.Invoke(welds);
                tr.Commit();
            }
            if (failed > 0)
                Acad.ed.WriteMessage("\n" + failed + " connector(s) skipped because of errors (first: " + firstError + ").");
            return welds.Count;
        }

        // Writes each mapped weld property from the connected part on its side; a property missing on the part is written as null.
        internal static void SetWeldProp(Connector conn, Weld weld, MappingProfile profile)
        {
            StringCollection pNames = new StringCollection();
            StringCollection pVals = new StringCollection();
            foreach (PropertyMapping mapping in profile.Mappings)
            {
                Dictionary<string, string> partProps = mapping.Side == 1 ? weld.Port1.Props : weld.Port2.Props;
                string value = null;
                if (partProps != null) partProps.TryGetValue(mapping.Source, out value);
                pNames.Add(mapping.Target);
                pVals.Add(value);
            }
            if (pNames.Count == 0) return;
            int subPartRowID = conn.FindWeldRowId();
            Acad.dlm.SetProperties(subPartRowID, pNames, pVals);
        }

        // The profile without mappings to weld properties that the weld classes in Project Setup don't have
        // (writing one of those would fail for every weld). Skipped targets are reported once.
        internal static MappingProfile WritableProfile(MappingProfile profile)
        {
            HashSet<string> weldProps;
            try
            {
                weldProps = new HashSet<string>(ProjectSchema.ReadWeldProperties(Acad.dlm), StringComparer.Ordinal);
            }
            catch (System.Exception)
            {
                return profile;
            }
            if (weldProps.Count == 0) return profile;
            List<string> missing = profile.Mappings.Where(m => !weldProps.Contains(m.Target)).Select(m => m.Target).ToList();
            if (missing.Count == 0) return profile;
            Acad.ed.WriteMessage("\nSkipped weld properties that are not in Project Setup: " + string.Join(", ", missing));
            MappingProfile writable = profile.Clone(profile.Name);
            writable.Mappings.RemoveAll(m => !weldProps.Contains(m.Target));
            return writable;
        }

        // Null (with a message) when no Plant project is open.
        private static WeldPropSettings LoadSettings()
        {
            string path = SettingsStore.GetSettingsPath();
            if (path == null)
            {
                Acad.ed.WriteMessage("\nNo Plant 3D project is open.");
                return null;
            }
            string message;
            WeldPropSettings settings = SettingsStore.Load(path, out message);
            if (message != null) Acad.ed.WriteMessage("\n" + message);
            return settings;
        }

        // Opens the Weld Property Mapping window for the current project's settings.
        [CommandMethod ("WeldPropMapping")]
        public static void OpenMappingWindow()
        {
            string path = SettingsStore.GetSettingsPath();
            WeldPropSettings settings = LoadSettings();
            if (settings == null) return;

            List<SourceProperty> sources;
            List<string> weldTargets;
            try
            {
                sources = ProjectSchema.ReadSourceProperties(Acad.dlm);
                weldTargets = ProjectSchema.ReadWeldProperties(Acad.dlm);
            }
            catch (System.Exception ex)
            {
                Acad.ed.WriteMessage("\nCould not read the classes from Project Setup: " + ex.Message);
                sources = new List<SourceProperty>();
                weldTargets = new List<string>();
            }

            UserPreferences prefs = UserPreferences.Load();
            bool dark = prefs.Theme == UserPreferences.ThemeDark
                || (prefs.Theme != UserPreferences.ThemeLight && AutoCadUsesDarkTheme());
            var viewModel = new MappingViewModel(settings, sources, weldTargets, PlantApplication.CurrentProject.Name);
            var window = new MappingWindow(viewModel, dark, prefs);
            AcApp.ShowModalWindow(window);
            if (window.Result == MappingWindowResult.Cancel) return;

            try
            {
                SettingsStore.Save(settings, path);
                Acad.ed.WriteMessage("\nSaved " + path);
            }
            catch (System.Exception ex)
            {
                Acad.ed.WriteMessage("\nCould not save " + path + ": " + ex.Message);
                return;
            }

            if (window.Result == MappingWindowResult.SaveAndUpdate)
            {
                MappingProfile profile = WritableProfile(settings.GetActiveProfile());
                LoopThroughWelds((connector, weld) => { SetWeldProp(connector, weld, profile); });
                Acad.ed.WriteMessage("\nWeld properties updated with profile \"" + profile.Name + "\".");
            }
            else if (window.Result == MappingWindowResult.SaveAndNumber)
            {
                NumberAllWelds(settings);
            }
        }

        // AutoCAD's COLORTHEME: 0 = dark, 1 = light.
        private static bool AutoCadUsesDarkTheme()
        {
            try
            {
                return Convert.ToInt32(AcApp.GetSystemVariable("COLORTHEME")) == 0;
            }
            catch (System.Exception)
            {
                return true;
            }
        }

        // Switches "autoUpdate" of the project settings: new welds get their mapped properties when the command that
        // created them ends (see WeldAutoUpdate).
        [CommandMethod ("WeldPropAuto")]
        public static void ToggleAutoUpdate()
        {
            WeldPropSettings settings = LoadSettings();
            if (settings == null) return;
            string path = SettingsStore.GetSettingsPath();
            settings.AutoUpdate = !settings.AutoUpdate;
            try
            {
                SettingsStore.Save(settings, path);
            }
            catch (System.Exception ex)
            {
                Acad.ed.WriteMessage("\nCould not save " + path + ": " + ex.Message);
                return;
            }
            Acad.ed.WriteMessage(settings.AutoUpdate
                ? "\nWeld auto-update is ON for this project: new welds get their properties (profile \"" + settings.GetActiveProfile().Name + "\")."
                : "\nWeld auto-update is OFF for this project. Use SetWeldProp to update welds.");
        }

        [CommandMethod ("SetWeldProp")]
        public static void SetWeldProp()
        {
            WeldPropSettings settings = LoadSettings();
            if (settings == null) return;
            MappingProfile profile = WritableProfile(settings.GetActiveProfile());
            LoopThroughWelds((connector, weld) => { SetWeldProp(connector, weld, profile); });
        }
        [CommandMethod ("SetWeldNumber")]
        public static void WeldNumerAssign()
        {
            WeldPropSettings settings = LoadSettings();
            if (settings == null) return;
            NumberAllWelds(settings);
        }

        // Fills the mapped properties of all welds, then renumbers them all (WeldNumbering.NumberAll).
        private static void NumberAllWelds(WeldPropSettings settings)
        {
            MappingProfile profile = WritableProfile(settings.GetActiveProfile());
            List<Weld> numbered = null;
            LoopThroughWelds((connector, weld) => { SetWeldProp(connector, weld, profile); }, (weldList) =>
            {
                WeldNumbering.NumberAll(weldList, settings.Numbering);
                numbered = weldList;
            });
            if (numbered == null) return;
            WeldAutoUpdate.SetNumberIndex(Acad.db, new WeldNumberIndex(numbered));
            Acad.ed.WriteMessage("\n" + numbered.Count + " weld(s) numbered.");
        }
    }
}
