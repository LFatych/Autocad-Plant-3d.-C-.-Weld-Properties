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
            ObjectId[] objIds = selRes.Value.GetObjectIds();
            List<Weld> welds = new List<Weld>();

            using (DocumentLock docLock = Acad.doc.LockDocument()) ;
            using (Transaction tr = Acad.db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId objId in objIds)
                {
                    Connector connector = tr.GetObject(objId, OpenMode.ForRead) as Connector;
                    if (connector is null) continue;
                    Dictionary<string, string> connectorProps = Acad.dlm.FindAcPpRowId(connector.ObjectId).GetP3dProps();
                    if (!connectorProps.IsWeld()) continue;

                    Weld weld = connector.AsWeld(tr);
                    SetProp?.Invoke(connector, weld);
                    if (WeldNumAssign != null)
                    {
                        weld.WeldType = connectorProps["JointType"];
                        welds.Add(weld);
                    }
                }
                WeldNumAssign?.Invoke(welds);
                tr.Commit();
            }

        }

        // Writes each mapped weld property from the connected part on its side; a property missing on the part is written as null.
        private static void SetWeldProp(Connector conn, Weld weld, MappingProfile profile)
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
                weldTargets = ProjectSchema.ReadWeldTargets(Acad.dlm);
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
            if (!window.Saved) return;

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

            if (window.UpdateWeldsRequested)
            {
                MappingProfile profile = settings.GetActiveProfile();
                LoopThroughWelds((connector, weld) => { SetWeldProp(connector, weld, profile); });
                Acad.ed.WriteMessage("\nWeld properties updated with profile \"" + profile.Name + "\".");
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

        [CommandMethod ("SetWeldProp")]
        public static void SetWeldProp()
        {
            WeldPropSettings settings = LoadSettings();
            if (settings == null) return;
            MappingProfile profile = settings.GetActiveProfile();
            LoopThroughWelds((connector, weld) => { SetWeldProp(connector, weld, profile); });
        }
        [CommandMethod ("SetWeldNumber")]
        public static void WeldNumerAssign()
        {
            WeldPropSettings settings = LoadSettings();
            if (settings == null) return;
            MappingProfile profile = settings.GetActiveProfile();
            int bw = settings.Numbering.ButtweldStart;
            int tw = settings.Numbering.TapStart;
            int sw = settings.Numbering.SocketweldStart;

            LoopThroughWelds((connector, weld) => { SetWeldProp(connector, weld, profile); }, (weldList) =>
            {
                List<Weld> buttWelds = weldList.Where(w => w.WeldType == "Buttweld").ToList();
                List<Weld> tapWelds = weldList.Where(w => w.WeldType == "Tap").ToList();
                List<Weld> socketWelds = weldList.Where(w => w.WeldType == "Socketweld").ToList();

                buttWelds?.SetNum(bw);
                tapWelds?.SetNum(tw);
                socketWelds?.SetNum(sw);
            });
        }
    }
}
