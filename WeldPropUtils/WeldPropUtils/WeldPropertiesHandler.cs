using Autodesk.AutoCAD.ApplicationServices; 
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.PnP3dObjects;

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;

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

        private static void SetWeldProp(Connector conn, Weld weld)
        {
            StringCollection pNames = new StringCollection
            {
                "Material1", "OD1", "WallThickness1", "LDS1", "SPEC1",
                "Material2", "OD2", "WallThickness2", "LDS2", "SPEC2"
            };
            StringCollection pVals = new StringCollection
            {
                weld.Port1.Material, weld.Port1.OD, weld.Port1.WallThickness, weld.Port1.Lds, weld.Port1.Spec,
                weld.Port2.Material, weld.Port2.OD, weld.Port2.WallThickness, weld.Port2.Lds, weld.Port2.Spec
            };
            int subPartRowID = conn.FindWeldRowId();
            Acad.dlm.SetProperties(subPartRowID, pNames, pVals);
        }
        [CommandMethod ("SetWeldProp")]
        public static void SetWeldProp()
        {
            LoopThroughWelds((connector, Weld) => { SetWeldProp(connector, Weld); });
        }
        [CommandMethod ("SetWeldNumber")]
        public static void WeldNumerAssign()
        {
            int bw = 11;
            int tw = 51;
            int sw = 71;

            LoopThroughWelds(SetWeldProp, (weldList) =>
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
