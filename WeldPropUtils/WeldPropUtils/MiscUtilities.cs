//autocad namespaces
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.ProcessPower.DataLinks;
//Plant3d namespaces
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.P3dProjectParts;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using portCol = Autodesk.ProcessPower.PnP3dObjects.PortCollection;
using pPart = Autodesk.ProcessPower.PnP3dObjects.Part;
using pPort = Autodesk.ProcessPower.PnP3dObjects.Port;

namespace WeldPropUtils
{
    // The active drawing and project, resolved on every use (never cached: the user can switch drawings or projects).
    public static class Acad
    {
        public static Document doc => Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Database db => doc.Database;
        public static Editor ed => doc.Editor;
        public static PlantProject currProj => PlantApplication.CurrentProject;
        public static PipingProject pipeProj => currProj.ProjectParts["Piping"] as PipingProject;
        public static DataLinksManager dlm => pipeProj.DataLinksManager;
    }

    public static class MiscUtilities
    {
        // Non-empty properties of a project row. Throws on failure; WeldPropertiesHandler.ProcessWelds skips that weld
        // and reports all failures once (no MessageBox per failure).
        public static Dictionary<string, string> GetP3dProps(this int rowId)
        {
            Dictionary<string, string> propsDict = new Dictionary<string, string>();
            List<KeyValuePair<string, string>> props = Acad.dlm.GetAllProperties(rowId, true);
            foreach (KeyValuePair<string, string> prop in props)
            {
                if (!string.IsNullOrEmpty(prop.Key) && !string.IsNullOrEmpty(prop.Value))
                {
                    propsDict[prop.Key] = prop.Value;
                }
            }
            return propsDict;
        }

        public static bool IsWeld(this Dictionary<string, string> props)
        {
            return props.TryGetValue("JointType", out string jointType) &&
                (jointType == "Buttweld" || jointType == "Tap"
                || jointType == "Socketweld");
        }

        public static Weld AsWeld(this Connector conn, Transaction tr)
        {
            ConnectionManager connMgr = new ConnectionManager();

            portCol ports = conn.GetPorts(PortType.Both);
            Weld weld = new Weld("", conn.FindWeldRowId(), "", new structPort(), new structPort());
            int portIndex = 1;

            foreach (pPort port in ports)
            {
                if (connMgr.IsConnected(new Pair { ObjectId = conn.ObjectId, Port = port }))
                {
                    pPart connPart = conn.GetConnPart(tr, port);
                    if (connPart != null)
                    {
                        var portProps = port.GetPortProps(tr, connPart);
                        weld.PortPropCollect(portIndex, portProps);
                    }
                }
                portIndex++;
            }
            weld.NormalizePorts();
            return weld;
        }

        private static structPort GetPortProps(this pPort port, Transaction tr, pPart connPart)
        {
            int connPartRowID = Acad.dlm.FindAcPpRowId(connPart.ObjectId);
            Dictionary<string, string> connPartProps = connPartRowID.GetP3dProps();
            string portName = null;
            foreach (pPort p in connPart.GetPorts(PortType.Both))
            {
                if (p.Position == port.Position)
                {
                    portName = p.Name;
                    break;
                }
            }

            //if the pipe is connected to nozzle of the equipment, we will retrieve properties of equipment
            //insteaad of nozzle. That is why we need to dig into equipment and find our nozzle
            //actually connectors usually have name S1, S2, but for nozzles it is longer, so we could use it
            if (portName?.Length > 2)
            {
                Equipment eqp = tr.GetObject(connPart.ObjectId, OpenMode.ForRead) as Equipment;
                NozzleSubPartCollection nozzleCol = eqp.AllSubParts;
                foreach (NozzleSubPart n in nozzleCol)
                {
                    int nId = Acad.dlm.FindAcPpRowId(Acad.dlm.MakeAcPpObjectId(connPart.ObjectId, 1));
                    Dictionary<string, string> nozzleProps = nId.GetP3dProps();
                    if (nozzleProps.FirstOrDefault(p => p.Key == "PortName").Value == portName)
                    {
                        connPartProps = nozzleProps;
                        break;
                    }
                }
            }
            return new structPort
            {
                Material = connPartProps.FirstOrDefault(p => p.Key == "Material").Value,
                Lds = connPartProps.FirstOrDefault(p => p.Key == "PartSizeLongDesc").Value,
                Spec = connPartProps.FirstOrDefault(p => p.Key == "Spec").Value,
                WallThickness = connPartProps.FirstOrDefault(p => p.Key == "WallThickness").Value,
                OD = connPartProps.FirstOrDefault(p => p.Key == "MatchingPipeOd").Value,
                Props = connPartProps
            };
        }

        private static pPart GetConnPart(this Connector conn, Transaction tr, pPort port)
        {
            ConnectionIterator connIter = ConnectionIterator.NewIterator(conn.ObjectId, port);
            while (!connIter.Done())
            {
                ObjectId connectedPartdId = connIter.ObjectId;
                if (connectedPartdId != ObjectId.Null)
                {
                    pPart connPart = tr.GetObject(connectedPartdId, OpenMode.ForRead) as pPart;
                    if (connPart != null && !(connPart is Connector))
                    {
                        return connPart;
                    }
                }
                connIter.Next();
            }
            return null;
        }

        public static int FindWeldRowId(this Connector conn)
        {
            SubPartCollection subPartCol = conn.AllSubParts;

            foreach (SubPart subPart in subPartCol)
            {
                if (subPart is WeldSubPart)
                    return Acad.dlm.FindAcPpRowId(Acad.dlm.MakeAcPpObjectId(conn.ObjectId, 1));
            }
            return 0;
        }
    }
}
