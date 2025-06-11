using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WeldPropUtils
{
    public struct structPort
    {
        public string OD { get; set; }
        public string WallThickness { get; set; }
        public string Material { get; set; }
        public string Spec { get; set; }
        public string Lds { get; set; }

        public structPort(string od, string wallThickness,string material, string spec, string lds)
        {
            OD = od;
            WallThickness = wallThickness;
            Material = material;
            Spec = spec;
            Lds = lds;
        }
    }

    public class Weld
    {
        public string WeldNumber { get; set; }
        public int WeldId { get; set; }
        public string WeldType { get; set; }
        public structPort Port1 { get; set; }
        public structPort Port2 { get; set; }

        public Weld(string weldNumber, int weldId, string weldType, structPort port1, structPort port2)
        {
            WeldNumber = weldNumber;
            WeldId = weldId;
            WeldType = weldType;
            Port1 = port1;
            Port2 = port2;
        }


        //we need to have larger port ever first
        private int ComparePorts(structPort port1, structPort port2)
        {
            int odComparison = string.Compare(port1.OD, port2.OD);

            if (odComparison != 0) return odComparison;
            int wallThicknessComparison = string.Compare(port1.WallThickness, port2.WallThickness);
            if (wallThicknessComparison != 0) return wallThicknessComparison;

            return string.Compare(port1.Material, port2.Material);
        }

        public void NormalizePorts()
        {
            if (ComparePorts(Port1, Port2) < 0)
            {
                SwapPorts();
            }
        }

        private void SwapPorts()
        {
            (Port1,Port2) = (Port2,Port1);
        }

        public void PortPropCollect (int portIndex, structPort port)
        {
            if (portIndex == 1)
            {
                Port1 = port;
            }
            else if (portIndex == 2)
            {
                Port2 = port;
            }
            else
            {
                throw new ArgumentException("Port index must be 1 or 2");
            }
        }
    }
}
