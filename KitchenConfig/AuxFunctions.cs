using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace KitchenConfig
{
    static public class AuxFunctions
    {
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;

        public static string GetWorkingDir(AssemblyDocument mainDoc)
        {
            return thisApp.DesignProjectManager.ActiveDesignProject.WorkspacePath;
            //return System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            //return System.IO.Path.GetDirectoryName(mainDoc.FullFileName); // used in forge implementation
        }

        public static Matrix CreateMatrix(string axisX, string axisY, string axisZ, double posX, double posY, double posZ)
        {
            TransientGeometry oTG = thisApp.TransientGeometry;

            Matrix oMatrix = oTG.CreateMatrix();

            Vector oAxisX1 = oTG.CreateVector(1, 0, 0);
            Vector oAxisY1 = oTG.CreateVector(0, 1, 0);
            Vector oAxisZ1 = oTG.CreateVector(0, 0, 1);
            Point oPoint1 = oTG.CreatePoint(0, 0, 0);

            Vector oAxisX2 = ConvertInVector(axisX);
            Vector oAxisY2 = ConvertInVector(axisY);
            Vector oAxisZ2 = ConvertInVector(axisZ);
            Point oPoint2 = oTG.CreatePoint(posX, posY, posZ);

            oMatrix.SetToAlignCoordinateSystems(oPoint1, oAxisX1, oAxisY1, oAxisZ1, oPoint2, oAxisX2, oAxisY2, oAxisZ2);
            return oMatrix.Copy();
        }

        private static Vector ConvertInVector(string Direction)
        {
            TransientGeometry oTG = thisApp.TransientGeometry;

            switch (Direction)
            {
                case "-X":
                    return oTG.CreateVector(-1, 0, 0);
                case "X":
                    return oTG.CreateVector(1, 0, 0);
                case "-Y":
                    return oTG.CreateVector(0, -1, 0);
                case "Y":
                    return oTG.CreateVector(0, 1, 0);
                case "-Z":
                    return oTG.CreateVector(0, 0, -1);
                case "Z":
                default:
                    return oTG.CreateVector(0, 0, 1);
            }
        }

        public static void PrepareGlassPlace(ComponentOccurrence activeOcc, string doorName, bool hasGlass)
        {
            foreach (ComponentOccurrence doorOcc in activeOcc.SubOccurrences)
            {
                if (doorOcc.Name == doorName)
                {
                    PartComponentDefinition doorDef = (PartComponentDefinition)doorOcc.Definition;
                    ExtrudeFeatures doorFeats = doorDef.Features.ExtrudeFeatures;
                    if (hasGlass)
                    {
                        doorFeats["Glass"].Suppressed = false;
                        doorFeats["Opening"].Suppressed = false;
                    }
                    else
                    {
                        doorFeats["Glass"].Suppressed = true;
                        doorFeats["Opening"].Suppressed = true;
                    }
                    continue;
                }
            }
        }

        public static double CalculateDoorDim(double dimIn)
        {
            return dimIn - 2 * 0.1; //0.1 = kant tape thickness
        }

        public static bool IsFileExist(string fileName)
        {
            return System.IO.File.Exists(fileName);
        }

        public static void MakePattern(ComponentOccurrence oOcc, double width, double depth)
        {
            // Set a reference to the AssemblyComponentDefinition
            AssemblyComponentDefinition oDef = mainDoc.ComponentDefinition;

            // Get the x\-axis of the assembly
            WorkAxis oXAxis = oDef.WorkAxes["X Axis"];

            // Get the y\-axis of the assembly
            WorkAxis oYAxis = oDef.WorkAxes["Y Axis"];

            ObjectCollection oParentOccs = thisApp.TransientObjects.CreateObjectCollection();
            oParentOccs.Add(oOcc);

            // Create a rectangular pattern of components\:
            // 2 columns in the x\-direction with an offset of sirina in
            // 2 rows in the y\-direction with an offset of dubina in
            oDef.OccurrencePatterns.AddRectangularPattern(oParentOccs, oXAxis, true, width, 2, oYAxis, true, depth, 2);
        }

        public static string ExtractName(string fileName)
        {
            return fileName.Substring(0, fileName.Length - 4);
        }

        public static string MakeDimensions(double dim1, double dim2)
        {
            string str1 = Math.Round(dim1 * 10, 0).ToString();
            string str2 = Math.Round(dim2 * 10, 0).ToString();

            return str1 + "x" + str2;
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }

        public static IvyxType ConvertStrInEnum(string input)
        {
            switch (input)
            {
                case "closed":
                    return IvyxType.closed;
                case "drawer":
                    return IvyxType.drawer;
                case "leftDoor":
                    return IvyxType.leftDoor;
                case "rightDoor":
                    return IvyxType.rightDoor;
                case "doubleDoor":
                    return IvyxType.doubleDoor;
                case "open":
                    return IvyxType.open;
                case "cassette":
                    return IvyxType.cassette;
                default:
                    return IvyxType.open;
            }
        }

        public static double CalculateIvyxHeight(double percent)
        {
            return CreateElement.realHeigth * percent / 100;
        }

        public static string ConvertCmInMm(double length)
        {
            double ABSLen;
            ABSLen = Math.Round(length * 10, 0);
            return ABSLen.ToString();
        }
    }
}
