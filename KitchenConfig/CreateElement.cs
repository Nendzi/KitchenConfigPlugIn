using System;
using System.Collections.Generic;
using System.Diagnostics;
using Inventor;

namespace KitchenConfig
{
    public static class CreateElement
    {
        static AssemblyComponentDefinition MainAssy;
        public static string mainFolder;
        static string drawerFolder;
        static string sinkFolder;
        static string workingTopFolder;
        static string legsFolder;
        static string sidesFolder;
        static double width;
        static double height;
        static double position = 0;
        public static double realHeigth;
        static double depth = 60;
        static string ivyxHeigth;
        static string ivyxType;
        static string[] ivyxTypeName;
        static double[] ivyxHeightInPercent;
        static int sinkType;
        //static int handleType;
        //static bool hasShelf;
        //static int shelfQty = 2;

        static InventorServer thisApp;
        public static AssemblyDocument oActiveDoc;

        public static void Composer(AssemblyDocument oDoc, InventorServer oApp)
        {
            oActiveDoc = oDoc;
            thisApp = oApp;
            CreateDoor.Inject(oDoc, oApp);
            CreateMirrorDoor.Inject(oDoc, oApp);
            CreateCassette.Inject(oDoc, oApp);
            CreateClosedFront.Inject(oDoc, oApp);
            CreateDrawer.Inject(oDoc, oApp);
            AuxFunctions.Inject(oDoc, oApp);

            Init();

            // Creating corpus for kitchnen element
            PlaceElement("Bottom panel", sidesFolder);
            PlaceElement("Right panel", sidesFolder);
            PlaceElement("Left panel", sidesFolder);
            PlaceElement("Thin back plate", sidesFolder);
            PlaceElement("Legs", legsFolder);
            switch (sinkType)
            {
                case -1:
                    PlaceElement("Top element", workingTopFolder);
                    break;
                case 0:
                    PlaceSink(sinkType, sinkFolder);
                    break;
                case 1:
                    PlaceSink(sinkType, sinkFolder);
                    break;
                case 2:
                    PlaceSink(sinkType, sinkFolder);
                    break;
            }
            // Insert dividers
            List<double> distances = DividersDistribution();
            var rb = 1;
            foreach (var distance in distances)
            {
                PlaceElement("Divider", sidesFolder, distance, rb);
                rb++;
            }

            height = 0;
            for (int i = 0; i < ivyxTypeName.Length; i++)
            {
                height = AuxFunctions.CalculateIvyxHeight(ivyxHeightInPercent[i]);
                switch (AuxFunctions.ConvertStrInEnum(ivyxTypeName[i]))
                {
                    case IvyxType.closed:
                        CreateClosedFront.PlaceClosedFront(width, height, position, i);
                        break;
                    case IvyxType.drawer:
                        CreateDrawer.PlaceDrawer(width, height, position, i);
                        break;
                    case IvyxType.leftDoor:
                        CreateDoor.PlaceDoor(width, height, position, i);
                        break;
                    case IvyxType.rightDoor:
                        CreateMirrorDoor.PlaceDoor(width, height, position, i, 1);
                        break;
                    case IvyxType.doubleDoor:
                        CreateDoor.PlaceDoor(width / 2, height, position, i);
                        CreateMirrorDoor.PlaceDoor(width, height, position, i, 0.5);
                        break;
                    case IvyxType.cassette:
                        CreateCassette.PlaceCassette(width, height, position, i);
                        break;
                    case IvyxType.open:
                    default:
                        break;
                }
                position += AuxFunctions.CalculateIvyxHeight(ivyxHeightInPercent[i]);
            }

            string NameForDisk;
            AssemblyDocument topAssy = oDoc;

            foreach (ComponentOccurrence doorOcc in topAssy.ComponentDefinition.Occurrences)
            {
                if (doorOcc.Definition.Type == ObjectTypeEnum.kAssemblyComponentDefinitionObject)
                {
                    NameForDisk = doorOcc.Name;
                    MainAssy = (AssemblyComponentDefinition)doorOcc.Definition;
                    MainAssy.Document.SaveAs(mainFolder + @"\" + NameForDisk + ".iam", false);
                }
            }
            oDoc.Update2(true);
            topAssy.SaveAs(mainFolder + @"\Main element.iam", false);
        }

        private static void PlaceElement(string name, string location, double distance = 0, int rb = 0)
        {
            string fileName = "";
            string orjX = "";
            string orjY = "";
            string orjZ = "";
            double posY = 0;
            double posZ = 0;
            double posX = 0;

            //Load element
            switch (name)
            {
                case "Bottom panel":
                    orjX = "Z";
                    orjY = "X";
                    orjZ = "Y";
                    posX = 1.8;
                    posY = 2.8;
                    posZ = 0;
                    fileName = "Gen Side and Panel.ipt";
                    break;
                case "Right panel":
                    orjX = "-X";
                    orjY = "Z";
                    orjZ = "Y";
                    posX = width;
                    posY = 2.8;
                    posZ = 0;
                    fileName = "Gen Side and Panel.ipt";
                    break;
                case "Left panel":
                    orjX = "X";
                    orjY = "-Z";
                    orjZ = "Y";
                    posX = 0;
                    posY = 2.8;
                    posZ = height - 3.8;
                    fileName = "Gen Side and Panel.ipt";
                    break;
                case "Thin back plate":
                    orjX = "X";
                    orjY = "-Y";
                    orjZ = "-Z";
                    posX = width / 2;
                    posY = depth;
                    posZ = (height - 3.8) / 2;
                    fileName = "Gen Back plate.ipt";
                    break;
                case "Top element":
                    orjX = "X";
                    orjY = "Z";
                    orjZ = "-Y";
                    posX = width / 2;
                    posY = depth;
                    posZ = height - 3.8;
                    fileName = "Gen Top plate.ipt";
                    break;
                case "Legs":
                    orjX = "-Y";
                    orjY = "Z";
                    orjZ = "-X";
                    posX = 4;
                    posY = 4 + 2.8;
                    posZ = -10;
                    fileName = "Leg01.ipt";
                    break;
                case "Divider":
                    orjX = "Z";
                    orjY = "X";
                    orjZ = "Y";
                    posX = 1.8;
                    posY = 2.8;
                    posZ = distance;
                    fileName = "Gen Side and Panel.ipt";
                    break;
            }

            ComponentOccurrences allOccurrs = oActiveDoc.ComponentDefinition.Occurrences;

            string fullNameOfFile = System.IO.Path.Combine(location, fileName);
            PartDocument oMainDoc = thisApp.Documents.Open(fullNameOfFile, false) as PartDocument;

            PartComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            ComponentOccurrence newOcc = allOccurrs.AddByComponentDefinition(oCompDef as ComponentDefinition, AuxFunctions.CreateMatrix(orjX, orjY, orjZ, posX, posY, posZ));

            //resize element
            UserParameters elemUserParameters = oCompDef.Parameters.UserParameters;

            switch (name)
            {
                case "Bottom panel":
                case "Divider":
                    elemUserParameters["Height"].Value = width - 2 * 1.8;
                    elemUserParameters["Depth"].Value = depth - 0.4 - 2.8;
                    if (rb == 0)
                    {
                        newOcc.Name = name;
                    }
                    else
                    {
                        newOcc.Name = name + rb.ToString();
                    }
                    fileName = "Bottom " + AuxFunctions.MakeDimensions(width, depth) + ".ipt";
                    break;
                case "Right panel":
                    elemUserParameters["Height"].Value = height - 3.8;
                    elemUserParameters["Depth"].Value = depth - 0.4 - 2.8;
                    newOcc.Name = name;
                    fileName = "Lateral " + AuxFunctions.MakeDimensions(height, depth) + ".ipt";
                    break;
                case "Left panel":
                    elemUserParameters["Height"].Value = height - 3.8;
                    elemUserParameters["Depth"].Value = depth - 0.4 - 2.8;
                    newOcc.Name = name;
                    fileName = "Lateral " + AuxFunctions.MakeDimensions(height, depth) + ".ipt";
                    break;
                case "Thin back plate":
                    elemUserParameters["Height"].Value = height - 3.8;
                    elemUserParameters["Width"].Value = width;
                    newOcc.Name = name;
                    break;
                case "Top element":
                    elemUserParameters["Width"].Value = width;
                    newOcc.Name = name;
                    break;
                case "Legs":
                    newOcc.Name = name;
                    AuxFunctions.MakePattern(newOcc, width - 2 * 4, depth - 2.8 - 2 * 4);
                    break;
            }

            newOcc.Definition.Document.Update2(true);

            string fileNameOnDisk = mainFolder + @"\" + fileName;
            if (!AuxFunctions.IsFileExist(fileNameOnDisk))
            {
                oCompDef.Document.SaveAs(fileNameOnDisk, true);
            }

            newOcc.Replace(fileNameOnDisk, false);

            oMainDoc.ReleaseReference();
            oMainDoc.Close();
        }

        private static void PlaceSink(int sinkType, string location)
        {
            switch (sinkType)
            {
                case 0:
                    break;
                case 1:
                    break;
                case 2:
                    break;
            }
        }

        private static void Init()
        {
            sinkFolder = AuxFunctions.GetWorkingDir(oActiveDoc) + @"\Sink\BuiltOn";
            workingTopFolder = AuxFunctions.GetWorkingDir(oActiveDoc) + @"\Top";
            legsFolder = AuxFunctions.GetWorkingDir(oActiveDoc) + @"\Legs";
            sidesFolder = AuxFunctions.GetWorkingDir(oActiveDoc) + @"\Sides and Shelfs";
            drawerFolder = AuxFunctions.GetWorkingDir(oActiveDoc) + @"\Drawers";

            mainFolder = AuxFunctions.GetWorkingDir(oActiveDoc) + @"\Result";

            width = oActiveDoc.ComponentDefinition.Parameters.UserParameters["width"].Value; //mm
            height = oActiveDoc.ComponentDefinition.Parameters.UserParameters["height"].Value; //mm
            realHeigth = height - 3.8;
            ivyxType = oActiveDoc.ComponentDefinition.Parameters.UserParameters["ivyxType"].Value;
            ivyxHeigth = oActiveDoc.ComponentDefinition.Parameters.UserParameters["ivyxHeigth"].Value;

            ivyxTypeName = ParseIvyx(ivyxType);
            ivyxHeightInPercent = ParseHeigth(ivyxHeigth);

            sinkType = -1; //-1, 0, 1, 2

            CreateDoor.hasGlass = true;
            CreateMirrorDoor.hasGlass = true;

            //handleType = 1;// -1, 0 , 1

            //hasShelf = true;
            //shelfQty = 1;
        }

        private static string[] ParseIvyx(string inputData)
        {
            string[] output = inputData.Split(',');
            return output;
        }
        private static double[] ParseHeigth(string inputData)
        {
            string[] vs = inputData.Split(',');
            double[] output = new double[vs.Length];

            for (int i = 0; i < vs.Length; i++)
            {
                output[i] = Convert.ToDouble(vs[i]);
            }

            return output;
        }
        private static List<double> DividersDistribution()
        {
            List<double> output = new List<double>();
            double[] offsets = new double[ivyxTypeName.Length - 1];
            string ivyxName;
            string prevIvyxName;
            double distance;

            for (int i = 1; i < ivyxHeightInPercent.Length; i++)
            {
                ivyxName = ivyxTypeName[i];
                prevIvyxName = ivyxTypeName[i - 1];
                if (ivyxName.ToLower().Contains("door"))
                {
                    if (prevIvyxName == "closed" || prevIvyxName == "open" || prevIvyxName == "drawer")
                    {
                        offsets[i - 1] = 1;
                    }
                    else if (prevIvyxName == "cassette" || prevIvyxName.ToLower().Contains("door"))
                    {
                        offsets[i - 1] = 0.5;
                    }
                }
                else if (ivyxName == "cassette")
                {
                    offsets[i - 1] = 1;
                }
                else if (ivyxName == "open")
                {
                    offsets[i - 1] = 0;
                }
                else if (ivyxName == "closed")
                {
                    offsets[i - 1] = -1;
                }
                else
                {
                    if (prevIvyxName == "cassette" || prevIvyxName == "closed" || prevIvyxName == "drawer")
                    {
                        offsets[i - 1] = -1;
                    }
                    else if (prevIvyxName == "open")
                    {
                        offsets[i - 1] = 1;
                    }
                    else { offsets[i - 1] = 0; }
                }
            }

            double elevation = 0;
            double position;
            for (int i = 0; i < offsets.Length; i++)
            {
                elevation += AuxFunctions.CalculateIvyxHeight(ivyxHeightInPercent[i]);
                position = 1.8 * offsets[i];
                distance = elevation + position;
                if (position < 0) continue;
                output.Add(distance - 1.8);
            }

            return output;
        }
    }
}
