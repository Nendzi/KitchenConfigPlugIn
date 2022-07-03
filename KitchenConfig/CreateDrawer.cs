using System;
using Inventor;
using System.Diagnostics;

namespace KitchenConfig
{
    public static class CreateDrawer
    {
        static string drawerFolder;
        static string rukohvatiFolder;
        static string kantTrakeFolder;
        static string sliderFolder;
        //static int handleType;
        static double width;
        static double height;
        static double position;
        static double depth = 60; //cm
        static ComponentOccurrence newOcc;
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;
        static AssemblyDocument CurrentDoc;


        public static void PlaceDrawer(double widthIn, double heightIn, double positionIn, int rn)
        {
            width = AuxFunctions.CalculateDoorDim(widthIn);
            height = AuxFunctions.CalculateDoorDim(heightIn);
            position = positionIn;

            Init();

            MakeDrawerAssembly(rn + 1);
            //newOcc.Edit(); //dont use in forge
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            PlaceElement("DrawerFront", drawerFolder);
            PlaceElement("Handler", rukohvatiFolder);
            PlaceElement("ABS:1", kantTrakeFolder);
            PlaceElement("ABS:2", kantTrakeFolder);
            PlaceElement("ABS:3", kantTrakeFolder);
            PlaceElement("ABS:4", kantTrakeFolder);
            PlaceElement("DrawerSide:1", drawerFolder);
            PlaceElement("DrawerSide:2", drawerFolder);
            PlaceElement("DrawerBottom", drawerFolder);
            PlaceElement("DrawerBackwall", drawerFolder);

            CurrentDoc.Update2(true);
            //newOcc.ExitEdit(ExitTypeEnum.kExitToTop); ako nema edit nema ni exitedit
        }

        private static void Init()
        {
            drawerFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Drawers";
            rukohvatiFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Handlers";
            kantTrakeFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Kant tapes";
            sliderFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Drawers\Sliders";
        }

        private static void PlaceElement(string name, string location)
        {
            //Trace.TraceInformation($"You are placing {name} on {location}");
            string fileName = "";
            string orjX = "";
            string orjY = "";
            string orjZ = "";
            double posX = 0;
            double posY = 0;
            double posZ = 0;

            //Load element
            switch (name)
            {
                case "DrawerFront":
                    orjX = "X";
                    orjY = "-Z";
                    orjZ = "Y";
                    posX = width / 2 - 0.3;
                    posY = position + height / 2 - 0.3;
                    posZ = 0;
                    fileName = "GenDrawerFront.ipt";
                    break;
                case "Handler":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = 0;
                    posZ = 0;
                    fileName = "Handler01.ipt";
                    break;
                case "DrawerBackwall":
                    orjX = "-Z";
                    orjY = "-X";
                    orjZ = "Y";
                    posX = width / 2;
                    posY = position + 10 / 2 + 1.8 - 0.5;
                    posZ = -depth + 4.2;
                    fileName = "GenDrawerSide.ipt";
                    break;
                case "ABS:1":
                    orjX = "-Z";
                    orjY = "-X";
                    orjZ = "Y";
                    posX = 0;
                    posY = position + height / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:2":
                    orjX = "-Z";
                    orjY = "X";
                    orjZ = "-Y";
                    posX = width - 2 * 0.3;
                    posY = position + height / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:3":
                    orjX = "-Z";
                    orjY = "-Y";
                    orjZ = "-X";
                    posX = width / 2 - 0.3;
                    posY = position;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:4":
                    orjX = "-Z";
                    orjY = "Y";
                    orjZ = "X";
                    posX = width / 2 - 0.3;
                    posY = position + height - 2 * 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;
                case "DrawerSide:1":
                    orjX = "X";
                    orjY = "-Z";
                    orjZ = "Y";
                    posX = width - 4.6;
                    posY = position + 10 / 2 + 1.8 - 0.5;
                    posZ = -(depth / 2 - 0.3);
                    fileName = "GenDrawerSide.ipt";
                    break;
                case "DrawerSide:2":
                    orjX = "-X";
                    orjY = "Z";
                    orjZ = "Y";
                    posX = 4.6;
                    posY = position + 10 / 2 + 1.8 - 0.5;
                    posZ = -(depth / 2 - 0.3);
                    fileName = "GenDrawerSide.ipt";
                    break;
                case "DrawerBottom":
                default:
                    orjX = "X";
                    orjY = "-Z";
                    orjZ = "-Y";
                    posX = width / 2;
                    posY = position + 1.8 + 0.3;
                    posZ = -(depth / 2) + 1.2;
                    fileName = "DrawerBottom.ipt";
                    break;
            }
            //TODO - CurrentDoc je ActiveEditDocument
            ComponentOccurrences allOccurrs = CurrentDoc.ComponentDefinition.Occurrences;

            PartComponentDefinition oCompDef;
            ComponentOccurrence newOcc;
            PartDocument oMainDoc = null;

            if (fileName == "Handler01.ipt")
            {
                try
                {
                    newOcc = (ComponentOccurrence)allOccurrs.AddUsingiMates(location + @"\" + fileName, false);
                }
                catch (Exception)
                {

                }

                //Occurrence already exist in assembly so get it by name
                newOcc = allOccurrs.ItemByName[AuxFunctions.ExtractName(fileName) + ":1"];
                oCompDef = (PartComponentDefinition)newOcc.Definition;
            }
            else
            {
                oMainDoc = thisApp.Documents.Open(location + @"\" + fileName, false) as PartDocument;

                oCompDef = oMainDoc.ComponentDefinition;
                newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix(orjX, orjY, orjZ, posX, posY, posZ));
            }

            //resize element
            UserParameters elemUserParameters = oCompDef.Parameters.UserParameters;

            switch (name)
            {
                case "DrawerFront":
                    elemUserParameters["Width"].Value = width;
                    elemUserParameters["Heigth"].Value = height;
                    elemUserParameters["DistanceForHandle"].Value = 9;
                    fileName = "Drawer Front " + AuxFunctions.ConvertCmInMm(width) + "x" + AuxFunctions.ConvertCmInMm(height) + ".ipt";
                    newOcc.Name = name;
                    break;
                case "Handler":
                    newOcc.Name = name;
                    break;
                case "DrawerBackwall":
                    elemUserParameters["Height"].Value = 10;
                    elemUserParameters["Length"].Value = width - 9.2;
                    fileName = "Drawer Backwall " + AuxFunctions.ConvertCmInMm(width) + ".ipt";
                    newOcc.Name = name;
                    break;
                case "DrawerBottom":
                    elemUserParameters["Depth"].Value = depth;
                    elemUserParameters["Width"].Value = width;
                    fileName = "Drawer Bottom " + AuxFunctions.ConvertCmInMm(width) + "x" + AuxFunctions.ConvertCmInMm(depth) + ".ipt";
                    newOcc.Name = name;
                    break;
                case "ABS:1":
                    elemUserParameters["Length"].Value = height - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "ABS:2":
                    elemUserParameters["Length"].Value = height - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "ABS:3":
                    elemUserParameters["Length"].Value = width - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "ABS:4":
                    elemUserParameters["Length"].Value = width - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "DrawerSide:1":
                case "DrawerSide:2":
                default:
                    elemUserParameters["Length"].Value = depth - 4.2;
                    elemUserParameters["Height"].Value = 10;
                    fileName = "Drawer Lateralwall " + AuxFunctions.ConvertCmInMm(depth) + ".ipt";
                    newOcc.Name = name;
                    break;
            }

            if (fileName == "ABS.ipt")
            {                
                fileName = "ABS " + AuxFunctions.ConvertCmInMm(elemUserParameters["Length"].Value) + " mm.ipt";
            }

            newOcc.Definition.Document.Update2(true);

            if (!AuxFunctions.IsFileExist(CreateElement.mainFolder + @"\" + fileName))
            {
                oCompDef.Document.SaveAs(CreateElement.mainFolder + @"\" + fileName, true);
                //Trace.TraceInformation($"I am saving in Create door {CreateElement.mainFolder + @"\" + fileName}");
            }

            // s obzirom da fajle već postoji treba ga zameniti
            newOcc.Replace(CreateElement.mainFolder + @"\" + fileName, false);
            //Trace.TraceInformation($"I am replacing in Create door {CreateElement.mainFolder + @"\" + fileName}");

            if (oMainDoc != null)
            {
                oMainDoc.ReleaseReference();
                oMainDoc.Close();
            }
        }

        private static void MakeDrawerAssembly(int redniBroj)
        {
            ComponentOccurrences allOccurrs = mainDoc.ComponentDefinition.Occurrences;
            AssemblyDocument oMainDoc = thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, drawerFolder + @"\GenDrawer.iam", false) as AssemblyDocument;
            AssemblyComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix("X", "Z", "-Y", 0.4, 1, 0.4));
            oMainDoc.Close();
            newOcc.Name = "Drawer_" + redniBroj.ToString();
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }

    }
}
