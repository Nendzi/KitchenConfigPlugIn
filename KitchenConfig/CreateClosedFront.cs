using System;
using Inventor;
using System.Diagnostics;

namespace KitchenConfig
{
    public static class CreateClosedFront
    {
        static string closedFolder;
        static string kantTapeFolder;
        static string slidersFolder;
        static double width;
        static double height;
        static double position;
        static ComponentOccurrence newOcc;
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;
        static AssemblyDocument CurrentDoc;

        public static void PlaceClosedFront(double widthIn, double heightIn, double positionIn, int rn)
        {
            width = AuxFunctions.CalculateDoorDim(widthIn);
            height = AuxFunctions.CalculateDoorDim(heightIn);
            position = positionIn;

            Init();

            MakeClosedFront(rn + 1);
            //newOcc.Edit(); // comment for forge
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            PlaceElement("Front", closedFolder);
            PlaceElement("ABS:1", kantTapeFolder);
            PlaceElement("ABS:2", kantTapeFolder);
            PlaceElement("ABS:3", kantTapeFolder);
            PlaceElement("ABS:4", kantTapeFolder);

            CurrentDoc.Update2(true);
            //newOcc.ExitEdit(ExitTypeEnum.kExitToTop); ako nema edit nema ni exitedit
        }

        private static void Init()
        {
            closedFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Drawers";
            kantTapeFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Kant tapes";
            slidersFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"Drawers\Sliders";
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
                case "Front":
                    orjX = "X";
                    orjY = "-Z";
                    orjZ = "Y";
                    posX = width / 2 - 0.3;
                    posY = position + height / 2 - 0.3;
                    posZ = 0;
                    fileName = "GenDrawerFront.ipt";
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
            }

            ComponentOccurrences allOccurrs = CurrentDoc.ComponentDefinition.Occurrences;

            PartComponentDefinition oCompDef;
            ComponentOccurrence newOcc;
            PartDocument oMainDoc = null;
            oMainDoc = thisApp.Documents.Open(location + @"\" + fileName, false) as PartDocument;
            oCompDef = oMainDoc.ComponentDefinition;
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix(orjX, orjY, orjZ, posX, posY, posZ));

            //resize element
            UserParameters elemUserParameters = oCompDef.Parameters.UserParameters;

            switch (name)
            {
                case "Front":
                    elemUserParameters["Width"].Value = width;
                    elemUserParameters["Heigth"].Value = height;
                    foreach (HoleFeature item in oCompDef.Features.HoleFeatures)
                    {
                        item.Suppressed = true;
                    }                   
                    fileName = "Closed Front " + AuxFunctions.ConvertCmInMm(width) + "x" + AuxFunctions.ConvertCmInMm(height) + ".ipt";
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
                default:
                    elemUserParameters["Length"].Value = width - 2 * 0.3;
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
                //Trace.TraceInformation($"I am saving in Create front {CreateElement.mainFolder + @"\" + fileName}");
            }

            newOcc.Replace(CreateElement.mainFolder + @"\" + fileName, false);
            //Trace.TraceInformation($"I am replacing in Create front {CreateElement.mainFolder + @"\" + fileName}");

            if (oMainDoc != null)
            {
                oMainDoc.ReleaseReference();
                oMainDoc.Close();
            }
        }

        private static void MakeClosedFront(int redniBroj)
        {
            ComponentOccurrences allOccurrs = mainDoc.ComponentDefinition.Occurrences;

            AssemblyDocument oMainDoc = thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, closedFolder + @"\GenDrawerFront.iam", false) as AssemblyDocument;

            AssemblyComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix("X", "Z", "-Y", 0.4, 1, 0.4));

            oMainDoc.Close();

            newOcc.Name = "Closed_" + redniBroj.ToString();
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }
    }
}
