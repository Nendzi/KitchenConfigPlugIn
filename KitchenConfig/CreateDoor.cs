using System;
using Inventor;
using System.Diagnostics;

namespace KitchenConfig
{
    public static class CreateDoor
    {
        static string doorFolder;
        static string handlerFolder;
        static string kantTapeFolder;
        static string glassFolder;
        public static bool hasGlass;
        static double width;
        static double height;
        static double position;
        static ComponentOccurrence newOcc;
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;
        static AssemblyDocument CurrentDoc;


        public static void PlaceDoor(double widthIn, double heightIn, double positionIn, int rn)
        {
            width = AuxFunctions.CalculateDoorDim(widthIn); // obračunaj stvarnu sirinu i visinu vrata. to je potrebno zbog dužina kant trake
            height = AuxFunctions.CalculateDoorDim(heightIn);
            position = positionIn;

            Init();

            MakeDoorAssembly(rn + 1);
            //newOcc.Edit(); // dont use in forge
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            PlaceElement("Gen Door", doorFolder);
            AuxFunctions.PrepareGlassPlace(newOcc, "Gen Door", hasGlass);
            PlaceElement("Handler", handlerFolder);
            PlaceElement("ABS:1", kantTapeFolder);
            PlaceElement("ABS:2", kantTapeFolder);
            PlaceElement("ABS:3", kantTapeFolder);
            PlaceElement("ABS:4", kantTapeFolder);
            if (hasGlass)
            {
                PlaceElement("Glass", glassFolder);
            }

            CurrentDoc.Update2(true);
            //newOcc.ExitEdit(ExitTypeEnum.kExitToTop); ako nema edit nema ni exitedit
        }

        private static void Init()
        {
            doorFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Doors";
            handlerFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Handlers";
            kantTapeFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Kant tapes";
            glassFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Doors";
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
                case "Gen Door":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = position;
                    posZ = -1.8;
                    fileName = "Gen Door.ipt";
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
                case "Glass":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = position;
                    posZ = 0;
                    fileName = "Glass.ipt";
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
            //TODO - CurrentDoc je ActiveEditDocument
            ComponentOccurrences allOccurrs = CurrentDoc.ComponentDefinition.Occurrences;

            PartComponentDefinition oCompDef;
            ComponentOccurrence newOcc;
            PartDocument oMainDoc = null;

            if (fileName == "Glass.ipt" || fileName == "Handler01.ipt")
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
                case "Gen Door":
                    elemUserParameters["Width"].Value = width;
                    elemUserParameters["Height"].Value = height;
                    elemUserParameters["HandlerHolesDistance"].Value = 9;
                    fileName = "Left door Front " + AuxFunctions.ConvertCmInMm(width) + "x" + AuxFunctions.ConvertCmInMm(height) + ".ipt";
                    newOcc.Name = name;
                    break;
                case "Handler":
                    newOcc.Name = name;
                    break;
                case "Glass":
                    elemUserParameters["Height"].Value = height - 0.6 - 2 * 7;
                    elemUserParameters["Width"].Value = width - 0.6 - 2 * 7;
                    fileName = "Glass " + AuxFunctions.ConvertCmInMm(width) + "x" + AuxFunctions.ConvertCmInMm(height) + ".ipt";
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
                //Trace.TraceInformation($"I am saving in Create door {CreateElement.mainFolder + @"\" + fileName}");
            }

            newOcc.Replace(CreateElement.mainFolder + @"\" + fileName, false);
            //Trace.TraceInformation($"I am replacing in Create door {CreateElement.mainFolder + @"\" + fileName}");

            if (oMainDoc != null)
            {
                oMainDoc.ReleaseReference();
                oMainDoc.Close();
            }
        }

        private static void MakeDoorAssembly(int serialNumber)
        {
            ComponentOccurrences allOccurrs = mainDoc.ComponentDefinition.Occurrences;

            AssemblyDocument oMainDoc = thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, doorFolder + @"\GenDoor.iam", false) as AssemblyDocument;

            AssemblyComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix("X", "Z", "-Y", 0.4, 1, 0.4));

            oMainDoc.Close();

            newOcc.Name = "Left Door_" + serialNumber.ToString();
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }

    }
}
