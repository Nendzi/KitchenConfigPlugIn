using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace KitchenConfig
{
    static public class CreateMirrorDoor
    {
        static string doorFolder;
        static string handlerFolder;
        static string kantTapesFolder;
        static string glassFolder;
        public static bool hasGlass;
        //static int handleType;
        //static bool hasShelf;
        //static int shelfQty;
        static double width;
        static double elementWidth;
        static double height;
        static double position;
        static ComponentOccurrence newOcc;
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;
        static AssemblyDocument CurrentDoc;

        public static void PlaceDoor(double widthIn, double heightIn, double positionIn, int rn, double ratio)
        {
            elementWidth = widthIn;
            width = AuxFunctions.CalculateDoorDim(widthIn * ratio);
            height = AuxFunctions.CalculateDoorDim(heightIn);
            position = positionIn;

            Init();

            MakeDoorAssembly(rn + 1);
            //newOcc.Edit();
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            // Popunjavanje vrata elementima
            PlaceElement("Gen Door_MIR", doorFolder);
            AuxFunctions.PrepareGlassPlace(newOcc, "Gen Door_MIR", hasGlass);
            PlaceElement("Handler", handlerFolder);
            PlaceElement("ABS:1", kantTapesFolder);
            PlaceElement("ABS:2", kantTapesFolder);
            PlaceElement("ABS:3", kantTapesFolder);
            PlaceElement("ABS:4", kantTapesFolder);
            if (hasGlass)
            {
                PlaceElement("Glass", glassFolder);
            }

            CurrentDoc.Update2(true);
            //newOcc.ExitEdit(ExitTypeEnum.kExitToTop); isto kao i door
        }

        private static void Init()
        {
            doorFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Doors";
            handlerFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Handlers";
            kantTapesFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Kant tapes";
            glassFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Doors";
        }

        private static void PlaceElement(string name, string location)
        {
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
                case "Gen Door_MIR":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = position;
                    posZ = 0;
                    fileName = "Gen Door_MIR.ipt";
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
                    orjX = "Z";
                    orjY = "X";
                    orjZ = "Y";
                    posX = 0;
                    posY = position + height / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:2":
                    orjX = "Z";
                    orjY = "-X";
                    orjZ = "-Y";
                    posX = -(width - 2 * 0.3);
                    posY = position + height / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:3":
                    orjX = "Z";
                    orjY = "-Y";
                    orjZ = "X";
                    posX = -(width / 2 - 0.3);
                    posY = position;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:4":
                    orjX = "Z";
                    orjY = "Y";
                    orjZ = "X";
                    posX = -(width / 2 - 0.3);
                    posY = position + height - 2 * 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;
            }

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
                case "Gen Door_MIR":
                    elemUserParameters["Width"].Value = width;
                    elemUserParameters["Height"].Value = height;
                    elemUserParameters["HandlerHolesDistance"].Value = 9;
                    fileName = "Right door Front " + AuxFunctions.ConvertCmInMm(width) + "x" + AuxFunctions.ConvertCmInMm(height) + ".ipt";
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
                //Trace.TraceInformation($"I am saving in Create Mirror door {CreateElement.mainFolder + @"\" + fileName}");
            }

            newOcc.Replace(CreateElement.mainFolder + @"\" + fileName, false);
            //Trace.TraceInformation($"I am replacing in Create Mirror door {CreateElement.mainFolder + @"\" + fileName}");

            if (oMainDoc != null)
            {
                oMainDoc.ReleaseReference();
                oMainDoc.Close();
            }
        }

        private static void MakeDoorAssembly(int serialNumber)
        {
            ComponentOccurrences allOccurrs = mainDoc.ComponentDefinition.Occurrences;

            AssemblyDocument oMainDoc =
                thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, doorFolder + @"\GenDoor_MIR.iam", false) as AssemblyDocument;

            AssemblyComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix("X", "Z", "-Y", elementWidth-0.4, 2.8, 0.4));

            oMainDoc.Close();

            newOcc.Name = "Right Door_" + serialNumber.ToString();
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }
    }
}
