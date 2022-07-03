using System;
using Inventor;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KitchenConfig
{
    public class CreateCassette
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


        public static void PlaceCassette(double widthIn, double heightIn, double positionIn, int rn)
        {
            width = AuxFunctions.CalculateDoorDim(heightIn);
            height = AuxFunctions.CalculateDoorDim(widthIn);
            position = positionIn;
            if (height < 150)
            {
                hasGlass = false;
            }
            else
            {
                hasGlass = true;
            }

            Init();

            MakeCassetteAssembly(rn + 1);
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            // Popunjavanje vrata elementima
            PlaceElement("CassetteDoor", doorFolder);
            AuxFunctions.PrepareGlassPlace(newOcc, "CassetteDoor", hasGlass);
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
                case "CassetteDoor":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = position;
                    posY = 0;
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
                    posX = position;
                    posY = 0;
                    posZ = 0;
                    fileName = "Glass.ipt";
                    break;
                case "ABS:1":
                    orjX = "-Z";
                    orjY = "-X";
                    orjZ = "Y";
                    posX = position;
                    posY = height / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:2":
                    orjX = "-Z";
                    orjY = "X";
                    orjZ = "-Y";
                    posX = position + width - 2 * 0.3;
                    posY = height / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:3":
                    orjX = "-Z";
                    orjY = "-Y";
                    orjZ = "-X";
                    posX = position + width / 2 - 0.3;
                    posY = 0;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:4":
                    orjX = "-Z";
                    orjY = "Y";
                    orjZ = "X";
                    posX = position + width / 2 - 0.3;
                    posY = height - 2 * 0.3;
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
                case "CassetteDoor":
                    elemUserParameters["Width"].Value = width;
                    elemUserParameters["Height"].Value = height;
                    elemUserParameters["HandlerHolesDistance"].Value = 9;
                    elemUserParameters["HandlerFromTop"].Value = (height - 9) / 2;
                    fileName = "Cassette front " + AuxFunctions.ConvertCmInMm(height) + "x" + AuxFunctions.ConvertCmInMm(width) + ".ipt";
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

        private static void MakeCassetteAssembly(int redniBroj)
        {
            ComponentOccurrences allOccurrs = mainDoc.ComponentDefinition.Occurrences;
            AssemblyDocument oMainDoc = thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, doorFolder + @"\GenDoor.iam", false) as AssemblyDocument;
            AssemblyComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix("Z", "-X", "-Y", height, 1, 0.4)); //"X", "Z", "-Y", 0.4, 1, 0.4
            oMainDoc.Close();
            newOcc.Name = "Hor Door_" + redniBroj.ToString();
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }
    }
}
