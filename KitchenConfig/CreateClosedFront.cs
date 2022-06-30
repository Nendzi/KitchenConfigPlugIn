using System;
using Inventor;
using System.Diagnostics;

namespace KitchenConfig
{
    public static class CreateClosedFront
    {
        static string closedFolder;
        static string kantTrakeFolder;
        static string slidersFolder;
        static double sirina;
        static double visina;
        static double pozicija;
        static ComponentOccurrence newOcc;
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;
        static AssemblyDocument CurrentDoc;

        public static void PlaceClosedFront(double SirinaIn, double VisinaIn, double PositionIn, int rn)
        {
            sirina = AuxFunctions.CalculateDoorDim(SirinaIn); // obračunaj stvarnu sirinu i visinu vrata. to je potrebno zbog dužina kant trake
            visina = AuxFunctions.CalculateDoorDim(VisinaIn);
            pozicija = PositionIn;

            Init();

            MakeClosedFront(rn + 1);
            //newOcc.Edit(); forge se buni
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            // Popunjavanje vrata elementima
            PlaceElement("Front", closedFolder);
            PlaceElement("ABS:1", kantTrakeFolder);
            PlaceElement("ABS:2", kantTrakeFolder);
            PlaceElement("ABS:3", kantTrakeFolder);
            PlaceElement("ABS:4", kantTrakeFolder);

            CurrentDoc.Update2(true);
            //newOcc.ExitEdit(ExitTypeEnum.kExitToTop); ako nema edit nema ni exitedit
        }

        private static void Init()
        {
            closedFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Ladice";
            kantTrakeFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Kant trake";
            slidersFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"Ladice\Klizaci";
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
                    posX = sirina / 2 - 0.3;
                    posY = pozicija + visina / 2 - 0.3;
                    posZ = 0;
                    fileName = "GenFrontLadice.ipt";
                    break;

                case "ABS:1":
                    orjX = "-Z";
                    orjY = "-X";
                    orjZ = "Y";
                    posX = 0;
                    posY = pozicija + visina / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:2":
                    orjX = "-Z";
                    orjY = "X";
                    orjZ = "-Y";
                    posX = sirina - 2 * 0.3;
                    posY = pozicija + visina / 2 - 0.3;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:3":
                    orjX = "-Z";
                    orjY = "-Y";
                    orjZ = "-X";
                    posX = sirina / 2 - 0.3;
                    posY = pozicija;
                    posZ = 0;
                    fileName = "ABS.ipt";
                    break;

                case "ABS:4":
                    orjX = "-Z";
                    orjY = "Y";
                    orjZ = "X";
                    posX = sirina / 2 - 0.3;
                    posY = pozicija + visina - 2 * 0.3;
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
                    elemUserParameters["Width"].Value = sirina;
                    elemUserParameters["Heigth"].Value = visina;
                    foreach (HoleFeature item in oCompDef.Features.HoleFeatures)
                    {
                        item.Suppressed = true;
                    }                   
                    fileName = "Closed Front " + AuxFunctions.ConvertCmInMm(sirina) + "x" + AuxFunctions.ConvertCmInMm(visina) + ".ipt";
                    newOcc.Name = name;
                    break;
                case "ABS:1":
                    elemUserParameters["Length"].Value = visina - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "ABS:2":
                    elemUserParameters["Length"].Value = visina - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "ABS:3":
                    elemUserParameters["Length"].Value = sirina - 2 * 0.3;
                    newOcc.Name = name;
                    break;
                case "ABS:4":
                default:
                    elemUserParameters["Length"].Value = sirina - 2 * 0.3;
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

            // s obzirom da fajle već postoji treba ga zameniti
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

            AssemblyDocument oMainDoc = thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, closedFolder + @"\GenFrontLadice.iam", false) as AssemblyDocument;

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
