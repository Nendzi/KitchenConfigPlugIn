using System;
using Inventor;
using System.Diagnostics;

namespace KitchenConfig
{
    public static class CreateDoor
    {
        static string vrataFolder;
        static string rukohvatiFolder;
        static string kantTrakeFolder;
        static string stakloFolder;
        public static bool hasGlass;
        static double sirina;
        static double visina;
        static double pozicija;
        static ComponentOccurrence newOcc;
        static InventorServer thisApp;
        static AssemblyDocument mainDoc;
        static AssemblyDocument CurrentDoc;


        public static void PlaceDoor(double SirinaIn, double VisinaIn, double PositionIn,  int rn)
        {
            sirina = AuxFunctions.CalculateDoorDim(SirinaIn); // obračunaj stvarnu sirinu i visinu vrata. to je potrebno zbog dužina kant trake
            visina = AuxFunctions.CalculateDoorDim(VisinaIn);
            pozicija = PositionIn;

            Init();
            
            MakeDoorAssembly(rn+1);
            //newOcc.Edit(); forge se buni
            CurrentDoc = newOcc.Definition.Document as AssemblyDocument;

            // Popunjavanje vrata elementima
            PlaceElement("Gen vrata", vrataFolder);
            AuxFunctions.PrepareGlassPlace(newOcc, "Gen vrata", hasGlass);
            PlaceElement("Rukohvat", rukohvatiFolder);
            PlaceElement("ABS:1", kantTrakeFolder);
            PlaceElement("ABS:2", kantTrakeFolder);
            PlaceElement("ABS:3", kantTrakeFolder);
            PlaceElement("ABS:4", kantTrakeFolder);
            if (hasGlass)
            {
                PlaceElement("Staklo", stakloFolder);
            }

            CurrentDoc.Update2(true);
            //newOcc.ExitEdit(ExitTypeEnum.kExitToTop); ako nema edit nema ni exitedit
        }

        private static void Init()
        {
            vrataFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Vrata";
            rukohvatiFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Rukohvati";
            kantTrakeFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Kant trake";
            stakloFolder = AuxFunctions.GetWorkingDir(mainDoc) + @"\Vrata";
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
                case "Gen vrata":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = pozicija;
                    posZ = -1.8;
                    fileName = "Gen Vrata.ipt";
                    break;
                case "Rukohvat":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = 0;
                    posZ = 0;
                    fileName = "Rukohvat01.ipt";
                    break;
                case "Staklo":
                    orjX = "X";
                    orjY = "Y";
                    orjZ = "Z";
                    posX = 0;
                    posY = pozicija;
                    posZ = 0;
                    fileName = "Staklo.ipt";
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
            //TODO - CurrentDoc je ActiveEditDocument
            ComponentOccurrences allOccurrs = CurrentDoc.ComponentDefinition.Occurrences;

            PartComponentDefinition oCompDef;
            ComponentOccurrence newOcc;
            PartDocument oMainDoc = null;

            if (fileName == "Staklo.ipt" || fileName == "Rukohvat01.ipt")
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
                case "Gen vrata":
                    elemUserParameters["Sirina"].Value = sirina;
                    elemUserParameters["Visina"].Value = visina;
                    elemUserParameters["RazmakRukohvata"].Value = 9;
                    fileName = "Left door Front " + AuxFunctions.ConvertCmInMm(sirina) + "x" + AuxFunctions.ConvertCmInMm(visina) + ".ipt";
                    newOcc.Name = name;
                    break;
                case "Rukohvat":
                    newOcc.Name = name;
                    break;
                case "Staklo":
                    elemUserParameters["Visina"].Value = visina - 0.6 - 2 * 7;
                    elemUserParameters["Sirina"].Value = sirina - 0.6 - 2 * 7;
                    fileName = "Glass " + AuxFunctions.ConvertCmInMm(sirina) + "x" + AuxFunctions.ConvertCmInMm(visina) + ".ipt";
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

        private static void MakeDoorAssembly(int redniBroj)
        {
            ComponentOccurrences allOccurrs = mainDoc.ComponentDefinition.Occurrences;

            AssemblyDocument oMainDoc = thisApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, vrataFolder + @"\GenVrata.iam",false) as AssemblyDocument;

            AssemblyComponentDefinition oCompDef = oMainDoc.ComponentDefinition;

            //position element
            newOcc = allOccurrs.AddByComponentDefinition((ComponentDefinition)oCompDef, AuxFunctions.CreateMatrix("X", "Z", "-Y", 0.4, 1, 0.4));

            oMainDoc.Close();

            newOcc.Name = "Leva Vrata_"+redniBroj.ToString();
        }

        public static void Inject(AssemblyDocument doc, InventorServer app)
        {
            thisApp = app;
            mainDoc = doc;
        }

    }
}
