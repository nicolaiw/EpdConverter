using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using EpdToExcel.Core;
using EpdToExcel.Core.Models;

namespace EpdToExcel.Console.Test
{
    class Program
    {
        private static void L(string msg, System.ConsoleColor color = ConsoleColor.Green)
        {
            System.Console.ForegroundColor = color;
            System.Console.WriteLine(msg +  "   " + Environment.NewLine);
            System.Console.ForegroundColor = ConsoleColor.White;
        }

        static void Main(string[] args)
        {

            string text;
            using (var client = new WebClient())
            {
                text = client.DownloadString("http://www.oekobaudat.de/OEKOBAU.DAT/resource/unitgroups/838aaa22-0117-11db-92e3-0800200c9a66?format=xmll&version=03.00.000");
            //return;?format=xml");
            }

            // http://www.oekobaudat.de/OEKOBAU.DAT/resource/datastocks/cc02f499-6b0f-4556-bb4a-7abe48e55f71/processes/88559403-7658-48f2-bac9-7986b4d0f4c2?format=xml&lang=de
            // /OEKOBAU.DAT/resource/flowproperties/93a60a56-a3c8-11da-a746-0800200b9a66?format=html&amp;version=03.00.000
            // /unitgroups/ad38d542-3fe9-439d-9b95-2f5f7752acaf.xml?format=xml
            // http://www.oekobaudat.de/OEKOBAU.DAT/resource/flows/cf76b28f-3e3f-406a-aad0-df0b13c8d6e6?format=xml&version=33.00.000
            //return;
            System.Console.Write("Name des Ordners auf dem Desktop: ");
            var projectFolder = System.Console.ReadLine();
            var epdFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectFolder);


            while (!Directory.Exists(epdFolder))
            {
                L($"Ein Ordner mit dem Namen {projectFolder} existiert nicht.", ConsoleColor.Red);
                System.Console.Write("Bitte anderen Ordner wählen: ");
                projectFolder = System.Console.ReadLine();
                epdFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectFolder);
            }


            System.Console.Write("\nName den das Projekt erhalten soll: ");
            var projectName = System.Console.ReadLine();
            var projectFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectName + ".xlsx");


            while (File.Exists(projectFile))
            {
                L($"Ein Projekt mit dem Namen {projectName} existiert bereits", ConsoleColor.Red);
                System.Console.Write("Bitte neuen Namen vergeben: ");
                projectName = System.Console.ReadLine();
                projectFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectName + ".xlsx");
            }

            var epdFiles = new DirectoryInfo(epdFolder).GetFiles("*.xml");

            L($"\nFound {epdFiles.Count()} xml file(s).");

            L("Start importing ...");

            List<IEnumerable<Epd>> epds = new List<IEnumerable<Epd>>();
            for (int i = 0; i < epdFiles.Count(); i++)
            {
                try
                {
                    L($"{i +1}. {epdFiles[i].Name} done.");
                    epds.Add(EpdToXlsx.GetEpdFromXml(epdFiles[i].FullName, i + 1));
                }
                catch (Exception ex)
                {
                    L($"Import failed: {epdFiles[i].Name}.", ConsoleColor.Red);
                    L(ex.ToString(), ConsoleColor.Red);

                    System.Console.ReadLine();
                    L("Press any key to continue.", ConsoleColor.Cyan);
                }
            }


            try
            {
                L("Start exporting ...");
                L("This may take a while. Please wait ...");
                EpdToXlsx.ExportEpdsToXlsx(epds, projectFile);
                L("Export succeeded.");
                L("Press any key to close this window.", ConsoleColor.White);
            }
            catch (Exception ex)
            {
                L($"Export failed:", ConsoleColor.Red);
                L(ex.ToString(), ConsoleColor.Red);
            }


            System.Console.ReadLine();
        }
    }
}
