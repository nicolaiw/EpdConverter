using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
