using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using EpdToExcel.Core;
using EpdToExcel.Core.Models;

namespace EpdToExcel.Console.Test
{
    class Program
    {
        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();
        private static IntPtr ThisConsole = GetConsoleWindow();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int HIDE = 0;
        private const int MAXIMIZE = 3;
        private const int MINIMIZE = 6;
        private const int RESTORE = 9;


        private static void L(string msg, System.ConsoleColor color = ConsoleColor.Green)
        {
            System.Console.ForegroundColor = color;
            System.Console.WriteLine(msg + "   " + Environment.NewLine);
            System.Console.ForegroundColor = ConsoleColor.Gray;
        }


        static void Main(string[] args)
        {
            System.Console.SetWindowSize(System.Console.LargestWindowWidth, System.Console.LargestWindowHeight);
            ShowWindow(ThisConsole, MAXIMIZE);

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

            L("Indikatoren wählen:", ConsoleColor.Cyan);
            L("Leertaste = aus/ab -wählen | Enter = bestätigen | Navigation mit Pfeiltasten | Escape = ALLE aus/ab -wählen", ConsoleColor.DarkYellow);

            var indicatorMenu = new List<Tuple<string, Action<int, bool>>>();
            var selectedIndicators = new List<string>();
            var indicators = EpdToXlsx.IndicatorKeyNameMapping.Values.ToList();

            for (var i = 0; i < indicators.Count(); i++)
            {
                selectedIndicators.Add(EpdToXlsx.IndicatorKeyNameMapping.Keys.ElementAt(i));

                indicatorMenu.Add(
                    new Tuple<string, Action<int, bool>>(indicators[i], (index, selected) =>
                    {
                        var epdKey = EpdToXlsx.IndicatorKeyNameMapping.Keys.ElementAt(index);
                        selectedIndicators.Remove(epdKey);
                        if (selected)
                        {
                            selectedIndicators.Add(epdKey);
                        }
                    }));
            }

            ShowIndicatorSelectionList(indicatorMenu);

            var epdFiles = new DirectoryInfo(epdFolder).GetFiles("*.xml");

            L($"\nFound {epdFiles.Count()} xml file(s).");

            L("Start importing ...");

            List<IEnumerable<Epd>> epds = new List<IEnumerable<Epd>>();
            for (int i = 0; i < epdFiles.Count(); i++)
            {
                try
                {
                    L($"{i + 1}. {epdFiles[i].Name} done.");
                    epds.Add(EpdToXlsx.GetEpdFromXml(epdFiles[i].FullName, i + 1, selectedIndicators));
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


        private static void ClearCurrentLine(string currentText, bool selected)
        {
            // Clear the emphaziser " <--"
            System.Console.SetCursorPosition(currentText.Length, System.Console.CursorTop);
            System.Console.Write(new string(' ', 4));

            System.Console.SetCursorPosition(0, System.Console.CursorTop);
            System.Console.ForegroundColor = selected ? ConsoleColor.Cyan : ConsoleColor.Gray;

            var adjustedText = currentText.Remove(currentText.Length - 2, 1).Insert(currentText.Length - 2, selected ? "X" : " ");

            System.Console.Write(adjustedText);
            System.Console.SetCursorPosition(System.Console.CursorLeft - 2, System.Console.CursorTop);
        }


        private static void EmphaziseCurrentLine(string text, bool selected)
        {
            // Emphazise new line
            System.Console.SetCursorPosition(0, System.Console.CursorTop);
            System.Console.ForegroundColor = selected ? ConsoleColor.Cyan : ConsoleColor.Gray;

            var adjustedText = text.Remove(text.Length - 2, 1).Insert(text.Length - 2, selected ? "X" : " ");

            System.Console.Write(adjustedText +  " <--");
            System.Console.SetCursorPosition(System.Console.CursorLeft - 6, System.Console.CursorTop);
        }


        private static void ShowIndicatorSelectionList(List<Tuple<string, Action<int, bool>>> selectionList)
        {
            var adjustedSelectionList = selectionList.Select((t, i) => (i + 1).ToString() + ". " + t.Item1).ToList();
            var longestEntryLength = adjustedSelectionList.Select(t => t.Count()).Max();
            adjustedSelectionList = adjustedSelectionList.Select(t => t + new string(' ', longestEntryLength - t.Length) + " [X]").ToList();

            foreach (var item in adjustedSelectionList)
            {
                System.Console.ForegroundColor = ConsoleColor.Cyan;
                System.Console.WriteLine(item);
            }

            var initialCursorPosTop = System.Console.CursorTop;

            System.Console.SetCursorPosition(longestEntryLength + 2, initialCursorPosTop - selectionList.Count());

            var selected = new Dictionary<int, bool>();

            var allSelect = true;
            for (int i = 0; i < selectionList.Count; i++)
            {
                selected.Add(i, allSelect);
            }

            int currentEntry = 0;

            EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

            while (true)
            {
                switch (System.Console.ReadKey(true).Key)
                {
                    case ConsoleKey.Spacebar:

                        var newChar = selected[currentEntry] ? ' ' : 'X';
                        selected[currentEntry] = !selected[currentEntry];

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        System.Console.SetCursorPosition(System.Console.CursorLeft - 1, System.Console.CursorTop);

                        selectionList[currentEntry].Item2(currentEntry, selected[currentEntry]);

                        break;

                    case ConsoleKey.DownArrow:

                        ClearCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        if (currentEntry + 1 < selectionList.Count)
                        {
                            currentEntry++;
                            System.Console.SetCursorPosition(0, System.Console.CursorTop + 1);
                        }
                        else
                        {
                            currentEntry = 0;
                            System.Console.SetCursorPosition(0, System.Console.CursorTop - (selectionList.Count() - 1));
                        }

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        break;

                    case ConsoleKey.UpArrow:

                        ClearCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        if (currentEntry - 1 >= 0)
                        {
                            currentEntry--;
                            System.Console.SetCursorPosition(System.Console.CursorLeft, System.Console.CursorTop + -1);
                        }
                        else
                        {
                            currentEntry = selectionList.Count() - 1;
                            System.Console.SetCursorPosition(System.Console.CursorLeft, System.Console.CursorTop + currentEntry);
                        }

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        break;

                    case ConsoleKey.Escape:

                        allSelect = !allSelect;

                        var selectChar = allSelect ? 'X' : ' ';

                        System.Console.SetCursorPosition(System.Console.CursorLeft, System.Console.CursorTop - currentEntry - 1);

                        for (int i = 0; i < selectionList.Count(); i++)
                        {
                            System.Console.SetCursorPosition(System.Console.CursorLeft, System.Console.CursorTop + 1);

                            ClearCurrentLine(adjustedSelectionList[i], allSelect);

                            selectionList[i].Item2(i, allSelect);
                            selected[i] = allSelect;
                        }

                        System.Console.SetCursorPosition(System.Console.CursorLeft, System.Console.CursorTop - selectionList.Count() + currentEntry + 1);

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], allSelect);

                        break;
                    case ConsoleKey.Enter:

                        System.Console.SetCursorPosition(0, System.Console.CursorTop + (selectionList.Count() - currentEntry));
                        return;
                }
            }
        }
    }
}
