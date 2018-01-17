using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using EpdToExcel.Core;
using EpdToExcel.Core.Models;
using C = System.Console;

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
            C.ForegroundColor = color;
            C.WriteLine(msg + "   " + Environment.NewLine);
            C.ForegroundColor = ConsoleColor.Gray;
        }


        static void Main(string[] args)
        {
            C.SetWindowSize(System.Console.LargestWindowWidth, System.Console.LargestWindowHeight);
            ShowWindow(ThisConsole, MAXIMIZE);

            C.Write("Name des Ordners auf dem Desktop: ");
            var projectFolder = System.Console.ReadLine();
            var epdFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectFolder);

            while (!Directory.Exists(epdFolder))
            {
                L($"Ein Ordner mit dem Namen {projectFolder} existiert nicht.", ConsoleColor.Red);
                C.Write("Bitte anderen Ordner wählen: ");
                projectFolder = C.ReadLine();
                epdFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectFolder);
            }

            C.Write("\nName den das Projekt erhalten soll: ");
            var projectName = C.ReadLine();
            var projectFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), projectName + ".xlsx");

            while (File.Exists(projectFile))
            {
                L($"Ein Projekt mit dem Namen {projectName} existiert bereits", ConsoleColor.Red);
                C.Write("Bitte neuen Namen vergeben: ");
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
                    epds.Add(EpdToXlsx.GetEpdFromXml(epdFiles[i].FullName, i + 1, selectedIndicators, str => L(str, ConsoleColor.Yellow)));
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

            C.ReadLine();
        }


        private static void ClearCurrentLine(string currentText, bool selected)
        {
            // Clear the emphaziser " <--"
            C.SetCursorPosition(currentText.Length, C.CursorTop);
            C.Write(new string(' ', 4));

            C.SetCursorPosition(0, C.CursorTop);
            C.ForegroundColor = selected ? ConsoleColor.Cyan : ConsoleColor.Gray;

            var adjustedText = currentText.Remove(currentText.Length - 2, 1).Insert(currentText.Length - 2, selected ? "X" : " ");

            C.Write(adjustedText);
            C.SetCursorPosition(C.CursorLeft - 2, C.CursorTop);
        }


        private static void EmphaziseCurrentLine(string text, bool selected)
        {
            // Emphazise new line
            C.SetCursorPosition(0, C.CursorTop);
            C.ForegroundColor = selected ? ConsoleColor.Cyan : ConsoleColor.Gray;

            var adjustedText = text.Remove(text.Length - 2, 1).Insert(text.Length - 2, selected ? "X" : " ");

            C.Write(adjustedText +  " <--");
            C.SetCursorPosition(C.CursorLeft - 6, C.CursorTop);
        }


        private static void ShowIndicatorSelectionList(List<Tuple<string, Action<int, bool>>> selectionList)
        {
            var adjustedSelectionList = selectionList.Select((t, i) => (i + 1).ToString() + ". " + t.Item1).ToList();
            var longestEntryLength = adjustedSelectionList.Select(t => t.Count()).Max();
            adjustedSelectionList = adjustedSelectionList.Select(t => t + new string(' ', longestEntryLength - t.Length) + " [X]").ToList();

            foreach (var item in adjustedSelectionList)
            {
                C.ForegroundColor = ConsoleColor.Cyan;
                C.WriteLine(item);
            }

            var initialCursorPosTop = C.CursorTop;

            C.SetCursorPosition(longestEntryLength + 2, initialCursorPosTop - selectionList.Count());

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
                switch (C.ReadKey(true).Key)
                {
                    case ConsoleKey.Spacebar:

                        var newChar = selected[currentEntry] ? ' ' : 'X';
                        selected[currentEntry] = !selected[currentEntry];

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        C.SetCursorPosition(C.CursorLeft - 1, C.CursorTop);

                        selectionList[currentEntry].Item2(currentEntry, selected[currentEntry]);

                        break;

                    case ConsoleKey.DownArrow:

                        ClearCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        if (currentEntry + 1 < selectionList.Count)
                        {
                            currentEntry++;
                            C.SetCursorPosition(0, C.CursorTop + 1);
                        }
                        else
                        {
                            currentEntry = 0;
                            C.SetCursorPosition(0, C.CursorTop - (selectionList.Count() - 1));
                        }

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        break;

                    case ConsoleKey.UpArrow:

                        ClearCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        if (currentEntry - 1 >= 0)
                        {
                            currentEntry--;
                            C.SetCursorPosition(C.CursorLeft, C.CursorTop + -1);
                        }
                        else
                        {
                            currentEntry = selectionList.Count() - 1;
                            C.SetCursorPosition(C.CursorLeft, C.CursorTop + currentEntry);
                        }

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], selected[currentEntry]);

                        break;

                    case ConsoleKey.Escape:

                        allSelect = !allSelect;

                        var selectChar = allSelect ? 'X' : ' ';

                        C.SetCursorPosition(C.CursorLeft, C.CursorTop - currentEntry - 1);

                        for (int i = 0; i < selectionList.Count(); i++)
                        {
                            C.SetCursorPosition(C.CursorLeft, C.CursorTop + 1);

                            ClearCurrentLine(adjustedSelectionList[i], allSelect);

                            selectionList[i].Item2(i, allSelect);
                            selected[i] = allSelect;
                        }

                        C.SetCursorPosition(C.CursorLeft, C.CursorTop - selectionList.Count() + currentEntry + 1);

                        EmphaziseCurrentLine(adjustedSelectionList[currentEntry], allSelect);

                        break;
                    case ConsoleKey.Enter:

                        C.SetCursorPosition(0, C.CursorTop + (selectionList.Count() - currentEntry));
                        return;
                }
            }
        }
    }
}
