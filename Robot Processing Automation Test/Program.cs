
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Robot_Processing_Automation_Test.DLL;

namespace Robot_Processing_Automation_Test
{
    
    static class Program
    {
        private static void startNotepad(String bodyText)
        {
            Process myProcesses;

            Process[] myProcess = Process.GetProcessesByName("notepad");
            if (myProcess.Length == 0)
            {
                myProcesses = Process.Start("notepad");
                Console.WriteLine("Started a new instance of notepad");
            }
            else
            {
                myProcesses = myProcess[0];
                Console.WriteLine("Selected existing notepad");
            }
            Thread.Sleep(1000);

            IntPtr p = myProcesses.MainWindowHandle;

            WindowsDLL.SetForegroundWindow(p);
            Thread.Sleep(1000);
            IntPtr hwnd = WindowsDLL.GetForegroundWindow();
            WindowsDLL.SetForegroundWindow(hwnd);
            WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), bodyText);
        }
        private static void NavigateToSave()
        {
            WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "%F");
            Thread.Sleep(1000);
            for (int j = 0; j < 4; j++)
            {

                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{DOWN}");
            }

            Thread.Sleep(1000);
            WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{ENTER}");
        }

        private static void saveDialog(String nameOfFile, String filePath, IntPtr mainNotepadWindow)
        {
            //Puts a delay as it takes for the new window to open
            Thread.Sleep(1000);
            //Checks if the new window is different than the main notepad one
            if (WindowsDLL.GetForegroundWindow() != mainNotepadWindow)
            {
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), nameOfFile);
                Thread.Sleep(1000);
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{TAB}");
                Thread.Sleep(1000);
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "%d");
                Thread.Sleep(1000);
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), filePath);
                for (int i=0;i<10;i++)
                {

                    WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{TAB}");
                }
                Thread.Sleep(1000);
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{ENTER}");
                
            }
            else
            {
                Environment.Exit(0);
            }
        }
        private static void ifFileExists(IntPtr mainNotepadWindow)
        {
            //Puts a delay as it takes for the new window to open
            Thread.Sleep(1000);
            //Checks if the new window is different than the main notepad one which it shouldve gone to file file didnt exist
            if (WindowsDLL.GetForegroundWindow() != mainNotepadWindow)
            {
                Console.WriteLine("File already exists. Overwriting");
                //file exists
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{LEFT}");
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "{ENTER}");
            }
            else
            {
              //  Environment.Exit(0);
            }
        }
        private static void closeNotepad(IntPtr mainNotepadWindow)
        {
            Thread.Sleep(1000);
                if (WindowsDLL.GetForegroundWindow() == mainNotepadWindow)
                {
                WindowsDLL.SendKeys(WindowsDLL.GetForegroundWindow(), "%{F4}");
                }
         }

        
            static void Main()
        {
            //main driver for the program

            String notepadBodyText = "Hello World";
            String nameOfFile = "fileName";
            String notepadFilePath = "Documents";


            //Starts a new notepad or selects an existing notepad
            startNotepad(notepadBodyText);
            //used to compare if we are back to the mainnotepad
            IntPtr mainNotepadWindow = WindowsDLL.GetForegroundWindow();
            NavigateToSave();
            Console.WriteLine("Navgigation to Save Dialog successful");
            saveDialog(nameOfFile, notepadFilePath,mainNotepadWindow);
            Console.WriteLine("Entered file name: " + nameOfFile+"\nEntered Filepath: " + notepadFilePath);
            ifFileExists(mainNotepadWindow);
            closeNotepad(mainNotepadWindow);
            Console.WriteLine("Exited Notepad successfully");
        }
    }
}
