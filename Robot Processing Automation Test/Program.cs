using System;
using System.Diagnostics;
using System.Text;
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
                Console.Write("Started a new instance of notepad: ");
            }
            else
            {
                myProcesses = myProcess[0];
                Console.Write("Selected existing notepad: ");
            }
            
            
            
            IntPtr p = myProcesses.MainWindowHandle;
            WindowsDLL.SetForegroundWindow(p);
            Thread.Sleep(1000);
            int chars = 256;
            StringBuilder buff = new StringBuilder(chars);
            IntPtr notepadMain = WindowsDLL.GetForegroundWindow();
            if (WindowsDLL.GetWindowText(notepadMain, buff, chars) > 0)


                if (buff.ToString().Contains("Notepad"))
                {

                    Console.WriteLine(" " + buff.ToString());

                    WindowsDLL.SendKeys(notepadMain, bodyText);
                    NavigateToSave();
                }
                else
                {

                    Console.WriteLine("Notepad was selected but " + buff.ToString() + " got windows foreground");
                    Environment.Exit(0);

                }


        }
        private static void NavigateToSave()
        {
            int chars = 256;
            StringBuilder buff = new StringBuilder(chars);
            IntPtr notepadMain = WindowsDLL.GetForegroundWindow();
            if (WindowsDLL.GetWindowText(notepadMain, buff, chars) > 0)


                if (buff.ToString().Contains("Notepad"))
                {

                    WindowsDLL.SendKeys(notepadMain, "%F");
                    Thread.Sleep(1000);
                    for (int j = 0; j < 4; j++)
                    {

                        WindowsDLL.SendKeys(notepadMain, "{DOWN}");
                    }

                    Thread.Sleep(1000);
                    WindowsDLL.SendKeys(notepadMain, "{ENTER}");

                    Console.WriteLine("Navgigation to Save Dialog successful");
                }
        }

        private static void saveDialog(String nameOfFile, String filePath)
        {
            Thread.Sleep(1000);
            int chars = 256;
            StringBuilder buff = new StringBuilder(chars);
            IntPtr saveAsDialog = WindowsDLL.GetForegroundWindow();
            if (WindowsDLL.GetWindowText(saveAsDialog, buff, chars) > 0)
            {
                //if window texts equals SAVE AS then it will proceed to put file name, file path and try to save 
                if (buff.ToString().Equals("Save As"))
                {
                    WindowsDLL.SendKeys(saveAsDialog, nameOfFile);
                    Thread.Sleep(1000);
                    WindowsDLL.SendKeys(saveAsDialog, "{TAB}");
                    Thread.Sleep(1000);
                    WindowsDLL.SendKeys(saveAsDialog, "%d");
                    Thread.Sleep(1000);
                    WindowsDLL.SendKeys(saveAsDialog, filePath);
                    Thread.Sleep(1000);
                    Console.WriteLine("Entered file name: " + nameOfFile + "\nEntered Filepath: " + filePath);
                    for (int i = 0; i < 10; i++)
                    {
                        
                        WindowsDLL.SendKeys(saveAsDialog, "{TAB}");
                        Thread.Sleep(500);
                    }
                    Thread.Sleep(1000);
                    WindowsDLL.SendKeys(saveAsDialog, "{ENTER}");
                    Thread.Sleep(1000);


                }

                
                IntPtr newOrPreviousWindow = WindowsDLL.GetForegroundWindow();
                if (WindowsDLL.GetWindowText(newOrPreviousWindow, buff, chars) > 0)
                {

                    if (buff.ToString().Equals("Confirm Save As"))
                    {

                        ifFileExists();

                    }
                    else if (buff.ToString().Equals(nameOfFile + " - Notepad"))
                    {
                        Console.WriteLine("File Saved");


                    }
                   
                }
            }
        }
        private static void ifFileExists()
        {
            
            int chars = 256;
            StringBuilder buff = new StringBuilder(chars);
            IntPtr comfirmSaveAsDialog = WindowsDLL.GetForegroundWindow();
            if (WindowsDLL.GetWindowText(comfirmSaveAsDialog, buff, chars) > 0)
            {
                //if window texts equals Confirm SAVE AS then it will proceed to overwrite

                if (buff.ToString().Equals("Confirm Save As"))
                {
                    Console.WriteLine("File already exists. Overwriting");
                    //file exists
                    WindowsDLL.SendKeys(comfirmSaveAsDialog, "{LEFT}");
                    WindowsDLL.SendKeys(comfirmSaveAsDialog, "{ENTER}");
                    Console.WriteLine("File Saved");

                }
            }
        }


       

        static void Main()

        {
            //main driver for the program

            String notepadBodyText = "Hello World";
            String nameOfFile = "fileName.txt";
            String notepadFilePath = "Documents";
            startNotepad(notepadBodyText);
            saveDialog(nameOfFile, notepadFilePath);
            
        }
    }
}
