using System;
using System.Text;
using System.Runtime.InteropServices;


namespace Robot_Processing_Automation_Test.DLL
{
    public static class WindowsDLL
    {

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);


        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

  
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

     

        public static void SendKeys(IntPtr windowHandle, string key)
        {
            if (SetForegroundWindow(windowHandle))
               System.Windows.Forms.SendKeys.SendWait(key);
        }



        
    }
}
