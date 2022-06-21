using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Xml.Linq;
using Windows.ApplicationModel;

namespace ConsoleApp4
{
    internal class Program
    {
        public static void StartExcelInMSIXContainer()
        {

            // Invoke-CommandInDesktopPackage -PackageFamilyName "ConsoleApp4_n3sawgb4qe5x4" -AppId "App" -Command "cmd.exe" -PreventBreakaway
            var packageFamilyName = AppInfo.Current.PackageFamilyName;
            var appId = AppInfo.Current.Id;
            var psCmd = $"Invoke-CommandInDesktopPackage -PackageFamilyName {packageFamilyName} -AppId {appId} -Command excel.exe -PreventBreakaway";
            var startInfo = new ProcessStartInfo()
            {
                
                FileName = "powershell.exe",
                Arguments = $"-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile \"{psCmd}\"",
                UseShellExecute = false
            };
            var proc = Process.Start(startInfo);
        }

        static void Main(string[] args)
        {
            StartExcelInMSIXContainer();
            var file = @"c:\temp\test.xlsm";

            var activeObject = Marshal.GetActiveObject("Excel.Application");
            var eApp = (activeObject as Application);
            eApp.Visible = true;
            eApp.Workbooks.Open(file);
            var aSheet = eApp.ActiveSheet as Worksheet;
            aSheet.Activate();
            aSheet.Application.Run("sheet1.read_hkcu_silent");
            var res1 = aSheet.Range["A1"].Value;
            Console.WriteLine($"A1: {res1}");
            Console.ReadKey();
        }
    }
}
