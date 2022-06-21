# Sample: MSIX-Excel-Automation-Container

This sample demonstrates the launching of Microsoft Excel in the MSIX context of an app inter-operating with it via automation. This is necessary when Excel macros need to read the contents of the virtual ```HKEY Current User``` registry and file system of the app driving Excel.

This sample uses the ```Microsoft.Office.Interop.Excel``` NuGet library in a WPF .NET Framework Console app.

# Prerequisites

1. Visual Studio 2017 17.3 Preview 3 - Community Edition or greater
2. Microsoft Excel with Excel.exe on the PATH.

## Setup
1. Clone this repository and copy ```test.xlsm``` to your ```c:\temp``` folder. 
2. Using **Regedit**, create the Registry key HKEY_CURRENT_USER\SOFTWARE\Contoso\AB-XL CurrentVersion (String value) and set the value to 1.
3. Close **Regedit**
4. Build and Deploy the application. You don't need to run it yet.
5. Run the following PowerShell command (start PowerShell as admin) to start **Regedit** in the context of the ```ConsoleApp4``` app just deployed.
```PowerShell
Invoke-CommandInDesktopPackage -PackageFamilyName "ConsoleApp4_n3sawgb4qe5x4" -AppId "App" -Command "regedit.exe" -PreventBreakaway
```
6. Create the key HKEY_CURRENT_USER\SOFTWARE\Contoso\AB-XL CurrentVersion (String value) and "change" it from 1 to 5. (This updates the virtual registry for ConsoleApp4)
5. Close **Regedit**.

## Usage
1. In Visual Studio, set a breakpoint on line 38 - after calling the ```StartExcelInMSIXContainer()``` method.
2. Press F5 to start the app in the debugger. Note that the Packaging project is the Startup project and that the app is running in the MSIX context. The app should start and stop and the breakpoint. Wait until Excel has fully started before continuing.

3. Once Excel has loaded, press F5 to continue with the app, using the interop library, it will open ```c:\temp\test.xlsm``` and call the following macro:

```vb
Sub read_hkcu_silent()
Dim windowsShell
Dim regValue
Set windowsShell = CreateObject("WScript.Shell")

regValue = windowsShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Contoso\AB-XL\CurrentVersion")
ActiveSheet.Range("A1").Value = regValue

End Sub
```
3. The app will read this value written to cell A1 and display it in the console.
4. The value should be 5. **Demonstrating it is reading from the virtual registry.**
5. In the console, type any key to close the app.
6. Close Excel.
6. Open ```c:\temp\test.xlsm``` and click the button 'Read HKCU Registry'. You should see a message box display the number 1. **This demonstrates Excel reading from the real registry.**

## Application Code
```C#
public static void StartExcelInMSIXContainer()
{
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
```

## Areas to improve

Add wait functions after launching Excel before calling the code that opens the macro file. This necessary to ensure Excel is ready to connect and receive input.
