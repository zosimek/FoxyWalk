param($installPath, $toolsPath, $package, $project)


# Get the ID and security principal of the current user account
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent();
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID);

# Get the security principal for the administrator role
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator;

# Check to see if we are currently running as an administrator
if ($myWindowsPrincipal.IsInRole($adminRole))
{
    # We are running as an administrator, so change the title and background colour to indicate this
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)";
    $Host.UI.RawUI.BackgroundColor = "DarkBlue";
    Clear-Host;
}
else {
    # We are not running as an administrator, so relaunch as administrator

    # Create a new process object that starts PowerShell
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";

    # Specify the current script path and name as a parameter with added scope and support for scripts with spaces in it's path
    $newProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"

    # Indicate that the process should be elevated
    $newProcess.Verb = "runas";

    # Start the new process
    [System.Diagnostics.Process]::Start($newProcess);

    # Exit from the current, unelevated, process
    Exit;
}

# Run your code that needs to be elevated here...

# Update system PATH
$pathx = $toolsPath + '\Filters;'
$pathx += $toolsPath + '\FFMPEG;'
$pathx += $toolsPath + '\LAV\x64;'
$pathx += $toolsPath + '\LAV\x86;'
$pathx += $toolsPath + '\WebM;'
$pathx += $toolsPath + '\Xiph'

[Environment]::SetEnvironmentVariable("Path", $env:Path + ";" + $pathx, [EnvironmentVariableTarget]::Machine)

# Run filters registration
$argu = "-u"
$CMD = $toolsPath + '\Filters\reg_special.exe'
& $CMD $argu

$CMD = $toolsPath + '\FFMPEG\reg_special.exe'
& $CMD $argu

$CMD = $toolsPath + '\LAV\x64\reg_special.exe'
& $CMD $argu

$CMD = $toolsPath + '\LAV\x86\reg_special.exe'
& $CMD $argu

$CMD = $toolsPath + '\WebM\reg_special.exe'
& $CMD $argu

$CMD = $toolsPath + '\Xiph\reg_special.exe'
& $CMD $argu

# Remove from GAC
# [System.Reflection.Assembly]::Load("System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")            
# $publish = New-Object System.EnterpriseServices.Internal.Publish   

# $argx = $installPath + '\lib\net45\VisioForge.Types.dll'
# $publish.GacRemove($argx)      

# $argx = $installPath + '\lib\net45\VisioForge.MediaFramework.dll'
# $publish.GacRemove($argx)      

# $argx = $installPath + '\lib\net45\VisioForge.Shared.dll'
# $publish.GacRemove($argx)      

# $argx = $installPath + '\lib\net45\VisioForge.Tools.dll'
# $publish.GacRemove($argx)      

# $argx = $installPath + '\lib\net45\VisioForge.DirectX.dll'
# $publish.GacRemove($argx)      

# $argx = $installPath + '\lib\net45\VisioForge.Controls.dll'
# $publish.GacRemove($argx)      

# $argx = $installPath + '\lib\net45\VisioForge.Controls.UI.dll'
# $publish.GacRemove($argx) 

# $argx = $installPath + '\lib\net45\VisioForge.Controls.UI.Dialogs.dll'
# $publish.GacRemove($argx) 




# # open json.net splash page on package install
# # don't open if json.net is installed as a dependency

# try
# {
#   $url = "http://www.newtonsoft.com/json/install?version=" + $package.Version
#   $dte2 = Get-Interface $dte ([EnvDTE80.DTE2])

#   if ($dte2.ActiveWindow.Caption -eq "Package Manager Console")
#   {
#     # user is installing from VS NuGet console
#     # get reference to the window, the console host and the input history
#     # show webpage if "install-package newtonsoft.json" was last input

#     $consoleWindow = $(Get-VSComponentModel).GetService([NuGetConsole.IPowerConsoleWindow])

#     $props = $consoleWindow.GetType().GetProperties([System.Reflection.BindingFlags]::Instance -bor `
#       [System.Reflection.BindingFlags]::NonPublic)

#     $prop = $props | ? { $_.Name -eq "ActiveHostInfo" } | select -first 1
#     if ($prop -eq $null) { return }
  
#     $hostInfo = $prop.GetValue($consoleWindow)
#     if ($hostInfo -eq $null) { return }

#     $history = $hostInfo.WpfConsole.InputHistory.History

#     $lastCommand = $history | select -last 1

#     if ($lastCommand)
#     {
#       $lastCommand = $lastCommand.Trim().ToLower()
#       if ($lastCommand.StartsWith("install-package") -and $lastCommand.Contains("newtonsoft.json"))
#       {
#         $dte2.ItemOperations.Navigate($url) | Out-Null
#       }
#     }
#   }
#   else
#   {
#     # user is installing from VS NuGet dialog
#     # get reference to the window, then smart output console provider
#     # show webpage if messages in buffered console contains "installing...newtonsoft.json" in last operation

#     $instanceField = [NuGet.Dialog.PackageManagerWindow].GetField("CurrentInstance", [System.Reflection.BindingFlags]::Static -bor `
#       [System.Reflection.BindingFlags]::NonPublic)

#     $consoleField = [NuGet.Dialog.PackageManagerWindow].GetField("_smartOutputConsoleProvider", [System.Reflection.BindingFlags]::Instance -bor `
#       [System.Reflection.BindingFlags]::NonPublic)

#     if ($instanceField -eq $null -or $consoleField -eq $null) { return }

#     $instance = $instanceField.GetValue($null)

#     if ($instance -eq $null) { return }

#     $consoleProvider = $consoleField.GetValue($instance)
#     if ($consoleProvider -eq $null) { return }

#     $console = $consoleProvider.CreateOutputConsole($false)

#     $messagesField = $console.GetType().GetField("_messages", [System.Reflection.BindingFlags]::Instance -bor `
#       [System.Reflection.BindingFlags]::NonPublic)
#     if ($messagesField -eq $null) { return }

#     $messages = $messagesField.GetValue($console)
#     if ($messages -eq $null) { return }

#     $operations = $messages -split "=============================="

#     $lastOperation = $operations | select -last 1

#     if ($lastOperation)
#     {
#       $lastOperation = $lastOperation.ToLower()

#       $lines = $lastOperation -split "`r`n"

#       $installMatch = $lines | ? { $_.StartsWith("------- installing...newtonsoft.json ") } | select -first 1

#       if ($installMatch)
#       {
#         $dte2.ItemOperations.Navigate($url) | Out-Null
#       }
#     }
#   }
# }
# catch
# {
#   try
#   {
#     $pmPane = $dte2.ToolWindows.OutputWindow.OutputWindowPanes.Item("Package Manager")

#     $selection = $pmPane.TextDocument.Selection
#     $selection.StartOfDocument($false)
#     $selection.EndOfDocument($true)

#     if ($selection.Text.StartsWith("Attempting to gather dependencies information for package 'Newtonsoft.Json." + $package.Version + "'"))
#     {
#       # don't show on upgrade
#       if (!$selection.Text.Contains("Removed package"))
#       {
#         $dte2.ItemOperations.Navigate($url) | Out-Null
#       }
#     }
#   }
#   catch
#   {
#     # stop potential errors from bubbling up
#     # worst case the splash page won't open  
#   }
# }

# # still yolo