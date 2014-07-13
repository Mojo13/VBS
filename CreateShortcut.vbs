Option Explicit

On Error Resume Next

Dim objShell
Dim objDesktop
Dim  objLink
Dim strAppPath
Dim strWorkDir
Dim strIconPath



strWorkDir ="C:\windows"
strAppPath = "http://YourCompanyIntranetSiteURL/"  'you have to use your URL to Interanet site or path to specific program
strIconPath = "\\server\Xyz.ico"					'specify the path to the icon please change to your valid path

Set objShell = CreateObject("WScript.Shell")
objDesktop = objShell.SpecialFolders("Desktop")
Set objLink = objShell.CreateShortcut(objDesktop & "\YourShortcutName.lnk") 'change here To your shortcut name


objLink.Description = "your shortcut description" 'replaec with your description
objLink.IconLocation = strIconPath 
objLink.TargetPath = strAppPath
objLink.WindowStyle = 3
objLink.WorkingDirectory = strWorkDir
objLink.Save
