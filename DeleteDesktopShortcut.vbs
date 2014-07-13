''''''''''''''''''''''''''''''''''
' 
' This VB script removes the requested desktop shortcuts
' 
' Change only the file name (test.lnk)
'
' Script created by Holger Habermehl. October 23, 2012 
''''''''''''''''''''''''''''''''''

Set Shell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
DesktopPath = Shell.SpecialFolders("Desktop")
FSO.DeleteFile DesktopPath & "\test.lnk"
