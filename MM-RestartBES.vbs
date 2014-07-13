Const MAX_ITER = 600

Set shell = CreateObject("WScript.Shell")
shell.LogEvent 4, "RestartBES has started"

Set objWMIService = _
    GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set colService = objWMIService.ExecQuery( _
    "Select * from Win32_Service Where Name = 'Lotus Domino Server (CD0minodata)'")

For Each svc In colService
    svc.StopService
	'WScript.Sleep 5000
Next

iter = 0
Do
    Set colService = objWMIService.ExecQuery( _
        "Select * from Win32_Service Where Name = 'Lotus Domino Server (CD0minodata)'")

    done = True
    For Each svc In colService
        If svc.State <> "Stopped" Then
            done = False
            WScript.Sleep 1000
            Exit For
        End If

	 If svc.State = "Stopped" Then

		shell.LogEvent 0, "Domino Service Stopped Successfully"

	 End IF

    Next
	

    iter = iter + 1
    If iter > MAX_ITER Then
        
 	
  	Dim colProcessList
  	Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'nservice.exe'")
  	For Each objProcess in colProcessList 
    		objProcess.Terminate() 
	shell.LogEvent 4, "Domino service could not gracefully be Stopped, nservice.exe process was forcefully terminated!"
  	Next  


        Exit Do
    End If
Loop Until done

'WScript.Echo "Ready to Restart"

Set WSHShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "C:\WINDOWS\system32\shutdown.exe /r /t 1 /f"
shell.LogEvent 0, "Server Re-Starting"
