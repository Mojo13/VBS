' ================================================================================================================================ 
' 1. The script gathers errors and warnings from event logs on remote computers, 
'    gathers the information on the free space on disks. 
' 2. Can create the report in a format *.csv and send the message on mailbox 
'    (the attached file). 
'    Uses smtp protocol. 
' 3. Can create the table in a database and fill with its events. In this case 
'    data gathering is carried out from the moment of the last registered event 
'    in a event log. Automatic cleaning of the table from out-of-date events. 
'    Uses source ODBC. 
' 
'    Excuse for my English! :) 
' 
'    DmitriiSQ@gmail.com 
' 
' ================================================================================================================================ 
 
Option Explicit 
On Error Resume Next 
 
Dim objFSO, objReportFile, objWMIService, objEvent, objMessage, colDisks, objDisk, colLoggedEvents 
Dim dtmScriptStart, dtmStartDate, dtmEndDate, DateToCheck 
Dim strMailBody 
Dim strBuffer 
Dim DateOffset 
Dim DBCleanDays 
Dim intEventCount 
Dim objCN, objCmd, objRS, strDSN 
Dim strDeviceID, strMessage, strLogFile, strTimeWritten, strSourceName, strCategory, strEventCode, strUser 
Dim strTemp, strDBStatus, strMailTo, strReportFile, strTBLName, strComputer, strTempType 
Dim strReportDelim, strSMTPServerName, strSMTPPortNumber, strMailFrom, strMailSubject 
Dim strSearchURL, strSearchURLStart, strSearchURLEnd 
 
' ================================================================================================================================ 
Dim arrNameServers(99) 
arrNameServers(0) = "SERVER_1" 
arrNameServers(1) = "" 
arrNameServers(2) = "SERVER_2" 
 
' ================================================================================================================================ 
strReportFile = "C:\EventLogsCollector.csv"    ' Name and ABSOLUTE PATH to a file with the report 
DateOffset = 1                                 ' Quantity of days for which there is a search in event logs (if the database is not used) 
DBCleanDays = 365                              ' Removal of records in a database is higher, than X days 
strMailTo = "user1@domain.com;user2@domain.com"  ' The mail address of the receiver. Admits to specify some addresses through a semicolon 
strDSN = "TEST_ODBC_DSN"                       ' ODBC DSN. If value is not set or set empty ("") - the base is not used 
strTBLName = "EventLogsCollector"              ' Table name in a database 
 
strReportDelim = ";"                           ' Delimiter in a report file 
strSMTPServerName = "alias.domain.com"         ' Name SMTP of the server 
strSMTPPortNumber = 25                         ' Number of port SMTP of a server 
strMailFrom = "EventLogsCollector@domain.com"  ' The mail address of the sender 
strMailSubject = "The collector of souls"              ' Subject of the mail message 
 
 
strSearchURLStart = "http://social.technet.microsoft.com/Search/ru-RU?query="  ' The beginning of a line search URL 
strSearchURLEnd = "&refinement=82&ac=8"                                        ' The end of a line search URL 
 
' ================================================================================================================================ 
' Initialization of operation with a database 
strDBStatus = "OK!" 
Set objCN = CreateObject("ADODB.Connection") 
Set objCmd = CreateObject("ADODB.Command") 
Set objRS = CreateObject("ADODB.RecordSet") 
If strDSN <> ""  And Not IsNull(strDSN) Then 
  If IsNull(strTBLName) Or strTBLName = "" Then 
    strDBStatus = "The name of the table of a database is not defined" 
    strDSN = Null 
  Else 
    Err.Clear 
    objCN.Open ("DSN=" & strDSN) 
    If Err.Number <> 0 Then 
      strDBStatus = "It was not possible to be connected to object ODBC DSN " & strDSN 
      strDSN = Null 
    Else 
      objCmd.ActiveConnection = objCN 
      objRS.ActiveConnection = objCN 
      Err.Clear 
      objRS.Open("select OBJECT_ID('" & strTBLName & "')") 
      If Err.Number <> 0 Then 
        strDBStatus = "Error of procedure call OBJECT_ID" 
        strDSN = Null 
      Else 
        If IsNull(objRS.Fields.Item(0)) Then 
          objCmd.CommandText = "Create Table " & strTBLName & "(PK bigint PRIMARY KEY IDENTITY(1,1) NOT NULL, " & _ 
                                      "Computer nvarChar(28) NULL, " & _ 
                                      "LogFile nvarChar(56) NULL, " & _ 
                                      "DTEvent DateTime NULL, " & _ 
                                      "TypeName nvarChar(28) NULL, " & _ 
                                      "SourceName nvarChar(256) NULL, " & _ 
                                      "Category Int NULL, " & _ 
                                      "EventCode Int NULL, " & _ 
                                      "UserName nvarChar(256) NULL, " & _ 
                                      "MessageText nvarchar(4000) NULL, " & _ 
                                      "URL nvarchar(256) NULL)" 
          Err.Clear 
          objCmd.Execute 
          If Err.Number <> 0 Then 
            strDBStatus = "Error of creation of the table " & strTBLName & " in object ODBC DSN " & strDSN 
            strDSN = Null 
          End If 
          objCmd.CommandText = "CREATE NONCLUSTERED INDEX idx_" & strTBLName & "_NCL ON " & strTBLName & " (DTEvent DESC, Computer, LogFile, TypeName, SourceName)" 
          Err.Clear 
          objCmd.Execute 
          If Err.Number <> 0 Then 
            strDBStatus = "Index creation error idx_" & strTBLName & "_NCL in object ODBC DSN " & strDSN 
            strDSN = Null 
          End If 
        End If 
        objRS.Close 
      End If 
    End If 
  End If 
Else 
  strDBStatus = "The name of object ODBC DSN is not set. Work without base." 
End If 
 
 
' ================================================================================================================================ 
intEventCount = 0 
Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime") 
Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime") 
Set dtmScriptStart = CreateObject("WbemScripting.SWbemDateTime") 
 
dtmStartDate.SetVarDate CDate(Date() & " " & Time()), True 
dtmEndDate.SetVarDate WMItoVBTime(dtmStartDate) 
dtmScriptStart.SetVarDate WMItoVBTime(dtmStartDate) 
dtmStartDate.Day = dtmStartDate.Day - DateOffset 
 
' ================================================================================================================================ 
strBuffer =  """Computer""" & strReportDelim & """EventLog""" & strReportDelim & """Date and Time""" & strReportDelim & """Type""" & strReportDelim & _ 
      """Source""" & strReportDelim & """Category""" & strReportDelim & """Event ID""" & strReportDelim & """User""" & strReportDelim & """Description""" & strReportDelim & """URL""" & vbCrLf 
 
strMailBody =  "=================================================================================" & vbCrLf & _ 
          "The report on errors and warnings in event logs." & vbCrLf & vbCrLf & _ 
          "The list of analysable hosts: " & vbCrLf 
          For Each strComputer in arrNameServers 
            If strComputer <> Empty Then 
              strMailBody = strMailBody & VbTab & UCase(strComputer) & vbCrLf 
            End If 
          Next 
          strMailBody = strMailBody & _ 
          "=================================================================================" & vbCrLf & vbCrLf & vbCrLf 
           
strMailBody = strMailBody &   "The list of hard disks, the common and the free space" & vbCrLf & _ 
        "---------------------------------------------------------------------------------" & vbCrLf  
 
' ================================================================================================================================ 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objReportFile = objFSO.CreateTextFile(strReportFile, True) 
For Each strComputer in arrNameServers 
  If strComputer <> Empty Then 
    If strDSN <> "" And Not IsNull(strDSN) Then 
      objCmd.CommandText = "begin transaction EventLogsCollector" 
      Err.Clear 
      objCmd.Execute 
      If Err.Number <> 0 Then 
        strDBStatus = "Error in transaction formation." 
        strDSN = Null 
      End If 
    End If 
    Err.Clear 
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    If Err.Number = 0 Then 
      strMailBody = strMailBody & vbCrLf & Left(UCase(strComputer) & ":                                ", 20) 
      colDisks = Empty 
      Err.Clear 
      Set colDisks = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk Where DriveType = 3") 
      If Err.Number = 0 Then 
        For Each objDisk in colDisks 
          strDeviceID = objDisk.DeviceID 
          If strDeviceID = "" Or IsEmpty(strDeviceID) Or IsNull(strDeviceID) Then strDeviceID = "-:" 
            strMailBody =  strMailBody & strDeviceID & " " &  _ 
                    Right("     (" & Round(objDisk.Size / 1024 / 1024 / 1024, 0) & "Gb)",8) & " " & _ 
                    Right("     (" & Round(objDisk.FreeSpace / 1024 / 1024 / 1024, 0) & "Gb)",8) & " " & _ 
                    Right("     (" & FormatPercent(objDisk.FreeSpace / objDisk.Size) & " Free)",14) & vbCrLf & "                    " 
        Next 
      End If 
       
      dtmStartDate.SetVarDate WMItoVBTime(dtmScriptStart) 
      If strDSN <> "" And Not IsNull(strDSN) Then 
        Err.Clear 
        objRS.Open("select MAX(DTEvent) from " & strTBLName & " where DTEvent <= '" & WMItoVBTime(dtmScriptStart) & "' And Computer = '" & UCase(strComputer) & "'") 
        If Err.Number = 0 Then 
          If Not IsNull(objRS.Fields.Item(0)) Then 
            dtmStartDate.SetVarDate CDate(objRS.Fields.Item(0)), True  ' Date and time of start of a script from a database (Date and time of start of a script <= to maximum date-time from base) 
            dtmStartDate.Microseconds = dtmStartDate.Microseconds + 1000  ' Plus 1 msec 
          Else 
            dtmStartDate.Year = dtmStartDate.Year - 1 
          End If 
          objRS.Close 
        Else 
          dtmStartDate.Year = dtmStartDate.Year - 1 
        End If 
      Else 
        dtmStartDate.Day = dtmStartDate.Day - DateOffset  ' Initial date and time of search of events 
      End If 
       
      dtmEndDate.SetVarDate CDate(Date() & " " & Time()), True 
      Set colLoggedEvents = objWMIService.ExecQuery _ 
                        ("Select * from Win32_NTLogEvent Where Logfile <> 'Security' " & _ 
                        "AND TimeWritten >= '" & dtmStartDate & "' " & "AND TimeWritten <= '" & dtmEndDate & "' " & _ 
                        "AND (Type = 'Error' OR Type = '??????' OR Type = 'Warning' OR Type = '??????????????')")    '  Russian and English events 
        For Each objEvent in colLoggedEvents 
            intEventCount = intEventCount + 1    ' Simply counter, for statistics 
 
            Select Case LCase(objEvent.Type)    ' Correction on a case of Russian OS 
              Case "??????"      strTempType = "error" 
              Case "??????????????"  strTempType = "warning" 
              Case Else        strTempType = LCase(objEvent.Type) 
            End Select 
            strMessage = objEvent.Message & "" 
              If IsEmpty(strMessage) Or IsNull(strMessage) Then strMessage = "" 
              strMessage = Trim (Replace (Replace (Replace (Replace (Replace (strMessage, """", "`"), vbCrLf, " "), "'", "`"), vbCr, " "), vbLf, " ")) 
            strLogFile = objEvent.LogFile & "" 
            strTimeWritten = WMItoVBTime(objEvent.TimeWritten) & "" 
            strSourceName = objEvent.SourceName & "" 
            strCategory = objEvent.Category & "" 
            strEventCode = objEvent.EventCode & "" 
            strUser = objEvent.User & "" 
            strSearchURL = strSearchURLStart & strEventCode & "%20" & strSourceName & strSearchURLEnd 
           
            strBuffer =  strBuffer & _ 
                """" & UCase(strComputer) & """" & strReportDelim & _ 
                """" & strLogFile & """" & strReportDelim & _ 
                """" & strTimeWritten & """" & strReportDelim & _ 
                """" & strTempType & """" & strReportDelim & _ 
                """" & strSourceName & """" & strReportDelim & _ 
                """" & strCategory & """" & strReportDelim & _ 
                """" & strEventCode & """" & strReportDelim & _ 
                """" & strUser & """" & strReportDelim & _ 
                """" & strMessage & """" & strReportDelim & _ 
                """" & strSearchURL & """" & _ 
                vbCrLf 
 
             
            If strDSN <> "" And Not IsNull(strDSN) Then    ' Attempt of adding of the data in transaction 
              objCmd.CommandText = "insert into " & strTBLName & " (Computer, LogFile, DTEvent, " & _ 
                                         "TypeName, SourceName, Category, " & _ 
                                         "EventCode, UserName, MessageText, URL) " & _ 
                         "values (" & _ 
                               "'" & UCase(strComputer) & "', " & _ 
                               "'" & strLogFile & "', " & _ 
                               "'" & strTimeWritten & "', " & _ 
                               "'" & strTempType & "', " & _ 
                               "'" & strSourceName & "', " & _ 
                               "'" & strCategory & "', " & _ 
                               "'" & strEventCode & "', " & _ 
                               "'" & strUser & "', " & _ 
                               "'" & strMessage & "', " & _ 
                               "'" & strSearchURL & "')" 
              objCmd.Execute 
              If Err.Number <> 0 Then 
                strDBStatus = "Error of formation of transaction." 
              End If 
            End If 
    LogSave(strBuffer) 
        Next 
    Else 
' -=At connection errors=- 
      strMailBody = strMailBody & vbCrLf & Left(UCase(strComputer) & ":                                ", 20) 
      strMailBody = strMailBody & Err.Description & vbCrLf 
      strBuffer =  strBuffer & _ 
            """" & UCase(strComputer) & """" & strReportDelim & _ 
            """" & Err.Description & """" & strReportDelim & _ 
            """" & """" & strReportDelim & _ 
            """" & """" & strReportDelim & _ 
            """" & """" & strReportDelim & _ 
            """" & """" & strReportDelim & _ 
            """" & """" & strReportDelim & _ 
            """" & """" & strReportDelim & _ 
            """" & """" & _ 
            vbCrLf 
  LogSave(strBuffer) 
    End If 
    If strDSN <> "" And Not IsNull(strDSN) Then 
      objCmd.CommandText = "commit transaction EventLogsCollector" 
      Err.Clear 
      objCmd.Execute 
      If Err.Number <> 0 Then 
        strDBStatus = "Error in transaction closing." 
      End If 
    End If 
  End If 
Next 
 
If strDSN <> "" And Not IsNull(strDSN) Then 
  objCmd.CommandText = "begin transaction EventLogsCollector" 
  Err.Clear 
  objCmd.Execute 
  If Err.Number <> 0 Then 
    strDBStatus = "Error of formation of transaction." 
    strDSN = Null 
  Else 
    If DBCleanDays > 0 Then 
        dtmEndDate.SetVarDate CDate(WMItoVBTime(dtmScriptStart)) - DBCleanDays 
        objCmd.CommandText = "delete from " & strTBLName & " where DTEvent < '" & WMItoVBTime(dtmEndDate) & "'" 
        objCmd.Execute 
    End If 
    objCmd.CommandText = "commit transaction EventLogsCollector" 
    Err.Clear 
    objCmd.Execute 
    If Err.Number <> 0 Then 
      strDBStatus = "Error in transaction closing." 
    End If 
  End If 
End If 
 
dtmEndDate.SetVarDate CDate(Date() & " " & Time()) 
strMailBody =  strMailBody & vbCrLf & vbCrLf & _ 
        "=================================================================================" & vbCrLf & _ 
        "Start time: " & WMItoVBTime(dtmScriptStart) & vbCrLf & _ 
        "  End time: " & WMItoVBTime(dtmEndDate) & vbCrLf & _ 
        "   Counter: " & intEventCount & vbCrLf & _ 
        " DB status: " & strDBStatus & vbCrLf & vbCrLf 
 
objReportFile.Close 
MailMSGSend()  ' Sending of the post message 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
objFSO.DeleteFile strReportFile, True  ' Removal of a file with the report 
WScript.Quit 
 
' ==================================================================================================================================== 
' Functions 
' ==================================================================================================================================== 
 
Private  Function LogSave(SV) 
  objReportFile.Write(SV) 
  strBuffer = Empty 
End Function 
 
Private Function MailMSGSend() 
  Set objMessage = CreateObject("CDO.Message")  
  objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
  objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServerName 
  objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSMTPPortNumber 
  objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0 
  objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 
  objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
  objMessage.Configuration.Fields.Update 
 
  objMessage.Subject = strMailSubject 
  objMessage.From = strMailFrom 
  objMessage.To = strMailTo 
  objMessage.TextBody = strMailBody 
  objMessage.AddAttachment strReportFile 
  objMessage.Send 
End Function 
 
Private Function WMItoVBTime(sDMTF) 
        Dim sYear, sMonth, sDate, sHour, sMin, sSec 
        sYear  = Mid(sDMTF, 1,4) 
        sMonth = Mid(sDMTF, 5,2) 
        sDate  = Mid(sDMTF, 7,2) 
        sHour  = Mid(sDMTF, 9,2) 
        sMin   = Mid(sDMTF,11,2) 
        sSec   = Mid(sDMTF,13,2) 
  WMItoVBTime = sYear & "-" & sMonth & "-" & sDate & " " & sHour & ":" & sMin & ":" & sSec 
End Function 