WScript.Echo "Hello"
  WScript.Echo

  Set objNet = WScript.CreateObject("WScript.Network")
  WScript.Echo "Your Computer Name is " & objNet.ComputerName
  WScript.Echo "Your Username is " & objNet.UserName