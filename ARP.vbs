
function getIpFromMacaddress
Dim TempFile = "C:\somefile.txt"
Dim macAddress = "00-d0-24-1b-85-76"
if io.file.exists(tempfile) then io.file.delete(tempfile)
Dim shell = createobject("wscipt.shell")
shell.run(string.format("cmd /c arp -a >> " & tempfile)
Dim sr as new io.streamreader(Tempfile)
dim s as string = sr.readtoend
sr.close
Dim lines as string = split(s, vbnewline)
for i as integer=0 to line.length-1
 if line(i).indexof(macAddress) <> -1 then
   return split(line, vbtab)(0).trim
 end if
next
return "Ip not found"
end function