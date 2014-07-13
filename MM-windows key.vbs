@echo off
rem created By Dave Gallop
color 17

title Installing Correct Key
echo This must be Run as an administrator

rem remove current key
slmgr.vbs /upk


Rem install new key
slmgr.vbs /ipk <6T2PM-PQQ66-V888D-VBHJ7-2P76X>

rem auto activate
rem slmgr.vbs /ato

