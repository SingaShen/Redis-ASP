@echo off
%systemroot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /u /codebase %~dp0vRedis.dll
iisreset
pause