@echo off
%systemroot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /codebase %~dp0vRedis.dll
iisreset
pause