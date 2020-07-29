C:\Windows\SysWOW64\regsvr32.exe /s src\bin\x86\Debug\netcoreapp3.1\PowerPointAddin.comhost.dll

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointExtensibility.PowerPointAddin" /f /v LoadBehavior /t REG_DWORD /d 3
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointExtensibility.PowerPointAddin" /f /v FriendlyName /t REG_SZ /d "PowerPoint Addin (.NET Core 3.1)"
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointExtensibility.PowerPointAddin" /f /v Description /t REG_SZ /d "Sample addin running using .NET Core 3.1"
