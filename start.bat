@ECHO OFF
powershell.exe -noprofile -executionpolicy bypass -file \\domain.local\NETLOGON\signature\generate_signature.ps1
cls
