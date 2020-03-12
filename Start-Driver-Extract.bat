powershell -noprofile -command "&{ start-process powershell -ArgumentList ' -ExecutionPolicy bypass -noprofile -file "%~dp0\Driver-Extract.ps1"' -verb RunAs}"

pause