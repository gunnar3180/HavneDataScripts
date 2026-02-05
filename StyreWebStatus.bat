@echo off
echo Henter data fra StyreWeb, vent...
"C:\Program Files (x86)\LINQPad5\lprun.exe" WebAutomation.linq
echo(
echo Sjekker konsistens på data...
echo(
echo Status båtplasser:
"C:\Program Files (x86)\LINQPad5\lprun.exe" Havneberegninger.linq
pause
