@echo off

echo Henter data fra StyreWeb, vent...
"C:\Program Files (x86)\LINQPad5\lprun.exe" WebAutomation.linq
echo(

echo Status båtplasser:
set status_file=C:\Users\Solviken\Downloads\StyreWeb Status.txt
"C:\Program Files (x86)\LINQPad5\lprun.exe" Havneberegninger.linq -file "%status_file%"
echo(

echo Styreweb båtplass-status lagret i %status_file%
pause
