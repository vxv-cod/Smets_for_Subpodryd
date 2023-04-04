pyinstaller -w -F -i "logo.ico" AutoNameSmetiForSubpodryd.py

xcopy %CD%\*.ico %CD%\dist /H /Y /C /R

xcopy .\dist .\ConsoleApp\ /H /Y /C /R
