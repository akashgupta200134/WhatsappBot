#NoTrayIcon
SetTitleMatchMode, 2

; Read the image path from temp file
FileRead, imagePath, C:\Users\Akash\Desktop\WhatsappBot\ahk_path.txt
StringReplace, imagePath, imagePath, `n, , All
StringReplace, imagePath, imagePath, `r, , All

WinWait, Open,, 15
if ErrorLevel
    ExitApp

WinActivate, Open
WinWaitActive, Open,, 5
Sleep, 500

; Clear the filename box and type the full path
ControlSetText, Edit1, %imagePath%, Open
Sleep, 400

; Click the Open button
ControlClick, Button1, Open
Sleep, 300