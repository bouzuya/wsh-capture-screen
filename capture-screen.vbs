Option Explicit

Const INTERVAL_SEC = 10
Const WINSHOT_PATH = "C:\Users\user\Documents\ws153a\winshot.exe"
Const WINSHOT_OPTS = "-J -D -X"
Const FFMPEG_PATH = "C:\Users\user\Documents\ffmpeg\bin\ffmpeg.exe"
Const CAPTURE_DIR = "C:\Users\user\Pictures\winshot\"

Dim objShell, objFso
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFso = WScript.CreateObject("Scripting.FileSystemObject")

' check *.LCK
While objFso.FileExists(CAPTURE_DIR & "capture-screen.lck")
    WScript.Sleep(INTERVAL_SEC * 1000)
Wend
Call objFso.CreateTextFile(CAPTURE_DIR & "capture-screen.lck")

' save as *.JPG
Dim dtmStart, dtmNow
dtmStart = Now()
dtmNow = Now()
While DateDiff("s", dtmStart, dtmNow) < 24 * 60 * 60
    Call objShell.Run("""" & WINSHOT_PATH & """ " & WINSHOT_OPTS, 0, True)
    WScript.Sleep(INTERVAL_SEC * 1000)
    dtmNow = Now()
Wend

' convert to *.AVI
Dim strMoviePath
strMoviePath = CAPTURE_DIR & ToISO8601Date(dtmStart) & ".avi"
WScript.Echo("""" & FFMPEG_PATH & """ -i """ & CAPTURE_DIR & "WS%06d.JPG"" """ & strMoviePath & """")
Call objShell.Run("""" & FFMPEG_PATH & """ -i """ & CAPTURE_DIR & "WS%06d.JPG"" """ & strMoviePath & """", 0, True)

' delete *.JPG
Call objFso.DeleteFile(CAPTURE_DIR & "WS*.JPG")

' delete *.LCK
Call objFso.DeleteFile(CAPTURE_DIR & "capture-screen.lck")

