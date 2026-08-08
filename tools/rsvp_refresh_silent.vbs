' Auto-refresh RSVP responses without showing a console window.
' Launched by Windows Task Scheduler every N minutes.
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = "C:\Users\user\Desktop\mamoru"
sh.Run """C:\Program Files\nodejs\node.exe"" tools\rsvp_export.mjs", 0, True
