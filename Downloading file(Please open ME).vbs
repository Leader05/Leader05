Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad", 9
WScript.Sleep 500 ' Notepad Zeit zum Laden geben
For i = 1 To 8000
  WshShell.SendKeys "YOU SHOULD NOT HAVE DONE THAT"
  WshShell.SendKeys "{ENTER}"
  Next