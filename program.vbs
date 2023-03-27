Set fso = CreateObject("Scripting.FileSystemObject")

set cmd = CreateObject("wscript.shell")

Set fso = CreateObject("Scripting.FileSystemObject")
cd = Replace(WScript.ScriptFullName, WScript.ScriptName, "img.jpg")

set file = fso.OpenTextFile(cd, 1)
text = file.ReadAll
file.Close

' ganti = Replace(text, "ÿ", f)
MsgBox text
ganti = Replace(text, "f", "ÿ")
MsgBox ganti

MsgBox "ok"

set file = fso.OpenTextFile(cd, 2)

file.Write ganti
file.Close

WScript.Echo "Complete."