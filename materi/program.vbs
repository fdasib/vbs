Set fso = CreateObject("Scripting.FileSystemObject")

set cmd = CreateObject("wscript.shell")

Set fso = CreateObject("Scripting.FileSystemObject")
cd = Replace(WScript.ScriptFullName, WScript.ScriptName, "coba.txt")

set file = fso.OpenTextFile(cd, 1)
text = file.ReadAll
file.Close

coba = InputBox("pilih opsi : " & vblf & "1. Enkripsi" & vblf & "2. Dekripsi")
MsgBox coba

For i = 1 To Len(text)
    intASCII = Asc(Mid(text, i, 1))
    if coba = "1" then
        if i = 1 then
            intEncrypted = intASCII + 10 'contoh nilai konstanta: 10
        elseif i = 2 then
            intEncrypted = intASCII + 150 'contoh nilai konstanta: 10
        elseif i = 3 then
            intEncrypted = intASCII + 55 'contoh nilai konstanta: 10
        elseif i = 8 then
            intEncrypted = intASCII + 99 'contoh nilai konstanta: 10
        Else
            intEncrypted = intASCII + 100 'contoh nilai konstanta: 10
        end if
    else
        if i = 1 then
            intEncrypted = intASCII - 10 'contoh nilai konstanta: 10
        elseif i = 2 then
            intEncrypted = intASCII - 150 'contoh nilai konstanta: 10
        elseif i = 3 then
            intEncrypted = intASCII - 55 'contoh nilai konstanta: 10
        elseif i = 8 then
            intEncrypted = intASCII - 99 'contoh nilai konstanta: 10
        Else
            intEncrypted = intASCII - 100 'contoh nilai konstanta: 10
        end if
    end if
    strEncrypted = strEncrypted & Chr(intEncrypted)
Next

set file = fso.OpenTextFile(cd, 2)

file.Write strEncrypted
file.Close

WScript.Echo "Complete."