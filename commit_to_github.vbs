Option Explicit

Dim shell, commitMsg, confirm, cmd

Set shell = CreateObject("WScript.Shell")

commitMsg = InputBox("Vnesi git commit sporočilo:", "Git Commit")

If commitMsg = "" Then
    WScript.Quit
End If

confirm = InputBox("Srček, ali res hočeš commitati in potisniti spremembe? (Y/N):", "Potrditev")

If UCase(confirm) <> "Y" Then
    MsgBox "Operacija preklicana. Ni bilo git commita.", vbInformation, "Preklicano"
    WScript.Quit
End If

cmd = "cmd /k git add . && git commit -m """ & Replace(commitMsg, """", "\""") & """ && git push -u origin master && echo. && echo V redu, commit končan! Fih && curl https://ascii.live/rick"

shell.Run cmd, 1, False