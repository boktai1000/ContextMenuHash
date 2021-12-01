'Get File hash via the right-click menu
'SHA256 hash for the file is copied to the clipboard automatically
'Created: June 4, 2019 by Ramesh Srinivasan - winhelponline.com

Option Explicit
Dim WshShell, sOut, sFileName, sCmd, oExec, strInput
Set WshShell = WScript.CreateObject("WScript.Shell")

If WScript.Arguments.Count = 0 Then
   strInput = InputBox("Type ADD to add the Get File Hash context menu item, or REMOVE to remove the item", "ADD")
   If ucase(strInput) = "ADD" Then
      sCmd = "wscript.exe " & chr(34) & WScript.ScriptFullName & Chr(34) & " " & """" & "%1" & """"
      WshShell.RegWrite "HKCU\Software\Classes\*\shell\gethash\", "Get File Hash", "REG_SZ"
      WshShell.RegWrite "HKCU\Software\Classes\*\shell\gethash\command\", sCmd, "REG_SZ"
      WScript.Quit
   ElseIf ucase(strInput) = "REMOVE" Then
      sCmd = "reg.exe delete HKCU\Software\Classes\*\shell\gethash" & " /f"
      WshShell.Run sCmd, 0
      WScript.Quit
   End If
Else
   sFileName = """" & WScript.Arguments(0) & """"
   sCmd = "cmd.exe /c certutil.exe -hashfile " & sFileName & " SHA256" & _
   " | findstr /v " & chr(34) & "completed successfully" & Chr(34) & " | clip"
   WshShell.Run sCmd, 0
End If