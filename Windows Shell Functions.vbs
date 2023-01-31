Function GetCurrentUsername()
  Set oWshShell = CreateObject("WScript.Shell")
  GetCurrentUsername = oWshShell.ExpandEnvironmentStrings("%UserName%")
End Function
