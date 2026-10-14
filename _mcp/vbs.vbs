Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\LabUser\zava-contract-management"

' Launch UI process in the background
WshShell.Run "cmd /c npm run ui", 0, False

' Launch MCP process in the background
WshShell.Run "cmd /c npm run mcp", 0, False

Dim msg
msg = "Services have started successfully!" & vbCrLf & vbCrLf & _
      "• UI: npm run ui" & vbCrLf & _
      "• MCP: npm run mcp"

WshShell.Popup msg, 4, "Services Started", 64