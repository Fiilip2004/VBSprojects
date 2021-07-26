Set oWordBasic = CreateObject("Word.Basic")

'FULL SCREEN
oWordBasic.SendKeys "{prtsc}"

'ONLY selected wondow
'oWordBasic.SendKeys "%{prtsc}"

oWordBasic.FileQuit
Set oWordBasic = nothing ' clean up'
