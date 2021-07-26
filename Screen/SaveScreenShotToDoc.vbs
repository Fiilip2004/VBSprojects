Const END_OF_STORY = 6
Const MOVE_SELECTION = 0
count = 1
Docopen = 0

Set oWordBasic = CreateObject("Word.Basic")
oWordBasic.SendKeys "%{prtsc}"
oWordBasic.FileQuit
Set oWordBasic = nothing ' clean up'

StrFullname = "C:\path\te.doc"

On Error resume Next
Dim oWdApp : Set oWdApp = GetObject(,"Word.Application")

Msgbox(err.description)
If err.Number <> 0 then
Set oWdApp = CreateObject("Word.Application")
Set oWordBasic = CreateObject("Word.Basic")
End If
oWdApp.Visible = true
Err.Clear

Do Until count > oWdApp.Documents.Count OR oWdApp.Documents.Count = 0
Msgbox(oWdApp.Documents(count).Name)
if StrComp(oWdApp.Documents(count).FullName,StrFullname,1) = 0 then
Set MyDoc = oWdApp.Documents(count)
Docopen = 1
MyDoc.Activate
Msgbox("Doc Open")
Exit Do
End If
count = count + 1
Loop

if Docopen = 0 then
Msgbox("Doc not open")
Set MyDoc = oWdApp.Documents.Open(StrFullname)
If err.number = 5273 then
Msgbox("Given directory " & Foldername & " does not exist, Create it First )" )
End If
If err.number = 5174 then
MsgBox("a")
Set MyDoc = oWdApp.Documents.Add()
Mydoc.Saveas(StrFullname)
End If
End If


Msgbox("Doc open")
Set oWdApp.Visible=False

Set objSelection = oWdApp.Selection
objSelection.EndKey END_OF_STORY, MOVE_SELECTION
oWdApp.sendkeys "{ENter}"
objSelection.paste
Mydoc.save

Set oWdApp.Visible=True

Set Mydoc = Nothing
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate "Microsoft Internet Explorer"

oWdApp = nothing