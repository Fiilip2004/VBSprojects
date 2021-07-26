Const END_OF_STORY = 6
Const MOVE_SELECTION = 0
StrFullname = "path\te.doc"

On Error resume Next
Dim oWdApp : Set oWdApp = GetObject(,"Word.Application")

Set oWdApp = CreateObject("Word.Application")
oWdApp.Visible=False

Err.Clear

Set MyDoc = oWdApp.Documents.Open(StrFullname)
MyDoc.Visible=False

'If IS NOT PAGE IT WILL ADD
if oWdApp.Documents.Count = 0 then
Set MyDoc = oWdApp.Documents.Add()
End If

Set MyDoc = oWdApp.Documents(1)
MyDoc.Activate

Set objSelection = oWdApp.Selection
objSelection.EndKey END_OF_STORY, MOVE_SELECTION
oWdApp.sendkeys "{ENter}"
objSelection.paste
Mydoc.Saveas(StrFullname)
oWdApp.quit