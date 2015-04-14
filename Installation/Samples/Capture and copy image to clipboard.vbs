'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture the desktop window to clipboard
oScreenCapture.CaptureDesktop "[Clipboard]"

'Open a word document and paste the image in the document
Set wrdApp = CreateObject ("Word.Application")
wrdApp.Visible = True
Set oWordDoc = wrdApp.Documents.Add
oWordDoc.Range.Paste

Set oScreenCapture = Nothing
