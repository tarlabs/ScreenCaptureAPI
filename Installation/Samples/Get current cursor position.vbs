'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

Dim x,y
x = CLng(0)
y = CLng(0)

'Get current cursor position
oScreenCapture.GetCursorPosition x , y

MsgBox "(X, Y)  - (" & x & ", " & y & ")"

'Destroy the object
Set oScreenCapture = Nothing