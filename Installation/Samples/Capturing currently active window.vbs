'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture the active window
oScreenCapture.CaptureActiveWindow "C:\ActiveWindow.png"

'Destroy the object
Set oScreenCapture = Nothing
