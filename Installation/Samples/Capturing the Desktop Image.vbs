'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture the desktop
oScreenCapture.CaptureDesktop "C:\Desktop.png"

'Destroy the object
Set oScreenCapture = Nothing