'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture the desktop rectangle
oScreenCapture.CaptureDesktopRect 0, 0, 200, 200, "C:\DesktopRect.png", "Desktop (0,0) - (200,200)"

'Destroy the object
Set oScreenCapture = Nothing

