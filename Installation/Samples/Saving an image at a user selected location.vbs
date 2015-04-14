'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture the desktop window and save it at user selected location
oScreenCapture.CaptureDesktop "[User]"

'Destroy the object
Set oScreenCapture = Nothing

