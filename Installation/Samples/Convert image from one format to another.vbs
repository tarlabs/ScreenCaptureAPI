'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture desktop
oScreenCapture.CaptureDesktop "C:\ImageBMP.bmp"

'Convert the image to PNG format
oScreenCapture.ConvertImage "C:\ImageBMP.bmp","C:\ImagePNG.PNG"

'Destroy the object
Set oScreenCapture = Nothing