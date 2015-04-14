'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture desktop
oScreenCapture.CaptureDesktop "C:\Image1.png"

'Wait for a min
WScript.Sleep 60000

'Capture desktop
oScreenCapture.CaptureDesktop "C:\Image2.png"

'Get count of pixels which are different in both images
PixelCountDiff = oScreenCapture.CompareImages ("C:\Image1.png", "C:\Image2.png", "[PixelDiffCount]")

'Get percentage of pixels which are different in both images
PixelDiffPerc = oScreenCapture.CompareImages ("C:\Image1.png", "C:\Image2.png", "[PixelDiffPerc]")

MsgBox PixelCountDiff 
MsgBox PixelDiffPerc 

'Save the difference image
Call oScreenCapture.CompareImages ("C:\Image1.png", "C:\Image2.png", "C:\Difference.png")


'Destroy the object
Set oScreenCapture = Nothing
