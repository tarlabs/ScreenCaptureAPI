'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Capture the desktop
Set oDesktopTime1  = oScreenCapture.CaptureDesktop ("[object]", "Desktop Screenshot at - " & Now)

'Wait for a 1 sec
WScript.Sleep 1000

'Capture the desktop
Set oDesktopTime2  = oScreenCapture.CaptureDesktop ("[object]", "Desktop Screenshot at - " & Now)

Public Const MERGE_Image1LeftImage2 = 1

'Combine images with a padding of 10 pixels
oScreenCapture.CombineImages oDesktopTime1, oDesktopTime2, "C:\Combined.png", MERGE_Image1LeftImage2, 10

'Dispose the images. Important else memory leak might happen
oDesktopTime1.Dispose()
oDesktopTime2.Dispose()

'Destroy the object
Set oDesktopTime1 = Nothing
Set oDesktopTime2 = Nothing
Set oScreenCapture = Nothing
