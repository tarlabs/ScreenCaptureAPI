'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Get a new blank bitmap
Set oNewBitmap = oScreenCapture.GetNewBitmap(100,1)

'Add text to the to the bitmap
Set oNewBitmap = oScreenCapture.AddTextToImage(oNewBitmap,"This is a test")

'Save the bitmap to file
oScreenCapture.SaveImageToFile oNewBitmap, "C:\users\tarunlalwani\desktop\newbitmap.png"

'Dispose the new bitmap
oNewBitmap.dispose

'Destroy the objects created
Set oNewBitmap = Nothing
Set oScreenCapture = Nothing