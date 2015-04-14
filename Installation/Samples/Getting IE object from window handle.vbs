'SAMPLE QTP script

'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Launch IE
SystemUtil.Run "iexplore.exe"

'Get the window handle of just opened browser
hwnd = Browser("creationtime:=0").GetROProperty("hwnd")

'Get the IE object from browser window handle
Set oIE = oScreenCapture.GetIEObjectFromHWND(hwnd)

If Not oIE Is Nothing Then
	oIE.Navigate2 "www.knowledgeinbox.com"
End If

'Destroy the object
Set oScreenCapture = Nothing

