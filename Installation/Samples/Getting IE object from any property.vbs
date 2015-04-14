'SAMPLE QTP script


'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")


'Launch IE
SystemUtil.Run "iexplore.exe", "http://www.google.co.in"


'Wait for browser to load
Browser("creationtime:=0").Sync


'Get the window handle of just opened browser
hwnd = Browser("creationtime:=0").GetROProperty("hwnd")


'Various ways to get the IE object using different properties
Set oIE = oScreenCapture.GetIEObjectFromProperty("hwnd", hwnd)
Set oIE = oScreenCapture.GetIEObjectFromProperty("document", Browser("creationtime:=0").object.document)
Set oIE = oScreenCapture.GetIEObjectFromProperty("document", Browser("creationtime:=0").Page("micclass:=Page").object)
Set oIE = oScreenCapture.GetIEObjectFromProperty("document.title", "Google")
Set oIE = oScreenCapture.GetIEObjectFromProperty("document.location", "http://www.google.co.in")




If Not oIE Is Nothing Then
      oIE.Navigate2 "www.knowledgeinbox.com"
End If


'Destroy the object
Set oScreenCapture = Nothing