'Launch IE 
oIE = CreateObject("InternetExplorer.Application")


'Browse to page with frames
oIE.Navigate2("http://www.mountaindragon.com/html/2frames.htm")
oIE.Visible = True
oIE.FullScreen = False


While oIE.Busy


End While


oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")


'Capture all documents with frames also
oScreenCapture.CaptureIEFrames(oIE, "c:\temp\IEFrames.jpg")


oScreenCapture = Nothing