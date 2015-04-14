'SAMPLE QTP script


'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")


Set oIE = CreateObject("InternetExplorer.Application")


oIE.Navigate2 "http://www.google.co.in"
oIE.Visible = True
oIE.FullScreen = False


While oIE.Busy


Wend


'CaptureSourceObjectOnly = False, will capture whole document
Call oScreenCapture.CaptureIEObject(oIE.Document.forms.f.btnG, "C:\test\fullCapture.jpg", "Full capture", True,True,False)


'CaptureSourceObjectOnly = True, will capture the specified object only
Call oScreenCapture.CaptureIEObject(oIE.Document.forms.f.btnG, "C:\test\specificObject.jpg", "Specific Object", True,True,True)


'In QTP directly capture the object
'Call oScreenCapture.CaptureIEObject(Browser("Google").Page("Google").WebEdit("q"), "C:\test\specificObject.jpg", "Specific Object", True,True,True)


oIE.Quit




Set oIE = Nothing
Set oScreenCapture = Nothing