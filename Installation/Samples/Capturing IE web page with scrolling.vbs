'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Open a IE window
Set oIE = CreateObject("InternetExplorer.Application")

'Make the IE window visible
oIE.Visible = True

'Navigate to a website
oIE.Navigate2 "http://www.knowledgeinbox.com"

'Wait for the page to load
While oIE.Busy
Wend

'Capture the web page from starting
oScreenCapture.CaptureIEFromCurrentPos = False

'Set the font details
oScreenCapture.TextFontName = "Arial"
oScreenCapture.TextFontSize = 8

'Keep height of text as 30 px
oScreenCapture.TextHeight = 30

'Add text to both top and bottom
oScreenCapture.TextPos = 2

'Add a delay of 50msec before every capture
oScreencapture.WaitBeforeCapture = 50

'Capture the web page with vertical and horizontal scroll enabled
oScreenCapture.CaptureIE oIE.HWND, "C:\IE Srcolling.png", "Scrolling Capture", True, True

'Destory the object
Set oScreenCapture = Nothing
