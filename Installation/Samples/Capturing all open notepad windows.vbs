'Create the screen capture object
Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")

'Get all windows with notepad title
Set allNotePads = oScreenCapture.FindWindowLike ("*notepad*")

'Capture each notepad window
For Each oNotePad In allNotepads
	oScreenCapture.CaptureWindow oNotepad.hwnd,"C:\" & oNotepad.Title & ".png"
Next

'Destroy the object
Set oScreenCapture = Nothing
