Option Compare Text

Imports System.IO
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Windows.Forms
Imports System.Runtime.InteropServices


<Guid("E7852903-FA4C-4a84-8832-21DC7A9A150D"), _
InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> _
Public Interface _ScreenCapture
    <DispId(30)> Function CompareImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal ScaleImages As Boolean = True)
    <DispId(1)> Function SaveImageToFile(ByVal oImage As Object, ByVal destFile As String) As Object
    <DispId(2)> Function AddTextToImage(ByVal oImage As Object, ByVal Text As String)
    <DispId(3)> Function CaptureActiveWindow(ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "")
    <DispId(4)> Function CaptureDesktop(ByVal destination As String, Optional ByVal Text As String = "")
    <DispId(5)> Function CaptureIE(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False)
    <DispId(6)> Function CaptureDesktopRect(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal destination As String, Optional ByVal text As String = "")
    <DispId(7)> Function CaptureWindow(ByVal hwnd As Long, ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = True, Optional ByVal HorizontalScroll As Boolean = True)
    <DispId(8)> Function CaptureWindowRect(ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal destination As String, Optional ByVal Text As String = "")
    <DispId(9)> Sub ConvertImage(ByVal sourceFile As String, ByVal destFile As String, Optional ByVal OverWriteDestination As Boolean = True)
    <DispId(10)> Function GetIEObjectFromHWND(ByVal hwnd As Long) As Object
    <DispId(11)> Function GetRootWindow(ByVal hwnd As Long) As Long
    <DispId(12)> Function CombineImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal Position As Integer = 3, Optional ByVal Padding As Integer = 0)
    <DispId(13)> Function GetNewBitmap(ByVal Width As Integer, ByVal Height As Integer) As Bitmap
    <DispId(14)> Function FindWindowLike(Optional ByVal WindowTitle As String = "*", Optional ByVal WindowClass As String = "*")
    <DispId(15)> Sub GetCursorPosition(ByRef X As Object, ByRef Y As Object)
    <DispId(16)> Property TextPos() As Long
    <DispId(17)> Property TextFontName() As String
    <DispId(18)> Property TextFontSize() As Integer
    <DispId(19)> Property ActivateWindowOnCapture() As Boolean
    <DispId(20)> Property WaitBeforeCapture() As Long
    <DispId(21)> Property CaptureIEFromCurrentPos() As Boolean
    <DispId(22)> Property TextHeight() As Long
End Interface


<Guid("EA4215C4-3DF4-4b3d-BE76-CB7183B48605"), _
ClassInterface(ClassInterfaceType.None), _
ProgId("KnowledgeInbox.ScreenCapture")> _
Public Class ScreenCapture
    Implements _ScreenCapture

    Public Const MERGE_Image1RightImage2 = 0
    Public Const MERGE_Image1LeftImage2 = 1
    Public Const MERGE_Image1BottomImage2 = 2
    Public Const MERGE_Image1TopImage2 = 3

    Private mActivateWindowOnCapture As Boolean = True
    Private mWaitBeforeCapture As Long = 300
    Private mCaptureIEFromCurrentPos As Boolean = False
    Private mTextHeight As Long = 20
    Private mTextPos As Integer = 2 '0-Top, 1-bottom, 2-Both
    Private mTextFontName As String = "Verdana"
    Private mTextFontSize As Integer = 10

    ''' <summary>
    ''' ActivateWindowOnCapture if set to True will activate and bring the window to top
    ''' before capturing its image
    ''' </summary>
    ''' <value>True/False</value>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Public Property ActivateWindowOnCapture() As Boolean Implements _ScreenCapture.ActivateWindowOnCapture
        Get
            ActivateWindowOnCapture = mActivateWindowOnCapture
        End Get

        Set(ByVal NewValue As Boolean)
            mActivateWindowOnCapture = NewValue
        End Set
    End Property


    ''' <summary>
    ''' WaitBeforeCapture specifies the delay to be used between activating and capturing a window
    ''' </summary>
    ''' <value>Default: 300msec</value>
    ''' <returns>waiting time in msec</returns>
    ''' <remarks></remarks>
    Public Property WaitBeforeCapture() As Long Implements _ScreenCapture.WaitBeforeCapture
        Get
            WaitBeforeCapture = mWaitBeforeCapture
        End Get
        Set(ByVal value As Long)
            mWaitBeforeCapture = value
        End Set
    End Property


    ''' <summary>
    ''' CaptureIEFromCurrentPos if set to False then it will scroll the web page to top
    ''' before doing a scrolling capture
    ''' </summary>
    ''' <value>True/False. Default: False</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CaptureIEFromCurrentPos() As Boolean Implements _ScreenCapture.CaptureIEFromCurrentPos
        Get
            CaptureIEFromCurrentPos = mCaptureIEFromCurrentPos
        End Get
        Set(ByVal value As Boolean)
            mCaptureIEFromCurrentPos = value
        End Set
    End Property

    ''' <summary>
    ''' Defines the height of text that is added to the image when using AddTextToImage function
    ''' </summary>
    ''' <value>Default value: 20px</value>
    ''' <returns>Hieght of the text to be added using AddTextToImage</returns>
    ''' <remarks></remarks>
    Public Property TextHeight() As Long Implements _ScreenCapture.TextHeight
        Get
            TextHeight = mTextHeight
        End Get
        Set(ByVal value As Long)
            mTextHeight = value
        End Set
    End Property

    ''' <summary>
    ''' Position of the image to be added to the text
    ''' </summary>
    ''' <value>
    ''' 1 - Add text to the top
    ''' 2 - Add text to the bottom (Default)
    ''' 3 - Add text to both top and bottom
    ''' </value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TextPos() As Long Implements _ScreenCapture.TextPos
        Get
            TextPos = mTextPos
        End Get
        Set(ByVal value As Long)
            If value > 0 And value < 4 Then
                mTextPos = value
            Else
                Err.Raise(vbObjectError, "ScreenCapture.TextPos", "TextPos can only be set to 1-Top, 2-Bottom, 3-Both")
            End If
        End Set
    End Property

    Public Property TextFontName() As String Implements _ScreenCapture.TextFontName
        Get
            TextFontName = mTextFontName
        End Get
        Set(ByVal value As String)
            mTextFontName = value
        End Set
    End Property
    Public Property TextFontSize() As Integer Implements _ScreenCapture.TextFontSize
        Get
            TextFontSize = mTextFontSize
        End Get
        Set(ByVal value As Integer)
            mTextFontSize = value
        End Set
    End Property

    Structure PartialCapture
        Dim x As Integer, y As Integer
        Dim width As Integer, height As Integer
        Dim oImage As Bitmap
    End Structure

    Structure CompleteCapture
        Dim width As Integer, height As Integer
        Dim PartialCaptures() As PartialCapture
    End Structure


    ''' <summary>
    ''' Compare two images and get the difference (XOR) image. Also can get the pixel
    ''' Differnce count or difference percentage using [PixelDiffCount] or [PixelDiffPerc]
    ''' as destination value. See help for more details
    ''' </summary>
    ''' <param name="SrcImage1">File Path to image 1 or the .NET Bitmap object</param>
    ''' <param name="SrcImage2">File Path to image 2 or the .NET bitmap object</param>
    ''' <param name="Destination">Destination file name</param>
    ''' <param name="ScaleImages">ScaleImages in case of different size images being compared</param>
    ''' <returns>
    ''' -1 if ScaleImages=False and Images are of differnt size
    ''' .NET Bitmap object of Difference image if destination is "[object]"
    ''' Number of Pixel differences if destination is "[PixelDiffCount]"
    ''' Percentage of Pixel difference if destination is "[PixelDiffPerc]"
    ''' </returns>
    ''' <remarks>
    ''' Image comparison should be done with images taken by the same method (ex - Paint
    ''' or API from this library). Comparison of JPG should be avoided as the pixel difference
    ''' count would be high because of quality difference. Instead PNG, BMPP formats should be used.
    ''' </remarks>
    Public Function CompareImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal ScaleImages As Boolean = True) Implements _ScreenCapture.CompareImages
        Dim bmp1 As Bitmap ' = Bitmap.FromFile(SrcImage1)
        Dim bmp2 As Bitmap ' = Bitmap.FromFile(SrcImage2)

        If SrcImage1.GetType().Name = "String" Then
            bmp1 = Image.FromFile(SrcImage1)
        Else
            bmp1 = SrcImage1.clone
        End If

        If SrcImage2.GetType().Name = "String" Then
            bmp2 = Image.FromFile(SrcImage2)
        Else
            bmp2 = SrcImage2.Clone
        End If

        If (bmp1.Size <> bmp2.Size) And Not ScaleImages Then
            CompareImages = -1
        Else
            If (bmp1.Size <> bmp2.Size) Then
                If (bmp1.Width > bmp2.Width And bmp1.Height > bmp2.Height) Or (bmp1.Width + bmp1.Height > bmp2.Width + bmp2.Height) Then
                    bmp1 = New Bitmap(bmp1, bmp2.Width, bmp2.Height)
                Else
                    bmp2 = New Bitmap(bmp2, bmp1.Width, bmp1.Height)
                End If
            End If
        End If

        Dim bmp3 As Bitmap = New Bitmap(bmp1.Width, bmp1.Height)

        Dim bmpData1 As Imaging.BitmapData = bmp1.LockBits(New Rectangle(0, 0, bmp1.Width, bmp1.Height), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Dim bmpData2 As Imaging.BitmapData = bmp2.LockBits(New Rectangle(0, 0, bmp2.Width, bmp2.Height), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Dim bmpData3 As Imaging.BitmapData = bmp3.LockBits(New Rectangle(0, 0, bmp3.Width, bmp1.Height), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Dim bytes As Integer = bmpData1.Stride * bmp1.Height

        Dim buffer1(bytes - 1), buffer2(bytes - 1), buffer3(bytes - 1) As Byte

        System.Runtime.InteropServices.Marshal.Copy(bmpData1.Scan0, buffer1, 0, bytes)
        System.Runtime.InteropServices.Marshal.Copy(bmpData2.Scan0, buffer2, 0, bytes)

        For i As Integer = 0 To buffer1.GetUpperBound(0)
            buffer3(i) = buffer1(i) Xor buffer2(i)
        Next

        Dim MisMatchCount As Long
        MisMatchCount = 0

        System.Runtime.InteropServices.Marshal.Copy(buffer3, 0, bmpData3.Scan0, bytes)

        bmp1.UnlockBits(bmpData1)
        bmp2.UnlockBits(bmpData2)
        bmp3.UnlockBits(bmpData3)

        bmp1.Dispose()
        bmp1 = Nothing

        bmp2.Dispose()
        bmp2 = Nothing

        Select Case Destination.ToLower
            Case "[pixeldiffcount]", "[pixeldiffperc]"
                Dim i As Long
                Dim iCount As Long

                iCount = UBound(buffer3) - 1
                MisMatchCount = 0

                For i = 0 To iCount Step 3
                    If buffer3(i) <> 0 Or buffer3(i + 1) <> 0 Or buffer3(i + 2) <> 0 Then
                        MisMatchCount = MisMatchCount + 1
                    End If
                Next
                CompareImages = MisMatchCount
                If Destination.ToLower = "[pixeldiffperc]" Then
                    CompareImages = (MisMatchCount / (bmp3.Width * bmp3.Height)) * 100
                End If
            Case Else
                CompareImages = SaveImageToFile(bmp3, Destination)
        End Select

        bmp3.Dispose()
        bmp3 = Nothing
    End Function

    ''' <summary>
    ''' Add text to a give .NET bitmap object
    ''' </summary>
    ''' <param name="oImage">.NET bitmap object</param>
    ''' <param name="Text">Text to be added</param>
    ''' <returns>Image object with text added</returns>
    ''' <remarks>To change Text Height, Font, Size and Position see properties 
    ''' TextHeight, TextFontName, TextFontSize and TextPos
    ''' </remarks>
    Public Function AddTextToImage(ByVal oImage As Object, ByVal Text As String) Implements _ScreenCapture.AddTextToImage
        Select Case TypeName(oImage)
            Case "Bitmap", "Image"
            Case Else
                Err.Raise(vbObjectError + 1, "AddTextToImage", "Invalid parameter, object needs to be of type bitmap or image")
        End Select

        On Error GoTo errHandler

1:      If Text.Trim = "" Then
2:          AddTextToImage = oImage
3:          Exit Function
4:      End If

5:      Dim AddTextHeight As Integer
6:      If TextPos = 1 Or TextPos = 2 Then
7:          AddTextHeight = TextHeight
8:      Else
9:          AddTextHeight = 2 * TextHeight
10:     End If

11:     Dim oNewImage As Image = New Bitmap(CInt(oImage.Width), CInt(oImage.Height + AddTextHeight))
12:     Dim oGraphics As Graphics = Graphics.FromImage(oNewImage)

13:     Dim oFont As Font = New Font(TextFontName, TextFontSize)
14:     If TextPos = 1 Then
15:         oGraphics.FillRectangle(Brushes.White, 0, 0, oImage.Width, TextHeight)
16:         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, 2)
17:         oGraphics.DrawImage(oImage, 0, TextHeight) '+ oImage.Height)
18:     ElseIf TextPos = 2 Then
19:         oGraphics.FillRectangle(Brushes.White, 0, oImage.Height, oImage.Width, oImage.Height + TextHeight)

20:         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, oImage.Height + 2)
21:         oGraphics.DrawImage(oImage, 0, 0)
22:     Else
23:         oGraphics.FillRectangle(Brushes.White, 0, 0, oImage.Width, TextHeight)
24:         oGraphics.FillRectangle(Brushes.White, 0, oImage.Height + TextHeight, oImage.Width, oImage.Height + TextHeight)

25:         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, 2)
26:         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, 2 + oImage.Height + TextHeight)
27:         oGraphics.DrawImage(oImage, 0, TextHeight)
28:     End If

29:     oGraphics.Flush()
30:     oGraphics.Dispose()
31:     oGraphics = Nothing

32:     AddTextToImage = oNewImage
        Exit Function
errHandler:
        MsgBox(Err.Erl)
        MsgBox(Err.Description)
        MsgBox(Err.GetException().Message)
    End Function

    ''' <summary>
    ''' Capture the currently active window.
    ''' </summary>
    ''' <param name="destination">
    ''' Path to the destination file name. 
    ''' </param>
    ''' <param name="Client">
    ''' Optional. If set to True then the Title bar of the window is not captured. Default is False
    ''' </param>
    ''' <param name="Text">
    ''' Optional. Text to be added to the image.
    ''' </param>
    ''' <returns>
    ''' Nothing of file name is given. .NET bitmap object if destination is "[object]"
    ''' </returns>
    ''' <remarks>
    ''' Destination can have special keyword values
    ''' [object] - Return the .NET bitmap object
    ''' [clipboard] - Copy the image to clipboard
    ''' [user] - Ask user the path of the file to be saved.
    ''' </remarks>
    Public Function CaptureActiveWindow(ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "") Implements _ScreenCapture.CaptureActiveWindow
        Dim oBitmap As Bitmap
        oBitmap = GetImageFromHwnd(GetForegroundWindow(), Client)
        oBitmap = AddTextToImage(oBitmap, Text)
        CaptureActiveWindow = SaveImageToFile(oBitmap, destination)
        oBitmap.Dispose()
        oBitmap = Nothing
    End Function

    ''' <summary>
    ''' Capture complete desktop
    ''' </summary>
    ''' <param name="destination">Path to the destination file name. </param>
    ''' <param name="Text">Optional. Text to be added to the image.</param>
    ''' <returns>Nothing of file name is given. .NET bitmap object if destination is "[object]"</returns>
    ''' <remarks>
    ''' Destination can have special keyword values
    ''' [object] - Return the .NET bitmap object
    ''' [clipboard] - Copy the image to clipboard
    ''' [user] - Ask user the path of the file to be saved.
    ''' </remarks>

    Public Function CaptureDesktop(ByVal destination As String, Optional ByVal Text As String = "") Implements _ScreenCapture.CaptureDesktop
        Dim oBitmap As Bitmap
        oBitmap = GetImageFromHwnd(GetDesktopWindow(), False)
        oBitmap = AddTextToImage(oBitmap, Text)
        CaptureDesktop = SaveImageToFile(oBitmap, destination)
        oBitmap.Dispose()
        oBitmap = Nothing
    End Function

    ''' <summary>
    ''' Capture window with a given handle
    ''' </summary>
    ''' <param name="hwnd">Handle of the window to be captured</param>
    ''' <param name="destination">Path to the destination file name. </param>
    ''' <param name="Client">Optional. If set to True then the Title bar of the window is not captured. Default is False</param>
    ''' <param name="Text">Optional. Text to be added to the image.</param>
    ''' <param name="VerticalScroll">Reserved for Future use</param>
    ''' <param name="HorizontalScroll">Reserved for Future use</param>
    ''' <returns>Nothing of file name is given. .NET bitmap object if destination is "[object]"</returns>
    ''' <remarks>
    ''' Destination can have special keyword values
    ''' [object] - Return the .NET bitmap object
    ''' [clipboard] - Copy the image to clipboard
    ''' [user] - Ask user the path of the file to be saved.
    ''' </remarks>

    Public Function CaptureWindow(ByVal hwnd As Long, ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = True, Optional ByVal HorizontalScroll As Boolean = True) Implements _ScreenCapture.CaptureWindow
        Dim oBitmap As Bitmap
        oBitmap = GetImageFromHwnd(hwnd, Client, VerticalScroll, HorizontalScroll)
        oBitmap = AddTextToImage(oBitmap, Text)
        CaptureWindow = SaveImageToFile(oBitmap, destination)
        oBitmap.Dispose()
        oBitmap = Nothing
    End Function

    ''' <summary>
    ''' Capture IE web page with options for scrolling.
    ''' </summary>
    ''' <param name="source">
    ''' This can object of Type InternetExplorer.Application
    ''' Or hwnd of the web browser
    ''' QTP Page or Browser object
    ''' </param>
    ''' <param name="destination"></param>
    ''' <param name="Text"></param>
    ''' <param name="VerticalScroll"></param>
    ''' <param name="HorizontalScroll"></param>
    ''' <returns>Nothing of file name is given. .NET bitmap object if destination is "[object]"</returns>
    ''' <remarks>
    ''' Destination can have special keyword values
    ''' [object] - Return the .NET bitmap object
    ''' [clipboard] - Copy the image to clipboard
    ''' [user] - Ask user the path of the file to be saved.
    ''' </remarks>
    Public Function CaptureIE(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False) Implements _ScreenCapture.CaptureIE
        Dim IE As Object
        IE = GetIEFromSource(source)
        CaptureIE = False

        If IE Is Nothing Then
            If IsNumeric(source) Then
                CaptureIE = SaveImageToFile(GetImageFromHwnd(source), destination)
                Exit Function
            Else
                CaptureIE = False
                Exit Function
            End If
        ElseIf TypeName(IE.document).ToLower() <> "htmldocument" And TypeName(IE.document).ToLower() <> "htmldocumentclass" Then
            If IsNumeric(source) Then
                CaptureIE = SaveImageToFile(GetImageFromHwnd(source), destination)
                Exit Function
            Else
                CaptureIE = False
                Exit Function
            End If
        End If

        Dim CompleteCapture As CompleteCapture
        'CaptureIEFromCurrentPos = True
        Dim oldActivateMode As Boolean, oldWait As Long
        oldWait = WaitBeforeCapture

        oldActivateMode = ActivateWindowOnCapture

        If oldActivateMode Then
            ActivateWindow(IE.hwnd)
            Sleep(oldWait)
        End If

        ActivateWindowOnCapture = False
        WaitBeforeCapture = 0

        CompleteCapture = CreateIEPartialCaptures(IE, VerticalScroll, HorizontalScroll)

        ActivateWindowOnCapture = oldActivateMode
        WaitBeforeCapture = oldWait

        'Combine all images into one
        Dim CombinedImage As Bitmap, oGraphics As Graphics
        CombinedImage = New Bitmap(CompleteCapture.width, CompleteCapture.height)

        oGraphics = Graphics.FromImage(CombinedImage)

        Dim Capture As PartialCapture

        For Each Capture In CompleteCapture.PartialCaptures
            oGraphics.DrawImage(Capture.oImage, New Rectangle(Capture.x, Capture.y, Capture.width, Capture.height), New Rectangle(2, 2, Capture.width, Capture.height), GraphicsUnit.Pixel)
            Capture.oImage.Dispose()
            Capture.oImage = Nothing
            Capture = Nothing
        Next

        CompleteCapture = Nothing

        oGraphics.Dispose()
        oGraphics = Nothing

        CombinedImage = AddTextToImage(CombinedImage, Text)

        CaptureIE = SaveImageToFile(CombinedImage, destination)

        CombinedImage.Dispose()
        CombinedImage = Nothing

        IE = Nothing
    End Function

    Private Function CreateIEPartialCaptures(ByVal IE As Object, Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False) As CompleteCapture
        Dim IETabHwnd As Long = GetIEWindow(IE.hwnd)
        Dim documentBody As Object
        documentBody = IE.document.body

        Dim clientHeight As Integer, clientWidth As Integer, pageHeight As Integer, pageWidth As Integer
        Dim DeltaX As Integer, DeltaY As Integer, numHScroll As Integer, numVScroll As Integer
        Dim Height As Integer, Width As Integer

        'Get the size of page currently displayed
        clientHeight = documentBody.clientHeight
        clientWidth = documentBody.clientWidth

        Dim StartPosLeft As Integer, StartPosTop As Integer
        If CaptureIEFromCurrentPos Then
            StartPosLeft = documentBody.scrollLeft
            StartPosTop = documentBody.scrollTop
        Else
            StartPosLeft = 0
            StartPosTop = 0
        End If

        'Get the total size of the page 
        pageHeight = documentBody.scrollHeight - StartPosTop
        pageWidth = documentBody.scrollWidth - StartPosLeft


        'Check the flags on if we need scroll horizontally
        'and vertically
        If HorizontalScroll And VerticalScroll Then
            'Use total width of the page
            Height = pageHeight
            Width = pageWidth
        ElseIf HorizontalScroll Then
            'We need scroll horizontally only use the complete
            'page width
            Width = pageWidth

            'Use the current screen height
            Height = clientHeight
        ElseIf VerticalScroll Then
            'We need scroll horizontally only use the complete
            'page height
            Height = pageHeight

            'Use the current screen width
            Width = clientWidth
        Else
            'No scrolling, use the current screen size only
            Height = clientHeight
            Width = clientWidth
        End If

        'ShowWindow(IE.hwnd, SW_SHOW)
        'BringWindowToTop(IE.hwnd)
        'SetForegroundWindow(IE.hwnd)
        'Sleep(WaitBeforeCapture)

        'If pageHeight and clientHeight are same then no scroll bars
        'are present and we need a single scroll only
        'else if VerticalScroll Flag is off then we don't need to scroll
        If pageHeight = clientHeight Or Not VerticalScroll Then
            DeltaY = 0
            numVScroll = 1
        Else
            'Delta is kept to eliminate any error/missing image 
            'part when doing mutliple scrolls
            DeltaY = 5
            'No. of vertical scrolls required to cover the whole page
            numVScroll = pageHeight \ (clientHeight - DeltaY) + 1
        End If


        'If pageWidth and clientWidth are same then no scroll bars
        'are present and we need a single scroll only
        'else if HorizontalScroll Flag is off then we don't need to scroll
        If pageWidth = clientWidth Or Not HorizontalScroll Then
            DeltaX = 0
            numHScroll = 1
        Else
            'Delta is kept to eliminate any error/missing image 
            'part when doing mutliple scrolls
            DeltaX = 5
            'No. of horizontal scrolls required to cover the whole page
            numHScroll = pageWidth \ (clientWidth - DeltaX) + 1
        End If

        'Variable to store the no. of compeleted scrolls
        Dim curHScroll, curVScroll

        'Scroll the page top most position
        documentBody.ScrollTop = StartPosTop

        curVScroll = 0

        Dim TempPartialCapture() As PartialCapture
        ReDim TempPartialCapture(-1)
        Do
            curHScroll = 0

            'scroll the page to left most position
            documentBody.scrollLeft = StartPosLeft
            Do
                curHScroll = curHScroll + 1

                ReDim Preserve TempPartialCapture(UBound(TempPartialCapture) + 1)

                TempPartialCapture(UBound(TempPartialCapture)) = New PartialCapture

                With TempPartialCapture(UBound(TempPartialCapture))
                    'X coordinate of the current image on the final image
                    .x = documentBody.ScrollLeft - StartPosLeft

                    'Y coordinate of the current image on the final image
                    .y = documentBody.scrollTop - StartPosTop

                    'Width of the current capture
                    .width = clientWidth

                    'Height of the current capture
                    .height = clientHeight

                    'Capture the currently displayed part of the page
                    .oImage = GetImageFromHwnd(IETabHwnd)

                    'Scroll to the right of the page by (clientWidth - deltaX)
                    documentBody.scrollLeft = documentBody.scrollLeft + clientWidth - DeltaX

                End With
                'Loop until all required number of horizontal scrolls are completed
            Loop While curHScroll < numHScroll

            curVScroll = curVScroll + 1

            'Scroll down on the page by (clientHeight - deltaY)
            documentBody.scrollTop = documentBody.scrollTop + clientHeight - DeltaY

            'Loop until all required number of vertical scrolls are complete
        Loop While curVScroll < numVScroll

        CreateIEPartialCaptures.height = Height
        CreateIEPartialCaptures.width = Width
        CreateIEPartialCaptures.PartialCaptures = TempPartialCapture
    End Function

    Private Function GetIEFromSource(ByVal source As Object) As Object
        GetIEFromSource = Nothing

        Try
            Select Case TypeName(source)
                Case "IWebBrowser2"
                    'IE COM iterface return the same way
                    GetIEFromSource = source
                Case "Step"
                    'This is a QTP page object. Object property
                    'on parent will give the COM interface
                    GetIEFromSource = source.GetTOProperty("parent").object
                Case "CoBrowser"
                    'This is a QTP page object. Object property
                    'will give the document object
                    GetIEFromSource = source.object
                Case Else
                    GetIEFromSource = GetIEObjectFromHWND(source)
            End Select
        Catch ex As Exception

        End Try
    End Function

    ''' <summary>
    ''' Get "InternetExplorer.Application" type object from the hwnd of IE window
    ''' or a control inside IE window
    ''' </summary>
    ''' <param name="hwnd">Handle of object inside the IE window handle</param>
    ''' <returns>Internet explorer COM object.</returns>
    ''' <remarks>
    ''' Returns Nothing in case no IE is found
    ''' </remarks>
    Public Function GetIEObjectFromHWND(ByVal hwnd As Long) As Object Implements _ScreenCapture.GetIEObjectFromHWND
        GetIEObjectFromHWND = Nothing
        hwnd = GetRootWindow(hwnd)
        Dim allIE As Object, oIE As Object
        Try
            allIE = CreateObject("Shell.Application").windows
            For Each oIE In allIE
                If oIE.hwnd = hwnd Then
                    GetIEObjectFromHWND = oIE
                    Exit Function
                End If
            Next
        Catch ex As Exception
            allIE = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Convert image from one format to another format
    ''' </summary>
    ''' <param name="sourceFile">Source image file</param>
    ''' <param name="destFile">Destination image file. File extension will determine the file format</param>
    ''' <param name="OverWriteDestination">Default: True. True/False</param>
    ''' <remarks>
    ''' The file formats supported for destination files are (png, bmp, jpg, tiff, , wmf, emf, EXIF, gif)
    ''' </remarks>
    Public Sub ConvertImage(ByVal sourceFile As String, ByVal destFile As String, Optional ByVal OverWriteDestination As Boolean = True) Implements _ScreenCapture.ConvertImage
        If Not File.Exists(sourceFile) Then
            Err.Raise(vbObjectError + 1, "ConvertImage", "Could not find source file - " & sourceFile)
        End If

        If destFile.Trim() = "" <> 0 Then
            Err.Raise(vbObjectError + 2, "ConvertImage", "Invalid Destintaion file - " & destFile)
        End If

        If File.Exists(destFile) And Not OverWriteDestination Then
            Err.Raise(vbObjectError + 3, "ConvertImage", "Destination file already exists. Use OverwriteDestination as True")
        End If

        Dim srcImage As Drawing.Image

        Try
            srcImage = System.Drawing.Image.FromFile(sourceFile)
            Try
                SaveImageToFile(srcImage, destFile)
            Catch ex As Exception
                srcImage.Dispose()
                srcImage = Nothing
                Err.Raise(vbObjectError + 4, "ConvertImage", "Failed to save converted image to destination - " & destFile)
            End Try
        Catch ex As Exception
            Err.Raise(vbObjectError + 2, "ConvertImage", "Failed to Load the source image - " & sourceFile)
        End Try
    End Sub

    Public Function SaveImageToFile(ByVal oImage As Object, ByVal destFile As String) As Object Implements _ScreenCapture.SaveImageToFile
        Select Case TypeName(oImage)
            Case "Bitmap", "Image"
            Case Else
                Err.Raise(vbObjectError + 1, "SaveImageToFile", "Invalid parametere, object needs to be of type bitmap or image")
        End Select

        SaveImageToFile = False
        Dim fileFormat As String
        fileFormat = Path.GetExtension(destFile).ToLower

        Dim imgFormat As Imaging.ImageFormat

        Select Case fileFormat
            Case ".bmp"
                imgFormat = Imaging.ImageFormat.Bmp
            Case ".emf"
                imgFormat = Imaging.ImageFormat.Emf
            Case ".exif"
                imgFormat = Imaging.ImageFormat.Exif
            Case ".gif"
                imgFormat = Imaging.ImageFormat.Gif
            Case ".jpg", ".jpeg"
                imgFormat = Imaging.ImageFormat.Jpeg
            Case ".png"
                imgFormat = Imaging.ImageFormat.Png
            Case ".tiff"
                imgFormat = Imaging.ImageFormat.Tiff
            Case ".wmf"
                imgFormat = Imaging.ImageFormat.Wmf
            Case Else
                imgFormat = Imaging.ImageFormat.Png
        End Select

        If destFile.ToLower().Trim = "[clipboard]" Then
            Clipboard.Clear()
            Clipboard.SetImage(oImage)
            SaveImageToFile = True
        ElseIf destFile.ToLower().Trim = "[object]" Then
            SaveImageToFile = oImage.Clone
        ElseIf destFile.ToLower().Trim = "[user]" Then
            Dim saveDialog As SaveFileDialog
            saveDialog = New SaveFileDialog
            saveDialog.AddExtension = True
            saveDialog.OverwritePrompt = False
            saveDialog.DefaultExt = "png"
            saveDialog.Filter = "PNG File(*.png)|*.png|Bitmap(*.bmp)|*.bmp|JPG(*.jpg)|*.jpg|TIFF File(*.tiff)|*.tiff|WMF Files(*.wmf)|*.wmf|EMF Files(*.emf)|*.emf|EXIF Files(*.EXIF)|*.EXIF|GIF Files(*.gif)|*.gif|All files(*.*)|*.*"
            saveDialog.InitialDirectory = GetSetting("ScreenCapture", "SaveDialog", "InitDir", "Desktop")
            saveDialog.Title = "Save Srceenshot"
            Dim resSave As DialogResult
            resSave = saveDialog.ShowDialog()
            If Not (resSave = Windows.Forms.DialogResult.Cancel) Then
                SaveImageToFile(oImage, saveDialog.FileName)
                SaveSetting("ScreenCapture", "SaveDailog", "InitDir", Path.GetDirectoryName(saveDialog.FileName))
            End If

            saveDialog.Dispose()
            saveDialog = Nothing
        Else
            oImage.Save(destFile, imgFormat)
            SaveImageToFile = True
        End If

        imgFormat = Nothing
    End Function

    Private Function GetImageFromHwnd(ByVal hwnd As Long, Optional ByVal client As Boolean = False, Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False) As Bitmap
        If IsWindow(hwnd) = 0 Then
            Err.Raise(vbObjectError + 1, "GetImageFromHwnd", "Window handle is not valid")
        End If

        If IsWindowVisible(hwnd) = 0 Then
            Err.Raise(vbObjectError + 1, "GetImageFromHwnd", "Window handle not visible")
        End If

        ShowCursor(False)

        Dim windowRECT As RECT
        Dim windowDC As Long

        If ActivateWindowOnCapture Then
            ActivateWindow(GetRootWindow(hwnd))
        End If

        Sleep(WaitBeforeCapture)

        If client Then
            GetClientRect(hwnd, windowRECT)
            windowDC = GetDC(hwnd)
        Else
            GetWindowRect(hwnd, windowRECT)
            windowDC = GetWindowDC(hwnd)
        End If

        Dim oBitmap As Bitmap = New Bitmap(windowRECT.Right - windowRECT.Left, windowRECT.Bottom - windowRECT.Top)
        'Bitmap bitmap = new Bitmap(rc.Width, rc.Height);

        Dim oGraphics As Graphics = Graphics.FromImage(oBitmap)
        'Graphics gfxBitmap = Graphics.FromImage(bitmap);

        '// get a device context for the bitmap
        'IntPtr hdcBitmap = gfxBitmap.GetHdc();
        Dim hdcBitmap As IntPtr = oGraphics.GetHdc()

        '// get a device context for the window
        'IntPtr hdcWindow = Win32.GetWindowDC(hWnd); 
        Dim hdcWindow As IntPtr = windowDC

        If hwnd = GetDesktopWindow() Then
            '// bitblt the window to the bitmap
            BitBlt(hdcBitmap, 0, 0, windowRECT.Right - windowRECT.Left, _
                                   windowRECT.Bottom - windowRECT.Top, _
            hdcWindow, 0, 0, (SRCCOPY))
        Else
            PrintWindow(hwnd, hdcBitmap, 0)
        End If

        '// release the bitmap's device context
        oGraphics.ReleaseHdc(hdcBitmap)

        oGraphics.Flush()

        'DeleteDC(hdcBitmap)
        'DeleteDC(hdcWindow)
        ReleaseDC(hwnd, hdcWindow)

        oGraphics.Dispose()

        GetImageFromHwnd = oBitmap

        ShowCursor(True)
        'oBitmap.Save("c:\tests.png", Imaging.ImageFormat.Png)
    End Function

    ''' <summary>
    ''' Get handle of the top level window using a handle object inside the window
    ''' </summary>
    ''' <param name="hwnd">Handle of the window</param>
    ''' <returns>Handle of the top level window</returns>
    ''' <remarks></remarks>
    Public Function GetRootWindow(ByVal hwnd As Long) As Long Implements _ScreenCapture.GetRootWindow
        'Do
        'GetRootWindow = hwnd
        'Debug.Print(hwnd)
        'hwnd = GetParent(hwnd)
        'Loop While hwnd <> 0
        GetRootWindow = GetAncestor(hwnd, GA_ROOT)
    End Function

    Private Sub ActivateWindow(ByVal hwnd As Long)
        Dim topWindow As Long
        topWindow = GetRootWindow(hwnd)

        Dim winPlacement As WINDOWPLACEMENT
        GetWindowPlacement(topWindow, winPlacement)

        Select Case winPlacement.showCmd
            Case SW_SHOWMINIMIZED
                'winPlacement.showCmd = SW_SHOW
                ShowWindow(topWindow, SW_RESTORE)
            Case SW_NORMAL, SW_MAXIMIZE
                ShowWindow(topWindow, SW_SHOW)
        End Select
        BringWindowToTop(topWindow)
        SetForegroundWindow(topWindow)
    End Sub

    Private Function GetIEWindow(ByVal hwnd As Long) As Long
        Dim IEhWnd As Long = GetRootWindow(hwnd)
        Dim hwnd_Server As Long
        hwnd_Server = FindWindowEx(IEhWnd, 0, "TabWindowClass", "")

        If hwnd_Server = 0 Then
            hwnd_Server = FindWindowEx(IEhWnd, 0, "Shell DocObject View", "")
        Else
            Dim hwnd_Server2 As Long
            hwnd_Server2 = 0
            Do
                hwnd_Server2 = FindWindowEx(hwnd_Server, hwnd_Server2, "Shell DocObject View", "")
                If IsWindowVisible(hwnd_Server2) <> 0 Then
                    hwnd_Server = hwnd_Server2
                    Exit Do
                End If
            Loop While hwnd_Server2 <> 0
        End If

        If hwnd_Server <> 0 Then
            hwnd_Server = FindWindowEx(hwnd_Server, 0, "Internet Explorer_Server", "")
        End If

        GetIEWindow = hwnd_Server
    End Function

    Public Function CaptureWindowRect(ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal destination As String, Optional ByVal Text As String = "") Implements _ScreenCapture.CaptureWindowRect
        Dim oBitmap As Bitmap
        oBitmap = CaptureWindow(hwnd, "[object]")
        Dim newBitmap As Bitmap = New Bitmap(CInt(width), CInt(height))
        Dim oGraphics As Graphics = Graphics.FromImage(newBitmap)
        oGraphics.DrawImage(oBitmap, New Rectangle(0, 0, width, height), New Rectangle(x, y, width, height), GraphicsUnit.Pixel)
        oGraphics.Flush()
        oGraphics.Dispose()
        oBitmap.Dispose()
        newBitmap = AddTextToImage(newBitmap, Text)
        CaptureWindowRect = SaveImageToFile(newBitmap, destination)
        newBitmap.Dispose()
        newBitmap = Nothing
    End Function

    'Public Function CompareImages(ByVal Image1 As Object, ByVal Image2 As Object, Optional ByVal ScaleImages As Boolean = False) As Integer
    '    Dim oImage1 As Bitmap
    '    Dim oImage2 As Bitmap
    '    If Image1.GetType().Name = "String" Then
    '        oImage1 = Image.FromFile(Image1)
    '    Else
    '        oImage1 = Image1.clone
    '    End If

    '    Dim imgConverter As ImageConverter = New ImageConverter
    '    Dim img1Bytes As Byte()
    '    Dim img2Bytes As Byte()

    '    If Image2.GetType().Name = "String" Then
    '        oImage2 = Image.FromFile(Image2)
    '    Else
    '        oImage2 = Image2.clone
    '    End If

    '    If (oImage1.Size <> oImage2.Size) And Not ScaleImages Then
    '        CompareImages = 1
    '    Else
    '        If (oImage1.Size <> oImage2.Size) Then
    '            If (oImage1.Width > oImage2.Width And oImage1.Height > oImage2.Height) Or (oImage1.Width + oImage1.Height > oImage2.Width + oImage2.Height) Then
    '                oImage1 = New Bitmap(oImage1, oImage2.Width, oImage2.Height)
    '            Else
    '                oImage2 = New Bitmap(oImage2, oImage1.Width, oImage1.Height)
    '            End If
    '        End If

    '        Dim oXORImage As Graphics = Graphics.FromImage(oImage1)
    '        Dim oXORImage2 As Graphics = Graphics.FromImage(oImage2)
    '        Dim hdc1 As Long, hdc2 As Long
    '        hdc1 = oXORImage.GetHdc
    '        hdc2 = oXORImage2.GetHdc

    '        'oBitmap = CreateCompatibleDC(hdc1)
    '        'oBmp = CreateCompatibleBitmap(hdc1, oImage1.Width, oImage1.Height)

    '        StretchBlt(hdc1, 0, 0, oImage1.Width, oImage1.Height, hdc2, 0, 0, oImage1.Width, oImage1.Height, SRCINVERT)
    '        DeleteDC(hdc1)
    '        DeleteDC(hdc2)


    '        Dim i As Integer
    '        CompareImages = 0
    '        ReDim img1Bytes(0)
    '        ReDim img2Bytes(0)

    '        img1Bytes = imgConverter.ConvertTo(oImage1, img1Bytes.GetType())
    '        img2Bytes = imgConverter.ConvertTo(oImage2, img1Bytes.GetType())

    '        Dim oHash As System.Security.Cryptography.HashAlgorithm
    '        oHash = Security.Cryptography.HashAlgorithm.Create("MD5")
    '        img1Bytes = oHash.ComputeHash(img1Bytes)
    '        img2Bytes = oHash.ComputeHash(img2Bytes)

    '        For i = LBound(img1Bytes) To UBound(img1Bytes)
    '            If img1Bytes(i) <> img2Bytes(i) Then CompareImages = 2 : Exit For
    '        Next

    '        img2Bytes = Nothing
    '        img1Bytes = Nothing
    '    End If

    '    oImage1.Dispose()
    '    oImage2.Dispose()

    '    oImage1 = Nothing
    '    oImage2 = Nothing
    'End Function

    Public Function CaptureDesktopRect(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal destination As String, Optional ByVal text As String = "") Implements _ScreenCapture.CaptureDesktopRect
        CaptureDesktopRect = CaptureWindowRect(GetDesktopWindow, x, y, width, height, destination, text)
    End Function

    'Private Function CaptureURL(ByVal URL As String, ByVal Destination As String, Optional ByVal Text As String = "")
    '    Dim oWebBrw As WebBrowser
    '    CaptureURL = Nothing
    '    Try
    '        oWebBrw = New WebBrowser
    '        oWebBrw.Navigate(URL)

    '        If oWebBrw.Document Is Nothing Then
    '            Err.Raise(vbObjectError + 1, "CaptureURL", "Error while browsing URL - " & URL)
    '        End If

    '        Dim scrollWidth As Integer
    '        Dim scrollHeight As Integer
    '        scrollHeight = oWebBrw.Document.Body.ScrollRectangle.Height
    '        scrollWidth = oWebBrw.Document.Body.ScrollRectangle.Width
    '        oWebBrw.Size = New Size(scrollWidth, scrollHeight)

    '        Dim bm As New Bitmap(scrollWidth, scrollHeight)
    '        oWebBrw.DrawToBitmap(bm, New Rectangle(0, 0, bm.Width, bm.Height))
    '        CaptureURL = SaveImageToFile(bm, Destination)
    '        bm.Dispose()
    '        bm = Nothing
    '        oWebBrw.Dispose()
    '        oWebBrw = Nothing
    '    Catch ex As Exception
    '    Finally
    '        'oWebBrw.Dispose()
    '        oWebBrw = Nothing
    '    End Try
    'End Function

    Public Sub GetCursorPosition(ByRef X As Object, ByRef Y As Object) Implements _ScreenCapture.GetCursorPosition
        Dim pos As POINTAPI

        GetCursorPos(pos)

        X = pos.x

        Y = pos.y
    End Sub

    Public Function CombineImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal Position As Integer = MERGE_Image1TopImage2, Optional ByVal Padding As Integer = 0) Implements _ScreenCapture.CombineImages
        Dim bmp1 As Bitmap
        Dim bmp2 As Bitmap
        Dim bmpDest As Bitmap

        If SrcImage1.GetType().Name = "String" Then
            bmp1 = Image.FromFile(SrcImage1)
        Else
            bmp1 = SrcImage1.clone
        End If

        If SrcImage2.GetType().Name = "String" Then
            bmp2 = Image.FromFile(SrcImage2)
        Else
            bmp2 = SrcImage2.Clone
        End If

        Select Case Position
            Case MERGE_Image1RightImage2, MERGE_Image1LeftImage2
                bmpDest = New Bitmap(bmp1.Width + bmp2.Width + Padding, Math.Max(bmp1.Height, bmp2.Height))
            Case Else 'MERGE_Image1BottomImage2, MERGE_Image1TopImage2
                bmpDest = New Bitmap(Math.Max(bmp1.Width, bmp2.Width), bmp1.Height + bmp2.Height + Padding)
                'Case Else
                Position = MERGE_Image1BottomImage2
        End Select

        Dim oGraphics As Graphics = Graphics.FromImage(bmpDest)

        Select Case Position
            Case MERGE_Image1LeftImage2
                oGraphics.DrawImage(bmp1, New Point(0, 0))
                oGraphics.DrawImage(bmp2, New Point(bmp1.Width + Padding, 0))
            Case MERGE_Image1RightImage2
                oGraphics.DrawImage(bmp1, New Point(bmp2.Width + Padding, 0))
                oGraphics.DrawImage(bmp2, New Point(0, 0))
            Case MERGE_Image1TopImage2
                oGraphics.DrawImage(bmp1, New Point(0, 0))
                oGraphics.DrawImage(bmp2, New Point(0, bmp1.Height + Padding))
            Case MERGE_Image1BottomImage2
                oGraphics.DrawImage(bmp1, New Point(0, bmp2.Height + Padding))
                oGraphics.DrawImage(bmp2, New Point(0, 0))
        End Select

        CombineImages = SaveImageToFile(bmpDest, Destination)

        bmpDest.Dispose()
        bmp1.Dispose()
        bmp2.Dispose()
        oGraphics.Dispose()

        bmpDest = Nothing
        bmp1 = Nothing
        bmp2 = Nothing
        oGraphics = Nothing
    End Function

    Public Function GetNewBitmap(ByVal Width As Integer, ByVal Height As Integer) As Bitmap Implements _ScreenCapture.GetNewBitmap
        Dim RetBmp As Bitmap
        RetBmp = New Bitmap(Width, Height)
        GetNewBitmap = RetBmp.Clone
        RetBmp.Dispose()
        RetBmp = Nothing
    End Function


    Public Function FindWindowLike(Optional ByVal WindowTitle As String = "*", Optional ByVal WindowClass As String = "*") Implements _ScreenCapture.FindWindowLike
        Dim WindowsEnum As New WindowsEnumerator
        Dim windowsInfo As List(Of WindowInfo)
        windowsInfo = WindowsEnum.GetTopLevelWindows()
        Dim matchWindows As New Collection
        Dim oWindow As WindowInfo

        For Each oWindow In windowsInfo
            If oWindow.Title Like WindowTitle AndAlso oWindow.ClassName Like WindowClass Then
                matchWindows.Add(oWindow)
            End If
        Next

        FindWindowLike = matchWindows

        WindowsEnum = Nothing
    End Function
End Class
