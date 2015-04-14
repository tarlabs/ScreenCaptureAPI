Option Compare Text

Imports System.IO
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

'<Guid("168de179-a9c1-40f9-a8bf-398bbf91695e"), _
'InterfaceType(ComInterfaceType.InterfaceIsIDispatch), ComVisible(True)> _
'Public Interface _ScreenCapture2
'    _ScreenCapture
'End Interface

<ComImport(), Guid("332C4425-26CB-11D0-B483-00C04FD90119")> _
Public Interface HTMLDocumentClass
End Interface


<Guid("E7852903-FA4C-4a84-8832-21DC7A9A150D"), _
InterfaceType(ComInterfaceType.InterfaceIsIDispatch), ComVisible(True)> _
Public Interface _ScreenCapture
    <DispId(50)> Function CaptureIEObject(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False, Optional ByVal CaptureSourceObjectOnly As Boolean = False)
    <DispId(51)> Function GetIEObjectFromProperty(ByVal PropName As String, ByVal PropValue As Object) As Object
    <DispId(52)> Function EnumIEFramesDocument(ByVal wb As Object) As Object
    <DispId(53)> Function CaptureIEFrames(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False)

    <DispId(30)> Function CompareImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal ScaleImages As Boolean = True)
    <DispId(1)> Function SaveImageToFile(ByVal oImage As Object, ByVal destFile As String) As Object
    <DispId(2)> Function AddTextToImage(ByVal oImage As Object, ByVal Text As String)
    <DispId(3)> Function CaptureActiveWindow(ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "")
    <DispId(4)> Function CaptureDesktop(ByVal destination As String, Optional ByVal Text As String = "")
    <DispId(5)> Function CaptureIE(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False)
    <DispId(6)> Function CaptureDesktopRect(ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32, ByVal destination As String, Optional ByVal text As String = "")
    <DispId(7)> Function CaptureWindow(ByVal hwnd As Int32, ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = True, Optional ByVal HorizontalScroll As Boolean = True)
    <DispId(8)> Function CaptureWindowRect(ByVal hwnd As Int32, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32, ByVal destination As String, Optional ByVal Text As String = "")
    <DispId(9)> Sub ConvertImage(ByVal sourceFile As String, ByVal destFile As String, Optional ByVal OverWriteDestination As Boolean = True)
    <DispId(10)> Function GetIEObjectFromHWND(ByVal hwnd As Int32) As Object
    <DispId(11)> Function GetRootWindow(ByVal hwnd As Int32) As Int32
    <DispId(12)> Function CombineImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal Position As Integer = 3, Optional ByVal Padding As Integer = 0)
    <DispId(13)> Function GetNewBitmap(ByVal Width As Integer, ByVal Height As Integer) As Bitmap
    <DispId(14)> Function FindWindowLike(Optional ByVal WindowTitle As String = "*", Optional ByVal WindowClass As String = "*")
    <DispId(15)> Sub GetCursorPosition(ByRef X As Object, ByRef Y As Object)
    <DispId(16)> Property TextPos() As Int32
    <DispId(17)> Property TextFontName() As String
    <DispId(18)> Property TextFontSize() As Integer
    <DispId(19)> Property ActivateWindowOnCapture() As Boolean
    <DispId(20)> Property WaitBeforeCapture() As Int32
    <DispId(21)> Property CaptureIEFromCurrentPos() As Boolean
    <DispId(22)> Property TextHeight() As Int32
    <DispId(23)> Property CombineDeltaX() As Int32
    <DispId(24)> Property CombineDeltaY() As Int32
End Interface


<Guid("EA4215C4-3DF4-4b3d-BE76-CB7183B48605"), _
ClassInterface(ClassInterfaceType.None), _
ProgId("KnowledgeInbox.ScreenCapture")> _
Public Class ScreenCapture
    'Implements _ScreenCapture2
    Implements _ScreenCapture

    Public Const MERGE_Image1RightImage2 = 0
    Public Const MERGE_Image1LeftImage2 = 1
    Public Const MERGE_Image1BottomImage2 = 2
    Public Const MERGE_Image1TopImage2 = 3

    Private mActivateWindowOnCapture As Boolean = True
    Private mWaitBeforeCapture As Int32 = 300
    Private mCaptureIEFromCurrentPos As Boolean = False
    Private mTextHeight As Int32 = 20
    Private mCombineDeltaX As Int32 = 2
    Private mCombineDeltaY As Int32 = 2
    Private mTextPos As Integer = 2 '0-Top, 1-bottom, 2-Both
    Private mTextFontName As String = "Verdana"
    Private mTextFontSize As Integer = 10



    ''' <summary>
    ''' The Y delta to be used in case of combination of image. Adjust this if you get lines
    ''' in combine images
    ''' </summary>
    ''' <value>True/False</value>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Private Property CombineDeltaY() As Int32 Implements _ScreenCapture.CombineDeltaY
        Get
            CombineDeltaY = mCombineDeltaY
        End Get

        Set(ByVal NewValue As Int32)
            mCombineDeltaY = NewValue
        End Set
    End Property


    ''' <summary>
    ''' The X delta to be used in case of combination of image. Adjust this if you get lines
    ''' in combine images
    ''' </summary>
    ''' <value>True/False</value>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Public Property CombineDeltaX() As Int32 Implements _ScreenCapture.CombineDeltaX
        Get
            CombineDeltaX = mCombineDeltaX
        End Get

        Set(ByVal NewValue As Int32)
            mCombineDeltaX = NewValue
        End Set
    End Property


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
    Public Property WaitBeforeCapture() As Int32 Implements _ScreenCapture.WaitBeforeCapture
        Get
            WaitBeforeCapture = mWaitBeforeCapture
        End Get
        Set(ByVal value As Int32)
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
    Public Property TextHeight() As Int32 Implements _ScreenCapture.TextHeight
        Get
            TextHeight = mTextHeight
        End Get
        Set(ByVal value As Int32)
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
    Public Property TextPos() As Int32 Implements _ScreenCapture.TextPos
        Get
            TextPos = mTextPos
        End Get
        Set(ByVal value As Int32)
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
        CompareImages = _CompareImages(SrcImage1, SrcImage2, Destination, ScaleImages)
    End Function

    Private Function _CompareImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal ScaleImages As Boolean = True)
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
            _CompareImages = -1
        Else
            If (bmp1.Size <> bmp2.Size) Then
                If (bmp1.Width > bmp2.Width And bmp1.Height > bmp2.Height) Or (bmp1.Width + bmp1.Height > bmp2.Width + bmp2.Height) Then
                    bmp1 = New Bitmap(bmp1, bmp2.Width, bmp2.Height)
                Else
                    bmp2 = New Bitmap(bmp2, bmp1.Width, bmp1.Height)
                End If
            End If
        End If

        Dim CompWidth As Integer, CompHeight As Integer
        Dim PixelWidth As Integer, PixelPadding As Integer, RowPixelCount As Integer
        Dim MisMatchCount As Integer = 0
        Dim MisMatchRedCount As Integer = 0
        Dim MisMatchGreenCount As Integer = 0
        Dim MisMatchBlueCount As Integer = 0

        CompWidth = bmp1.Width
        CompHeight = bmp2.Height
        Dim bmp3 As Bitmap = New Bitmap(CompWidth, CompHeight)

        Dim bmpData1 As Imaging.BitmapData = bmp1.LockBits(New Rectangle(0, 0, CompWidth, CompHeight), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Dim bmpData2 As Imaging.BitmapData = bmp2.LockBits(New Rectangle(0, 0, CompWidth, CompHeight), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Dim bmpData3 As Imaging.BitmapData = bmp3.LockBits(New Rectangle(0, 0, CompWidth, CompHeight), Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Dim bytes As Int32 = bmpData1.Stride * CompHeight

        PixelWidth = bmpData1.Stride \ bmpData1.Width
        PixelPadding = bmpData1.Stride Mod bmpData1.Width
        RowPixelCount = bmpData1.Stride - PixelPadding

        Dim buffer1(bytes - 1), buffer2(bytes - 1), buffer3(bytes - 1) As Byte

        System.Runtime.InteropServices.Marshal.Copy(bmpData1.Scan0, buffer1, 0, bytes)
        System.Runtime.InteropServices.Marshal.Copy(bmpData2.Scan0, buffer2, 0, bytes)

        'For i As Int32 = 0 To buffer1.GetUpperBound(0)
        'buffer3(i) = buffer1(i) Xor buffer2(i)
        'Next

        Dim i As Int32
        i = 0
        For iRow As Integer = 0 To CompHeight - 1
            For iCell As Integer = 0 To RowPixelCount - 1 Step 3
                buffer3(i) = buffer2(i) Xor buffer1(i)
                buffer3(i + 1) = buffer2(i + 1) Xor buffer1(i + 1)
                buffer3(i + 2) = buffer2(i + 2) Xor buffer1(i + 2)

                If buffer3(i) <> 0 Then MisMatchRedCount = MisMatchRedCount + 1
                If buffer3(i + 1) <> 0 Then MisMatchGreenCount = MisMatchGreenCount + 1
                If buffer3(i + 2) <> 0 Then MisMatchBlueCount = MisMatchBlueCount + 1

                If buffer3(i) <> 0 Or buffer3(i + 1) <> 0 Or buffer3(i + 2) <> 0 Then
                    MisMatchCount = MisMatchCount + 1
                End If

                i = i + PixelWidth
            Next
            i = i + PixelPadding
            'If i Mod RowPixelCount = 0 And i <> 0 Then
            '    i = i + PixelPadding
            'Else
            '    If buffer3(i) <> 0 Or buffer3(i + 1) <> 0 Or buffer3(i + 2) <> 0 Then
            '        MisMatchCount = MisMatchCount + 1
            '    End If
            '    i = i + 3
            'End If
        Next


        'Dim MisMatchCount As Int32
        'MisMatchCount = 0

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
                'Dim i As Int32
                'Dim iCount As Int32

                'iCount = UBound(buffer3)
                'i = 0
                'For iRow As Integer = 0 To CompHeight - 1
                '    For iCell As Integer = 0 To RowPixelCount - 1 Step 3
                '        If buffer3(i) <> 0 Or buffer3(i + 1) <> 0 Or buffer3(i + 2) <> 0 Then
                '            MisMatchCount = MisMatchCount + 1
                '        End If
                '        i = i + 3
                '    Next
                '    i = i + PixelPadding
                '    'If i Mod RowPixelCount = 0 And i <> 0 Then
                '    '    i = i + PixelPadding
                '    'Else
                '    '    If buffer3(i) <> 0 Or buffer3(i + 1) <> 0 Or buffer3(i + 2) <> 0 Then
                '    '        MisMatchCount = MisMatchCount + 1
                '    '    End If
                '    '    i = i + 3
                '    'End If
                'Next
                _CompareImages = MisMatchCount
                If Destination.ToLower = "[pixeldiffperc]" Then
                    _CompareImages = (MisMatchCount / (bmp3.Width * bmp3.Height)) * 100
                End If
            Case Else
                _CompareImages = SaveImageToFile(bmp3, Destination)
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
        AddTextToImage = _AddTextToImage(oImage, Text)
    End Function

    Private Function _AddTextToImage(ByVal oImage As Object, ByVal Text As String)
        Select Case TypeName(oImage)
            Case "Bitmap", "Image"
            Case Else
                Err.Raise(vbObjectError + 1, "AddTextToImage", "Invalid parameter, object needs to be of type bitmap or image")
        End Select

        On Error GoTo errHandler

      If Text.Trim = "" Then
          _AddTextToImage = oImage
          Exit Function
     End If

     Dim AddTextHeight As Integer
     If TextPos = 1 Or TextPos = 2 Then
         AddTextHeight = TextHeight
     Else
         AddTextHeight = 2 * TextHeight
     End If

     Dim oNewImage As Image = New Bitmap(CInt(oImage.Width), CInt(oImage.Height + AddTextHeight))
     Dim oGraphics As Graphics = Graphics.FromImage(oNewImage)

     Dim oFont As Font = New Font(TextFontName, TextFontSize)
     If TextPos = 1 Then
         oGraphics.FillRectangle(Brushes.White, 0, 0, oImage.Width, TextHeight)
         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, 2)
         oGraphics.DrawImage(oImage, 0, TextHeight) '+ oImage.Height)
     ElseIf TextPos = 2 Then
         oGraphics.FillRectangle(Brushes.White, 0, oImage.Height, oImage.Width, oImage.Height + TextHeight)

         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, oImage.Height + 2)
         oGraphics.DrawImage(oImage, 0, 0)
     Else
         oGraphics.FillRectangle(Brushes.White, 0, 0, oImage.Width, TextHeight)
         oGraphics.FillRectangle(Brushes.White, 0, oImage.Height + TextHeight, oImage.Width, oImage.Height + TextHeight)

         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, 2)
         oGraphics.DrawString(Text, oFont, Brushes.Black, 2, 2 + oImage.Height + TextHeight)
         oGraphics.DrawImage(oImage, 0, TextHeight)
     End If

     oGraphics.Flush()
     oGraphics.Dispose()
     oGraphics = Nothing

     _AddTextToImage = oNewImage
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
        CaptureActiveWindow = _CaptureActiveWindow(destination, Client, Text)
    End Function

    Private Function _CaptureActiveWindow(ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "")
        Dim oBitmap As Bitmap
        oBitmap = _GetImageFromHwnd(GetForegroundWindow(), Client)
        oBitmap = _AddTextToImage(oBitmap, Text)
        _CaptureActiveWindow = _SaveImageToFile(oBitmap, destination)
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
        CaptureDesktop = _CaptureDesktop(destination, Text)
    End Function

    Private Function _CaptureDesktop(ByVal destination As String, Optional ByVal Text As String = "")
        Dim oBitmap As Bitmap
        oBitmap = _GetImageFromHwnd(GetDesktopWindow(), False)
        oBitmap = _AddTextToImage(oBitmap, Text)
        _CaptureDesktop = _SaveImageToFile(oBitmap, destination)
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

    Public Function CaptureWindow(ByVal hwnd As Int32, ByVal destination As String, Optional ByVal Client As Boolean = False, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = True, Optional ByVal HorizontalScroll As Boolean = True) Implements _ScreenCapture.CaptureWindow
        Dim oBitmap As Bitmap
        oBitmap = _GetImageFromHwnd(hwnd, Client, VerticalScroll, HorizontalScroll)
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
        CaptureIE = _CaptureIEObject(source, destination, Text, VerticalScroll, HorizontalScroll)
    End Function

    Public Function CaptureIEObject(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False, Optional ByVal CaptureSourceObjectOnly As Boolean = False) Implements _ScreenCapture.CaptureIEObject
        CaptureIEObject = _CaptureIEObject(source, destination, Text, VerticalScroll, HorizontalScroll, CaptureSourceObjectOnly)
    End Function

    Private Function _CaptureIEObject(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False, Optional ByVal CaptureSourceObjectOnly As Boolean = False)
        Dim IE As SHDocVW.IWebBrowser2
        IE = _GetIEFromSource(source)

        Dim scrollingObject As Object
        _CaptureIEObject = Nothing

        If IE Is Nothing Then
            If IsNumeric(source) Then
                _CaptureIEObject = SaveImageToFile(_GetImageFromHwnd(source), destination)
                Exit Function
            Else
                _CaptureIEObject = False
                Exit Function
            End If
        ElseIf TypeName(IE.Document).ToLower() <> "htmldocument" And TypeName(IE.Document).ToLower() <> "htmldocumentclass" Then
            If IsNumeric(source) Then
                _CaptureIEObject = SaveImageToFile(_GetImageFromHwnd(source), destination)
                Exit Function
            Else
                _CaptureIEObject = SaveImageToFile(_GetImageFromHwnd(IE.HWND), destination)
                '_CaptureIEObject = False
                Exit Function
            End If
        End If

        If CaptureSourceObjectOnly Then
            scrollingObject = _GetScrollObjectFromSource(source)
        Else
            scrollingObject = Nothing
        End If

        Dim CompleteCapture As CompleteCapture
        'CaptureIEFromCurrentPos = True
        Dim oldActivateMode As Boolean, oldWait As Int32
        oldWait = WaitBeforeCapture

        oldActivateMode = ActivateWindowOnCapture

        If oldActivateMode Then
            _ActivateWindow(IE.HWND)
            Sleep(oldWait)
        End If

        ActivateWindowOnCapture = False
        WaitBeforeCapture = 0

        CompleteCapture = _CreateIEPartialCaptures(IE, VerticalScroll, HorizontalScroll, scrollingObject)

        If CompleteCapture.PartialCaptures Is Nothing Then Exit Function

        ActivateWindowOnCapture = oldActivateMode
        WaitBeforeCapture = oldWait

        'Combine all images into one
        Dim CombinedImage As Bitmap ', oGraphics As Graphics
        CombinedImage = _CombinePartialImages(CompleteCapture)
        'CombinedImage = New Bitmap(CompleteCapture.width, CompleteCapture.height)

        'oGraphics = Graphics.FromImage(CombinedImage)

        'Dim Capture As PartialCapture

        'For Each Capture In CompleteCapture.PartialCaptures
        '    oGraphics.DrawImage(Capture.oImage, New Rectangle(Capture.x, Capture.y, Capture.width, Capture.height), New Rectangle(Me.CombineDeltaX, Me.CombineDeltaY, Capture.width, Capture.height), GraphicsUnit.Pixel)
        '    Capture.oImage.Dispose()
        '    Capture.oImage = Nothing
        '    Capture = Nothing
        'Next

        CompleteCapture = Nothing

        CombinedImage = _AddTextToImage(CombinedImage, Text)

        _CaptureIEObject = _SaveImageToFile(CombinedImage, destination)

        CombinedImage.Dispose()
        CombinedImage = Nothing

        IE = Nothing
    End Function

    Function _CombinePartialImages(ByRef CompleteCapture As CompleteCapture)

        'If there is only one image then return the same and no need to process
        If CompleteCapture.PartialCaptures.Length = 1 Then
            _CombinePartialImages = CompleteCapture.PartialCaptures(0).oImage
            Exit Function
        End If

        Dim oGraphics As Graphics
        Dim CombinedImage As Bitmap
        CombinedImage = New Bitmap(CompleteCapture.width, CompleteCapture.height)

        oGraphics = Graphics.FromImage(CombinedImage)

        Dim Capture As PartialCapture

        For i As Integer = CompleteCapture.PartialCaptures.Length - 1 to 0 Step -1
            Capture = CompleteCapture.PartialCaptures(i)
            'Capture.oImage.Save("C:\\Temp\\Part" & i & ".png")
            oGraphics.DrawImage(Capture.oImage, New Rectangle(Capture.x, Capture.y, Capture.width, Capture.height), New Rectangle(Me.CombineDeltaX, Me.CombineDeltaY, Capture.width, Capture.height), GraphicsUnit.Pixel)
            Capture.oImage.Dispose()
            Capture.oImage = Nothing
            Capture = Nothing
        Next

        oGraphics.Dispose()
        oGraphics = Nothing
        _CombinePartialImages = CombinedImage
    End Function

    Function _CallByNameEx(ByVal ObjectRef As Object, ByVal ProcName As String, ByVal useCallType As Microsoft.VisualBasic.CallType, ByVal ParamArray Args() As Object) As Object
        Dim ProcNames() As String

        _CallByNameEx = Nothing

        ProcNames = ProcName.Split(".")

        Dim i As Integer

        For i = 0 To ProcNames.Length - 1
            ObjectRef = CallByName(ObjectRef, ProcNames(i), useCallType, Args)
        Next

        _CallByNameEx = ObjectRef
    End Function

    Public Function GetIEObjectFromProperty(ByVal PropName As String, ByVal PropValue As Object) As Object Implements _ScreenCapture.GetIEObjectFromProperty
        GetIEObjectFromProperty = _GetIEObjectFromProperty(PropName, PropValue)
    End Function

    Private Function _GetIEObjectFromProperty(ByVal PropName As String, ByVal PropValue As Object) As Object
        _GetIEObjectFromProperty = Nothing
        PropName = LCase(PropName)

        If PropName = "hwnd" Then
            PropValue = GetRootWindow(PropValue)
        End If

        Dim allIE As Object, oIE As SHDocVW.IWebBrowser2

        Try
            allIE = CreateObject("Shell.Application").windows
            For Each oIE In allIE
                Dim actualValue As Object
                Try
                    actualValue = _CallByNameEx(oIE, PropName, CallType.Get)

                    If PropValue.GetHashCode = actualValue.GetHashCode Then
                        _GetIEObjectFromProperty = oIE
                        Exit Function
                    ElseIf PropName = "document" Then
                        If PropValue.all.GetHashCode = actualValue.all.GetHashCode Then
                            _GetIEObjectFromProperty = oIE
                            Exit Function
                        End If
                    End If

                Catch ex As Exception
                End Try

            Next
        Catch ex As Exception
            allIE = Nothing
        End Try
    End Function

    Public Function EnumIEFramesDocument(ByVal wb As Object) As Object Implements _ScreenCapture.EnumIEFramesDocument
        EnumIEFramesDocument = _EnumIEFramesDocument(wb)
    End Function

    Public Function CaptureIEFrames(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False) Implements _ScreenCapture.CaptureIEFrames
        CaptureIEFrames = _CaptureIEFrames(source, destination, Text, VerticalScroll, HorizontalScroll)
    End Function

    Private Function _CaptureIEFrames(ByVal source As Object, ByVal destination As String, Optional ByVal Text As String = "", Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False)
        _CaptureIEFrames = False
        Dim IE As SHDocVW.IWebBrowser2
        IE = _GetIEFromSource(source)

        If IE Is Nothing Then Exit Function

        Dim allFrames As Collection
        allFrames = _EnumIEFramesDocument(IE.Document)

        Dim oFrameDoc As HTMLDocumentClass
        Dim oCaptures As Collection

        If allFrames.Count = 0 Then
            _CaptureIEFrames = _CaptureIEObject(source, destination, Text, VerticalScroll, HorizontalScroll)
            'ElseIf allFrames.Count = 1 Then
            '    _CaptureIEFrames = _CaptureIEObject(allFrames(1), destination, Text, VerticalScroll, HorizontalScroll)
        Else
            oCaptures = New Collection
            Dim i As Integer = 0
            oCaptures.Clear()
            oCaptures.Add(_CaptureIEObject(IE, "[object]", "Main document", VerticalScroll, HorizontalScroll, True))

            For Each oFrameDoc In allFrames
                Dim oCapturedImage As Object
                oCapturedImage = _CaptureIEObject(oFrameDoc, "[object]", "Frame - " & i, VerticalScroll, HorizontalScroll, True)
                If Not TypeOf (oCapturedImage) Is Boolean Then
                    i = i + 1
                    oCaptures.Add(oCapturedImage)
                End If
            Next

            Dim oFinalImage As Object

            If oCaptures.Count > 1 Then
                oFinalImage = _CombineImages(oCaptures.Item(1), oCaptures.Item(2), "[object]", MERGE_Image1TopImage2)


                For i = 3 To oCaptures.Count
                    oFinalImage = _CombineImages(oFinalImage, oCaptures.Item(i), "[object]", MERGE_Image1TopImage2)
                Next
            Else
                oFinalImage = oCaptures.Item(1)
            End If

            oFinalImage = _AddTextToImage(oFinalImage, Text)

            _CaptureIEFrames = _SaveImageToFile(oFinalImage, destination)
            End If
    End Function

    Private Function _EnumIEFramesDocument(ByVal wb As HTMLDocumentClass) As Collection
        Dim pContainer As olelib.IOleContainer = Nothing
        Dim pEnumerator As olelib.IEnumUnknown = Nothing
        Dim pUnk As olelib.IUnknown = Nothing
        Dim pBrowser As SHDocVW.IWebBrowser2 = Nothing
        Dim pFramesDoc As Collection = New Collection

        _EnumIEFramesDocument = Nothing

        pContainer = wb

        Dim i As Integer = 0

        ' Get an enumerator for the frames
        If pContainer.EnumObjects(olelib.OLECONTF.OLECONTF_EMBEDDINGS, pEnumerator) = 0 Then

            pContainer = Nothing

            ' Enumerate and refresh all the frames
            Do While pEnumerator.Next(1, pUnk) = 0

                On Error Resume Next

                ' Clear errors
                Err.Clear()

                ' Get the IWebBrowser2 interface
                pBrowser = pUnk

                If Err.Number = 0 Then
                    pFramesDoc.Add(pBrowser.Document)
                    i = i + 1
                End If

            Loop

            pEnumerator = Nothing

        End If

        _EnumIEFramesDocument = pFramesDoc
    End Function
    Private Function _GetScrollObjectFromSource(ByVal Source) As Object
        _GetScrollObjectFromSource = Nothing

        Try
            Select Case TypeName(Source)
                Case "IWebBrowser2", "HTMLFrameElementClass"
                    Dim temp As SHDocVW.IWebBrowser2
                    temp = Source
                    'IE COM iterface return the same way
                    If temp.TopLevelContainer Then
                        _GetScrollObjectFromSource = _GetIEFromSource(temp.Container)
                    Else
                        _GetScrollObjectFromSource = temp
                    End If
                    temp = Nothing
                Case "ImageLink", "TextLink", "HtmlViewLink", "HtmlArea", "HtmlButton", _
                    "HtmlCheckBox", "HtmlEdit", "CoWebElement", "HtmlFile", "HtmlList", "HtmlRadioGroup", _
                    "HtmlTable", "CoWebXML"
                    _GetScrollObjectFromSource = Source.object
                Case "Step"
                    _GetScrollObjectFromSource = Source.object.documentElement
                Case "CoFrame"
                    _GetScrollObjectFromSource = Source.object.body
                Case "CoBrowser"
                    'This is a QTP page object. Object property
                    'will give the document object
                    _GetScrollObjectFromSource = Nothing
                Case "HTMLWindow2"
                    _GetScrollObjectFromSource = _GetIEObjectFromProperty("document.parentWindow", Source.top)
                Case "Interger", "Long", "Int32"
                    _GetScrollObjectFromSource = Nothing
                Case "HTMLDocumentClass"
                    _GetScrollObjectFromSource = Source.body
                Case "HTMLIFrameClass"
                    Dim temp As SHDocVW.IWebBrowser2
                    temp = Source
                    _GetScrollObjectFromSource = temp.Document.body
                    temp = Nothing
                Case Else
                    If TypeName(Source).StartsWith("HTML", StringComparison.OrdinalIgnoreCase) Then
                        _GetScrollObjectFromSource = Source
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Function

    Private Function _GetIEFromSource(ByVal source As Object) As Object
        _GetIEFromSource = Nothing

        Try
            Select Case TypeName(source)
                Case "IWebBrowser2", "HTMLFrameElementClass"
                    Dim temp As SHDocVW.IWebBrowser2
                    temp = source
                    'IE COM iterface return the same way
                    If Not temp.TopLevelContainer Then
                        _GetIEFromSource = _GetIEFromSource(temp.Container)
                    Else
                        _GetIEFromSource = temp
                    End If
                    temp = Nothing
                Case "ImageLink", "TextLink", "HtmlViewLink", "HtmlArea", "HtmlButton", _
                    "HtmlCheckBox", "HtmlEdit", "CoWebElement", "HtmlFile", "HtmlList", "HtmlRadioGroup", _
                    "HtmlTable", "CoWebXML", "Step", "CoFrame"
                    _GetIEFromSource = _GetIEFromSource(source.object)
                Case "CoBrowser"
                    'This is a QTP page object. Object property
                    'will give the document object
                    _GetIEFromSource = source.object
                Case "HTMLWindow2"
                    _GetIEFromSource = _GetIEObjectFromProperty("document.parentWindow", source.top)
                Case "Integer", "Long", "Int32"
                    _GetIEFromSource = _GetIEObjectFromProperty("hwnd", source)
                Case "HTMLDocumentClass"
                    _GetIEFromSource = _GetIEObjectFromProperty("document.parentWindow", source.parentWindow.top)
                Case Else
                    If TypeName(source).StartsWith("HTML", StringComparison.OrdinalIgnoreCase) Then
                        _GetIEFromSource = _GetIEObjectFromProperty("document.parentWindow", source.document.parentWindow.top)
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Function


    Private Function _CreateIEPartialCaptures(ByVal IE As Object, Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False, Optional ByVal scrollingObject As Object = Nothing) As CompleteCapture
        _CreateIEPartialCaptures = Nothing

        IE.Visible = True
        Dim IETabHwnd As Int32 = _GetIEWindow(IE.hwnd)
        'Dim scrollingObject As Object
        'scrollingObject = IE.document.forms.f.q

        If scrollingObject Is Nothing Then
            If (IE.document.body.clientHeight = 0) Then
                scrollingObject = IE.document.documentElement
            ElseIf (IE.document.documentElement.clientHeight = 0) Then
                scrollingObject = IE.document.body
            ElseIf (IE.document.body.clientHeight > IE.document.documentElement.clientHeight) Then
                scrollingObject = IE.document.documentElement
            ElseIf (IE.document.documentElement.clientHeight >0) then
                scrollingObject = IE.document.documentElement
            Else
                scrollingObject = IE.document.body
            End If
        End If
        'scrollingObject = IE.document.body
        'If scrollingObject.clientHeight = 0 Then scrollingObject = scrollingObject.document.body

        Dim objectOffsetX As Int32, objectOffsetY As Int32
        scrollingObject.scrollIntoView()
        If scrollingObject.getClientRects().length <> 0 Then
          objectOffsetX = scrollingObject.getClientRects()(0).Left
          objectOffsetY = scrollingObject.getClientRects()(0).Top
        End If


        Dim clientHeight As Integer, clientWidth As Integer, pageHeight As Integer, pageWidth As Integer
        Dim DeltaX As Integer, DeltaY As Integer, numHScroll As Integer, numVScroll As Integer
        Dim Height As Integer, Width As Integer

        'Get the size of page currently displayed
        clientHeight = scrollingObject.clientHeight
        clientWidth = scrollingObject.clientWidth

        Dim StartPosLeft As Integer, StartPosTop As Integer
        If CaptureIEFromCurrentPos Then
            StartPosLeft = scrollingObject.scrollLeft
            StartPosTop = scrollingObject.scrollTop
        Else
            StartPosLeft = 0
            StartPosTop = 0
        End If

        'Get the total size of the page 
        pageHeight = scrollingObject.scrollHeight - StartPosTop
        pageWidth = scrollingObject.scrollWidth - StartPosLeft

        If clientHeight = 0 Then clientHeight = pageHeight
        If clientWidth = 0 Then clientWidth = pageWidth

        If pageWidth < clientWidth Then pageWidth = clientWidth
        If pageHeight < clientHeight Then pageHeight = clientHeight

        If pageHeight = 0 Or clientHeight = 0 Then Exit Function

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
            DeltaY = 0
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
            DeltaX = 0
            'No. of horizontal scrolls required to cover the whole page
            numHScroll = pageWidth \ (clientWidth - DeltaX) + 1
        End If

        'Variable to store the no. of compeleted scrolls
        Dim curHScroll, curVScroll

        'Scroll the page top most position
        scrollingObject.ScrollTop = StartPosTop

        curVScroll = 0

        Dim TempPartialCapture() As PartialCapture
        ReDim TempPartialCapture(-1)
        Do
            curHScroll = 0

            'scroll the page to left most position
            scrollingObject.scrollLeft = StartPosLeft
            Do
                curHScroll = curHScroll + 1

                ReDim Preserve TempPartialCapture(UBound(TempPartialCapture) + 1)

                TempPartialCapture(UBound(TempPartialCapture)) = New PartialCapture

                With TempPartialCapture(UBound(TempPartialCapture))
                    'X coordinate of the current image on the final image
                    .x = scrollingObject.ScrollLeft - StartPosLeft

                    'Y coordinate of the current image on the final image
                    .y = scrollingObject.scrollTop - StartPosTop

                    'Width of the current capture
                    .width = clientWidth

                    'Height of the current capture
                    .height = clientHeight

                    'Capture the currently displayed part of the page
                    .oImage = _CaptureWindowRect(IETabHwnd, objectOffsetX, objectOffsetY, clientWidth + 2, clientHeight + 2, "[object]") '_GetImageFromHwnd(IETabHwnd)

                    'Scroll to the right of the page by (clientWidth - deltaX)
                    scrollingObject.scrollLeft = scrollingObject.scrollLeft + clientWidth - DeltaX

                End With
                'Loop until all required number of horizontal scrolls are completed
            Loop While curHScroll < numHScroll

            curVScroll = curVScroll + 1

            'Scroll down on the page by (clientHeight - deltaY)
            scrollingObject.scrollTop = scrollingObject.scrollTop + clientHeight - DeltaY

            'Loop until all required number of vertical scrolls are complete
        Loop While curVScroll < numVScroll

        _CreateIEPartialCaptures.height = Height
        _CreateIEPartialCaptures.width = Width
        _CreateIEPartialCaptures.PartialCaptures = TempPartialCapture
    End Function

    'Private Function _GetIEFromSource(ByVal source As Object) As Object
    '    _GetIEFromSource = Nothing

    '    Try
    '        Select Case TypeName(source)
    '            Case "IWebBrowser2"
    '                'IE COM iterface return the same way
    '                _GetIEFromSource = source
    '            Case "Step"
    '                'This is a QTP page object. Object property
    '                'on parent will give the COM interface
    '                _GetIEFromSource = source.GetTOProperty("parent").object
    '            Case "CoBrowser"
    '                'This is a QTP page object. Object property
    '                'will give the document object
    '                _GetIEFromSource = source.object
    '            Case Else
    '                _GetIEFromSource = _GetIEObjectFromHWND(source)
    '        End Select
    '    Catch ex As Exception

    '    End Try
    'End Function

    ''' <summary>
    ''' Get "InternetExplorer.Application" type object from the hwnd of IE window
    ''' or a control inside IE window
    ''' </summary>
    ''' <param name="hwnd">Handle of object inside the IE window handle</param>
    ''' <returns>Internet explorer COM object.</returns>
    ''' <remarks>
    ''' Returns Nothing in case no IE is found
    ''' </remarks>
    Public Function GetIEObjectFromHWND(ByVal hwnd As Int32) As Object Implements _ScreenCapture.GetIEObjectFromHWND
        GetIEObjectFromHWND = _GetIEObjectFromHWND(hwnd)
    End Function

    Private Function _GetIEObjectFromHWND(ByVal hwnd As Int32) As Object
        _GetIEObjectFromHWND = Nothing
        hwnd = _GetRootWindow(hwnd)
        Dim allIE As Object, oIE As Object
        Try
            allIE = CreateObject("Shell.Application").windows
            For Each oIE In allIE
                If oIE.hwnd = hwnd Then
                    _GetIEObjectFromHWND = oIE
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
        Call ConvertImage(sourceFile, destFile, OverWriteDestination)
    End Sub

    Private Sub _ConvertImage(ByVal sourceFile As String, ByVal destFile As String, Optional ByVal OverWriteDestination As Boolean = True)
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
                _SaveImageToFile(srcImage, destFile)
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
        SaveImageToFile = _SaveImageToFile(oImage, destFile)
    End Function

    Private Function _SaveImageToFile(ByVal oImage As Object, ByVal destFile As String) As Object
        Select Case TypeName(oImage)
            Case "Bitmap", "Image"
            Case Else
                Err.Raise(vbObjectError + 1, "SaveImageToFile", "Invalid parametere, object needs to be of type bitmap or image")
        End Select

        _SaveImageToFile = False
        Dim fileFormat As String
        fileFormat = Path.GetExtension(destFile).ToLower

        Dim imgFormat As Imaging.ImageFormat
        'Dim oEncoder As Imaging.ImageCodecInfo
        'oEncoder = GetEncoderInfo("image/tiff")
        'Imaging.ImageCodecInfo.GetImageEncoders().
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
            _SaveImageToFile = True
        ElseIf destFile.ToLower().Trim = "[object]" Then
            _SaveImageToFile = oImage.Clone
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
                _SaveImageToFile(oImage, saveDialog.FileName)
                SaveSetting("ScreenCapture", "SaveDailog", "InitDir", Path.GetDirectoryName(saveDialog.FileName))
            End If

            saveDialog.Dispose()
            saveDialog = Nothing
        Else
            oImage.Save(destFile, imgFormat)
            _SaveImageToFile = True
        End If

        imgFormat = Nothing
    End Function

        Private Declare Function SendMessage _
     Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32

        Private Declare Function PostMessage _
     Lib "user32" Alias "PostMessageA" _
     (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32

    Private Sub ScrollWindow(ByVal hwnd As Integer, ByVal ScrollFactor As Integer)
        _ActivateWindow(GetRootWindow(hwnd))
        SendMessage(hwnd, &H20A, (ScrollFactor * 120) << 16,&H6B049A)
    End Sub

    Private Declare Function SetPixel Lib "gdi32" Alias "SetPixel" (ByVal hdc As IntPtr, ByVal X As Int32, ByVal Y As Int32, ByVal col As Int32) As Int32
    Private Declare Function GetPixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As IntPtr, ByVal X As Int32, ByVal Y As Int32) As Int32

    Private Declare Function ScreenToClient Lib "user32" Alias "ScreenToClient" (ByVal hWnd As IntPtr, ByRef lpPoint As Point) As Boolean
    Private Declare Function ClientToScreen Lib "user32" Alias "ClientToScreen" (ByVal hWnd As IntPtr, ByRef lpPoint As Point) As Boolean

<StructLayout(LayoutKind.Sequential)> private Structure SCROLLBARINFO
    public cbSize As Integer
    public rcScrollBar as RECT
    public dxyLineButton As Integer
    public xyThumbTop As Integer
    public xyThumbBottom As Integer
    public reserved As Integer
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)> Public rgstate() As integer
End Structure

    Private Structure MarkerInfo
        Dim hwnd As Integer
        Dim client As Boolean
        Dim MarkerColors() As Integer
        Dim MarkerPattern() As Integer
        Dim WindowWidth As Integer
        Dim WindowHeight As Integer
        Dim MarkerPos As Point
        Dim HorizontalMarker As Boolean

        Public Sub New(ByVal phwnd As Integer, ByVal pclient As Boolean)
            hwnd = phwnd
            client = pclient
            
            'MarkerPattern = New Integer() {&H2AB02, &HCCFFCC, &HFF00FF, &HAA45AA, &H134513}
            MarkerPattern = New Integer() {&HFFFFFF, &HFFFFFF, &HFFFFFF, &HFFFFFF, &HFFFFFF}
        End Sub
    End Structure

    Function GetOptimumMarkerPos(ByVal hwnd As Integer, ByVal client As Boolean) As Point
        Dim clientRect As RECT
        If client Then
            'Change to client
            GetWindowRect(hwnd, clientRect)
        Else
            GetWindowRect(hwnd, clientRect)
        End If

        GetOptimumMarkerPos = New Point(clientRect.Right - clientRect.Left - 30, (clientRect.Bottom - clientRect.Top) / 2 + 1)
    End Function

    Private Function CreateScrollingMarker(ByVal hwnd As Integer, ByVal client As Boolean) As MarkerInfo
        Dim hdc As Integer
        Dim oMI As MarkerInfo = New MarkerInfo(hwnd, client)
        If client Then
            hdc = GetWindowDC(hwnd)
        Else
            hdc = GetWindowDC(hwnd)
        End If

        oMI.MarkerPos = GetOptimumMarkerPos(hwnd, client)
        oMI.MarkerColors = New Integer() {GetPixel(hdc, oMI.MarkerPos.X, oMI.MarkerPos.Y), _
                                                    GetPixel(hdc, oMI.MarkerPos.X + 1, oMI.MarkerPos.Y), _
                                                    GetPixel(hdc, oMI.MarkerPos.X + 2, oMI.MarkerPos.Y), _
                                                    GetPixel(hdc, oMI.MarkerPos.X + 3, oMI.MarkerPos.Y), _
                                                    GetPixel(hdc, oMI.MarkerPos.X + 4, oMI.MarkerPos.Y)}
        'SetPixel(hdc, oMI.MarkerPos.X-1, oMI.MarkerPos.Y, &HFF00FF)
        'SetPixel(hdc, oMI.MarkerPos.X-2, oMI.MarkerPos.Y, &HFAAFFF)
        'SetPixel(hdc, oMI.MarkerPos.X-3, oMI.MarkerPos.Y, &HFCCFFF)
        'SetPixel(hdc, oMI.MarkerPos.X-4, oMI.MarkerPos.Y, &HF33FFF)
        'SetPixel(hdc, oMI.MarkerPos.X-5, oMI.MarkerPos.Y, &HF22FFF)


        SetPixel(hdc, oMI.MarkerPos.X, oMI.MarkerPos.Y, oMI.MarkerColors(0) Xor oMI.MarkerPattern(0))
        SetPixel(hdc, oMI.MarkerPos.X + 1, oMI.MarkerPos.Y, oMI.MarkerColors(1) Xor oMI.MarkerPattern(1))
        SetPixel(hdc, oMI.MarkerPos.X + 2, oMI.MarkerPos.Y, oMI.MarkerColors(2) Xor oMI.MarkerPattern(2))
        SetPixel(hdc, oMI.MarkerPos.X + 3, oMI.MarkerPos.Y, oMI.MarkerColors(3) Xor oMI.MarkerPattern(3))
        SetPixel(hdc, oMI.MarkerPos.X + 4, oMI.MarkerPos.Y, oMI.MarkerColors(4) Xor oMI.MarkerPattern(4))

        ReleaseDC(hwnd, hdc)
        CreateScrollingMarker = oMI
    End Function

    Private Function GetPixelColor(oBitmap As Bitmap, ByVal x As Integer , ByVal y As Integer) As Integer
        Dim oCol as Color= oBitmap.GetPixel(x, y)
        Dim iRed As Integer = CInt(oCol.R)
        Dim iBlue As Integer = CInt(oCol.B)
        Dim iGreen As Integer = CInt(oCol.G)
        GetPixelColor = (iBlue << 16) + (iGreen << 8) + iRed
    End Function

    Private Sub ClearMarker(oMI As MarkerInfo)
        
    End Sub

    Private Function GetScrollingMarkerDelta(ByRef oMI As MarkerInfo) As Integer
        Dim hdc As Integer
        Dim oWindowRect As RECT
        Dim oScreenPos As Point

        If oMI.client Then
            hdc = GetWindowDC(oMI.hwnd)
            GetWindowRect(oMI.hwnd, oWindowRect)
        Else
            hdc = GetWindowDC(oMI.hwnd)
            GetWindowRect(oMI.hwnd, oWindowRect)
        End If

        oMI.WindowHeight = oWindowRect.Bottom - oWindowRect.Top
        oMI.WindowWidth = oWindowRect.Right - oWindowRect.Left
        
        Const DeltaMaxTest As Integer = 250
        Dim iYMin As Integer = oMI.MarkerPos.Y - DeltaMaxTest
        'If iYMin < 0 Then iYMin = 0
        iYMin = 0
        Dim iYMax As Integer = oMI.WindowHeight 'oMI.MarkerPos.Y

        oScreenPos.X = oMI.MarkerPos.X
        oScreenPos.Y = iYMax 
        ClientToScreen(GetRootWindow(oMI.hwnd), oScreenPos)

        Dim oScreenBMP As Bitmap
        oScreenBMP = New Bitmap(5, iYMax - iYMin)
        Dim oGR As Graphics = Graphics.FromImage(oScreenBMP)
        oGR.CopyFromScreen(oScreenPos.X + oWindowRect.Left, oWindowRect.Left, 0, 0, New Size(5, oMI.WindowHeight))
        'Dim hdc2 As Integer
        
        'hdc2 = oGR.GetHdc()
        
        'oScreenBMP.Save("c:\test.png")
        GetScrollingMarkerDelta = 0
        Dim i As Integer
        i = oScreenBMP.Height
        For j As Integer = iYMax To iYMin + 1 Step -1
            i = i - 1
            Debug.Print (i.ToString() + " - " + oScreenBMP.GetPixel(3, i).ToString())
            If (GetPixelColor(oScreenBMP, 0, i) XOR oMI.MarkerColors(0)) = oMI.MarkerPattern(0) Then
            If (GetPixelColor(oScreenBMP, 1, i) XOR oMI.MarkerColors(1)) = oMI.MarkerPattern(1) Then
            If (GetPixelColor(oScreenBMP, 2, i) XOR oMI.MarkerColors(2)) = oMI.MarkerPattern(2) Then
            If (GetPixelColor(oScreenBMP, 3, i) XOR oMI.MarkerColors(3)) = oMI.MarkerPattern(3) Then
            If (GetPixelColor(oScreenBMP, 4, i) XOR oMI.MarkerColors(4)) = oMI.MarkerPattern(4) Then
                Sleep(50)
                GetScrollingMarkerDelta = Math.Abs(j - oMI.MarkerPos.Y - 1)
                
                'SetPixel(hdc, oMI.MarkerPos.X, i, oMI.MarkerColors(0))
                'SetPixel(hdc, oMI.MarkerPos.X + 1, i, oMI.MarkerColors(0))
                'SetPixel(hdc, oMI.MarkerPos.X + 2, i, oMI.MarkerColors(0))
                'SetPixel(hdc, oMI.MarkerPos.X + 3, i, oMI.MarkerColors(0))
                'SetPixel(hdc, oMI.MarkerPos.X + 4, i, oMI.MarkerColors(0))
                
                
                If (GetScrollingMarkerDelta = 0) then
                    GetScrollingMarkerDelta =0 
                End If
                Exit For
            End If
            End If
            End If
            End If
            End If
        Next
        'oGR.ReleaseHdc(hdc2)
        ReleaseDC(oMI.hwnd, hdc)
    End Function
    
    Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As IntPtr, ByVal fnBar As ScrollBarDirection, ByRef lpsi As SCROLLINFO) As Integer
    
    <StructLayout(LayoutKind.Sequential)>Private Structure SCROLLINFO
    Public cbSize As Integer
    Public fMask As Integer
    Public nMin As Integer
    Public nMax As Integer
    Public nPage As Integer
    Public nPos As Integer
    Public nTrackPos As Integer
    End Structure

    Private Enum ScrollBarDirection
        SB_HORZ = 0
        SB_VERT = 1
        SB_CTL = 2
        SB_BOTH = 3
    End Enum

    Private Enum ScrollInfoMask
        SIF_RANGE = &H1
        SIF_PAGE = &H2
        SIF_POS = &H4
        SIF_DISABLENOSCROLL = &H8
        SIF_TRACKPOS = &H10
        SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
    End Enum

    Private Function _GetImageFromHwnd(ByVal hwnd As Int32, Optional ByVal client As Boolean = False, Optional ByVal VerticalScroll As Boolean = False, Optional ByVal HorizontalScroll As Boolean = False) As Bitmap
        If IsWindow(hwnd) = 0 Then
            Err.Raise(vbObjectError + 1, "GetImageFromHwnd", "Window handle is not valid")
        End If

        If IsWindowVisible(hwnd) = 0 Then
            Err.Raise(vbObjectError + 1, "GetImageFromHwnd", "Window handle not visible")
        End If

        ShowCursor(False)
        HideCaret(hwnd)

        Dim windowRECT As RECT
        Dim windowDC As Int32
        Dim oCompleteCapture As CompleteCapture = New CompleteCapture()
        If client Then
            GetClientRect(hwnd, windowRECT)
        Else
            GetWindowRect(hwnd, windowRECT)
        End If


        oCompleteCapture.width = windowRECT.Right - windowRECT.Left
        oCompleteCapture.height = windowRECT.Bottom - windowRECT.Top

        Dim oPartialCapture As PartialCapture
        oPartialCapture = New PartialCapture
        oPartialCapture.x = 0
        oPartialCapture.y = 0
        oPartialCapture.width = oCompleteCapture.width
        oPartialCapture.height = oCompleteCapture.height
        ReDim oCompleteCapture.PartialCaptures(0)

        Dim bNeedToScroll As Boolean = False
        Dim oBitmap As Bitmap
        Dim oGraphics As Graphics

        If ActivateWindowOnCapture Then
            _ActivateWindow(GetRootWindow(hwnd))
        End If

        Dim bReAttemptScrollCount As Integer = 0

        'Do
            'Sleep(WaitBeforeCapture)
            If client Then
                windowDC = GetDC(hwnd)
            Else
                windowDC = GetWindowDC(hwnd)
            End If


            oBitmap = New Bitmap(windowRECT.Right - windowRECT.Left, windowRECT.Bottom - windowRECT.Top)
            'Bitmap bitmap = new Bitmap(rc.Width, rc.Height);

            oGraphics = Graphics.FromImage(oBitmap)
            'Graphics gfxBitmap = Graphics.FromImage(bitmap);

            '// get a device context for the bitmap
            'IntPtr hdcBitmap = gfxBitmap.GetHdc();
            Dim hdcBitmap As IntPtr = oGraphics.GetHdc()

            '// get a device context for the window
            'IntPtr hdcWindow = Win32.GetWindowDC(hWnd); 
            Dim hdcWindow As IntPtr = windowDC

            If hwnd = GetDesktopWindow() or GetParent(hwnd) <> 0 Then
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
            
        '    If VerticalScroll Then
        '        If ActivateWindowOnCapture Then
        '            _ActivateWindow(GetRootWindow(hwnd))
        '        End If
        '        Dim oMI As MarkerInfo = CreateScrollingMarker(GetRootWindow(hwnd), client)
        '        Dim iDelta As Integer = 0
        '        Sleep(2000)
        '        '_ActivateWindow(hwnd)
        '        ScrollWindow(hwnd, -1)
        '        'Sleep(WaitBeforeCapture)
        '        Dim oTime As Date = Now()
        '        iDelta = GetScrollingMarkerDelta(oMI)
        '        Debug.Print ("Delta initial is - " + iDelta.ToString())
        '        'Dim extraScroll As Integer
        '        'extraScroll = (oMI.WindowHeight/iDelta) - 1
        '        'oMI = CreateScrollingMarker(hwnd, client)
        '        'While extraScroll > 0
        '        '    extraScroll = extraScroll - 1
        '        '    ScrollDown(hwnd)
        '        'End While
        '        'oMI = CreateScrollingMarker(hwnd, client)
        '        'ScrollWindow(hwnd, 1)
        '        'iDelta = GetScrollingMarkerDelta(oMI)
        '        'oBitmap.Save("C:\main.png")
                
        '        Debug.Print ("Delta New is - " + iDelta.ToString())
        '        'Sleep (2000)
        '        Debug.Print ("GetScrollingMarkerDelta took - " + DateDiff(DateInterval.Second,oTime,Now()).ToString())
        '        If iDelta = 0 Then
        '            If bReAttemptScrollCount < 2 Then
        '                'This could be a false trigger. Check once more just to make sure
        '                'Don't retry after this
        '                bReAttemptScrollCount += 1
        '                bNeedToScroll = True
        '                Continue Do
        '            End If
        '            ReDim Preserve oCompleteCapture.PartialCaptures(oCompleteCapture.PartialCaptures.Length - 2)
        '            bNeedToScroll = False
        '        Else
        '            'we found something so retry should continue
        '            bReAttemptScrollCount = 0
                    
        '            bNeedToScroll = True
        '            oPartialCapture.oImage = oBitmap
        '            oBitmap.Save ("C:\IMG_" + oCompleteCapture.PartialCaptures.Length.ToString() + ".png")
        '            oCompleteCapture.PartialCaptures(oCompleteCapture.PartialCaptures.Length - 1) = oPartialCapture
        '            ReDim Preserve oCompleteCapture.PartialCaptures(oCompleteCapture.PartialCaptures.Length)
        '            Dim oldPartialCapture As PartialCapture
        '            oldPartialCapture = oPartialCapture
        '            oPartialCapture = New PartialCapture
        '            oPartialCapture.x = oldPartialCapture.x
        '            oPartialCapture.y = oldPartialCapture.y + iDelta
        '            oPartialCapture.width = oldPartialCapture.width
        '            oPartialCapture.height = oldPartialCapture.height
        '            oCompleteCapture.height += iDelta

        '            oCompleteCapture.PartialCaptures(oCompleteCapture.PartialCaptures.Length - 1) = oPartialCapture
        '        End If
        '    Else
        '        bNeedToScroll = False
        '        oPartialCapture.oImage = oBitmap
        '        oCompleteCapture.PartialCaptures(0) = oPartialCapture
        '    End If
        'Loop While bNeedToScroll

        '_GetImageFromHwnd = _CombinePartialImages(oCompleteCapture) 'oBitmap
        _GetImageFromHwnd = oBitmap

        ShowCursor(True)
        ShowCaret(hwnd)
        'oBitmap.Save("c:\tests.png", Imaging.ImageFormat.Png)
    End Function

    ''' <summary>
    ''' Get handle of the top level window using a handle object inside the window
    ''' </summary>
    ''' <param name="hwnd">Handle of the window</param>
    ''' <returns>Handle of the top level window</returns>
    ''' <remarks></remarks>
    Public Function GetRootWindow(ByVal hwnd As Int32) As Int32 Implements _ScreenCapture.GetRootWindow
        GetRootWindow = _GetRootWindow(hwnd)
    End Function

    Private Function _GetRootWindow(ByVal hwnd As Int32) As Int32
        'Do
        'GetRootWindow = hwnd
        'Debug.Print(hwnd)
        'hwnd = GetParent(hwnd)
        'Loop While hwnd <> 0
        _GetRootWindow = GetAncestor(hwnd, GA_ROOTOWNER)
    End Function

    Private Sub _ActivateWindow(ByVal hwnd As Int32)
        Dim topWindow As Int32
        topWindow = _GetRootWindow(hwnd)

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

    Private Function _GetIEWindow(ByVal hwnd As Int32) As Int32
        Dim IEhWnd As Int32 = _GetRootWindow(hwnd)
        Dim hwnd_Server As Int32

        hwnd_Server = FindWindowEx(IEhWnd, 0, "Frame Tab", vbNullString)

        If hwnd_Server = 0 Then
            'It is IE7 or IE6 which has 2 pixels extra lines in capture
            Me.CombineDeltaX = 2
            Me.CombineDeltaY = 2
            hwnd_Server = FindWindowEx(IEhWnd, 0, "TabWindowClass", vbNullString)
        Else
            'It is IE8 which has 4 pixels extra lines in capture
            Me.CombineDeltaX = 0
            Me.CombineDeltaY = 0
            hwnd_Server = FindWindowEx(hwnd_Server, 0, "TabWindowClass", vbNullString)
        End If

        If hwnd_Server = 0 Then
            hwnd_Server = FindWindowEx(IEhWnd, 0, "Shell DocObject View", vbNullString)
        Else
            Dim hwnd_Server2 As Int32
            hwnd_Server2 = 0
            Do
                hwnd_Server2 = FindWindowEx(hwnd_Server, hwnd_Server2, "Shell DocObject View", vbNullString)
                If IsWindowVisible(hwnd_Server2) <> 0 Then
                    hwnd_Server = hwnd_Server2
                    Exit Do
                End If
            Loop While hwnd_Server2 <> 0
        End If

        If hwnd_Server <> 0 Then
            hwnd_Server = FindWindowEx(hwnd_Server, 0, "Internet Explorer_Server", vbNullString)
        End If

        _GetIEWindow = hwnd_Server
    End Function

    Public Function CaptureWindowRect(ByVal hwnd As Int32, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32, ByVal destination As String, Optional ByVal Text As String = "") Implements _ScreenCapture.CaptureWindowRect
        CaptureWindowRect = _CaptureWindowRect(hwnd, x, y, width, height, destination, Text)
    End Function

    Private Function _CaptureWindowRect(ByVal hwnd As Int32, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32, ByVal destination As String, Optional ByVal Text As String = "")
        Dim oBitmap As Bitmap
        oBitmap = CaptureWindow(hwnd, "[object]", False, False, False)
        Dim newBitmap As Bitmap = New Bitmap(CInt(width), CInt(height))
        Dim oGraphics As Graphics = Graphics.FromImage(newBitmap)
        oGraphics.DrawImage(oBitmap, New Rectangle(0, 0, width, height), New Rectangle(x, y, width, height), GraphicsUnit.Pixel)
        oGraphics.Flush()
        oGraphics.Dispose()
        oBitmap.Dispose()
        newBitmap = _AddTextToImage(newBitmap, Text)
        _CaptureWindowRect = _SaveImageToFile(newBitmap, destination)
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
    '        Dim hdc1 As Int32, hdc2 As Int32
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
    Public Function CaptureDesktopRect(ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32, ByVal destination As String, Optional ByVal text As String = "") Implements _ScreenCapture.CaptureDesktopRect
        CaptureDesktopRect = _CaptureDesktopRect(x, y, width, height, destination, text)
    End Function

    Private Function _CaptureDesktopRect(ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32, ByVal destination As String, Optional ByVal text As String = "")
        _CaptureDesktopRect = _CaptureWindowRect(GetDesktopWindow, x, y, width, height, destination, text)
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
        Call _GetCursorPosition(X, Y)
    End Sub

    Private Sub _GetCursorPosition(ByRef X As Object, ByRef Y As Object)
        Dim pos As POINTAPI

        GetCursorPos(pos)

        X = pos.x

        Y = pos.y
    End Sub


    Public Function CombineImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal Position As Integer = MERGE_Image1TopImage2, Optional ByVal Padding As Integer = 0) Implements _ScreenCapture.CombineImages
        CombineImages = _CombineImages(SrcImage1, SrcImage2, Destination, Position, Padding)
    End Function

    Private Function _CombineImages(ByVal SrcImage1 As Object, ByVal SrcImage2 As Object, ByVal Destination As String, Optional ByVal Position As Integer = MERGE_Image1TopImage2, Optional ByVal Padding As Integer = 0)
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
                'Position = MERGE_Image1BottomImage2
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

        _CombineImages = _SaveImageToFile(bmpDest, Destination)

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
        GetNewBitmap = _GetNewBitmap(Width, Height)
    End Function

    Private Function _GetNewBitmap(ByVal Width As Integer, ByVal Height As Integer) As Bitmap
        Dim RetBmp As Bitmap
        RetBmp = New Bitmap(Width, Height)
        _GetNewBitmap = RetBmp.Clone
        RetBmp.Dispose()
        RetBmp = Nothing
    End Function

    Public Function FindWindowLike(Optional ByVal WindowTitle As String = "*", Optional ByVal WindowClass As String = "*") Implements _ScreenCapture.FindWindowLike
        FindWindowLike = _FindWindowLike(WindowTitle, WindowClass)
    End Function

    Private Function _FindWindowLike(Optional ByVal WindowTitle As String = "*", Optional ByVal WindowClass As String = "*")
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

        _FindWindowLike = matchWindows

        WindowsEnum = Nothing
    End Function

End Class
