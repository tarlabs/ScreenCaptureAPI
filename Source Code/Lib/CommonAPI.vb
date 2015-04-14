Imports System.Runtime.InteropServices
Imports System.Collections.Generic
'Imports System.Runtime.InteropServices
Imports System.Text

Public Class WindowInfo
    Public Title As String = ""
    Public ClassName As String = ""
    Public hWnd As Int32
End Class

Public Class WindowsEnumerator

    Private Delegate Function EnumCallBackDelegate(ByVal hwnd As Integer, ByVal lParam As Integer) As Integer

    ' Top-level windows.
    Private Declare Function EnumWindows Lib "user32" _
     (ByVal lpEnumFunc As EnumCallBackDelegate, ByVal lParam As Integer) As Integer

    ' Child windows.
    Private Declare Function EnumChildWindows Lib "user32" _
     (ByVal hWndParent As Integer, ByVal lpEnumFunc As EnumCallBackDelegate, ByVal lParam As Integer) As Integer

    ' Get the window class.
    Private Declare Function GetClassName _
     Lib "user32" Alias "GetClassNameA" _
     (ByVal hwnd As Integer, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer

    ' Test if the window is visible--only get visible ones.
    Private Declare Function IsWindowVisible Lib "user32" _
     (ByVal hwnd As Integer) As Integer

    ' Test if the window's parent--only get the one's without parents.
    Private Declare Function GetParent Lib "user32" _
     (ByVal hwnd As Integer) As Integer

    ' Get window text length signature.
    Private Declare Function SendMessage _
     Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32

    ' Get window text signature.
    Private Declare Function SendMessage _
     Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As StringBuilder) As Int32

    Private _listChildren As New List(Of WindowInfo)
    Private _listTopLevel As New List(Of WindowInfo)

    Private _topLevelClass As String = ""
    Private _childClass As String = ""

    ''' <summary>
    ''' Get all top-level window information
    ''' </summary>
    ''' <returns>List of window information objects</returns>
    Public Overloads Function GetTopLevelWindows() As List(Of WindowInfo)

        EnumWindows(AddressOf EnumWindowProc, &H0)

        Return _listTopLevel

    End Function

    Public Overloads Function GetTopLevelWindows(ByVal className As String) As List(Of WindowInfo)

        _topLevelClass = className

        Return Me.GetTopLevelWindows()

    End Function

    ''' <summary>
    ''' Get all child windows for the specific windows handle (hwnd).
    ''' </summary>
    ''' <returns>List of child windows for parent window</returns>
    Public Overloads Function GetChildWindows(ByVal hwnd As Int32) As List(Of WindowInfo)

        ' Clear the window list.
        _listChildren = New List(Of WindowInfo)

        ' Start the enumeration process.
        EnumChildWindows(hwnd, AddressOf EnumChildWindowProc, &H0)

        ' Return the children list when the process is completed.
        Return _listChildren

    End Function

    Public Overloads Function GetChildWindows(ByVal hwnd As Int32, ByVal childClass As String) As List(Of WindowInfo)

        ' Set the search
        _childClass = childClass

        Return Me.GetChildWindows(hwnd)

    End Function

    ''' <summary>
    ''' Callback function that does the work of enumerating top-level windows.
    ''' </summary>
    ''' <param name="hwnd">Discovered Window handle</param>
    ''' <returns>1=keep going, 0=stop</returns>
    Private Function EnumWindowProc(ByVal hwnd As Int32, ByVal lParam As Int32) As Int32

        ' Eliminate windows that are not top-level.
        If GetParent(hwnd) = 0 AndAlso CBool(IsWindowVisible(hwnd)) Then

            ' Get the window title / class name.
            Dim window As WindowInfo = GetWindowIdentification(hwnd)

            ' Match the class name if searching for a specific window class.
            If _topLevelClass.Length = 0 OrElse window.ClassName.ToLower() = _topLevelClass.ToLower() Then
                _listTopLevel.Add(window)
            End If

        End If

        ' To continue enumeration, return True (1), and to stop enumeration 
        ' return False (0).
        ' When 1 is returned, enumeration continues until there are no 
        ' more windows left.

        Return 1

    End Function

    ''' <summary>
    ''' Callback function that does the work of enumerating child windows.
    ''' </summary>
    ''' <param name="hwnd">Discovered Window handle</param>
    ''' <returns>1=keep going, 0=stop</returns>
    Private Function EnumChildWindowProc(ByVal hwnd As Int32, ByVal lParam As Int32) As Int32

        Dim window As WindowInfo = GetWindowIdentification(hwnd)

        ' Attempt to match the child class, if one was specified, otherwise
        ' enumerate all the child windows.
        If _childClass.Length = 0 OrElse window.ClassName.ToLower() = _childClass.ToLower() Then
            _listChildren.Add(window)
        End If

        Return 1

    End Function

    ''' <summary>
    ''' Build the WindowInfo object to hold information about the Window object.
    ''' </summary>
    Private Function GetWindowIdentification(ByVal hwnd As Integer) As WindowInfo

        Const WM_GETTEXT As Int32 = &HD
        Const WM_GETTEXTLENGTH As Int32 = &HE

        Dim window As New WindowInfo()

        Dim title As New StringBuilder()

        ' Get the size of the string required to hold the window title.
        Dim size As Int32 = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)

        ' If the return is 0, there is no title.
        If size > 0 Then
            title = New StringBuilder(size + 1)

            SendMessage(hwnd, WM_GETTEXT, title.Capacity, title)
        End If

        ' Get the class name for the window.
        Dim classBuilder As New StringBuilder(64)
        GetClassName(hwnd, classBuilder, 64)

        ' Set the properties for the WindowInfo object.
        window.ClassName = classBuilder.ToString()
        window.Title = title.ToString()
        window.hWnd = hwnd

        Return window

    End Function

End Class

Module Win32API
    Public Const GA_ROOT As Int32 = 2
    Public Declare Function GetAncestor Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByVal gaFlags As Int32) As Int32
    Public Const GA_PARENT As Int32 = 1
    Public Const GA_ROOTOWNER As Int32 = 3

    Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
      ByVal hwndParent As Int32, _
      ByVal hwndChildAfter As Int32, _
      ByVal lpszClass As String, _
      ByVal lpszWindow As String) As Int32
    Public Declare Function GetWindow Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByVal wCmd As Int32) As Int32
    Public Declare Function PrintWindow Lib "user32.dll" ( _
      ByVal hWnd As Int32, _
      ByVal hdcBlt As Int32, _
      ByVal nFlags As Int32) As Int32
    Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" ( _
      ByVal hwnd As Int32, _
      ByVal lpClassName As String, _
      ByVal nMaxCount As Int32) As Int32
    Public Declare Function GetActiveWindow Lib "user32.dll" () As Int32
    Public Declare Function GetDesktopWindow Lib "user32.dll" () As Int32
    Public Declare Function GetWindowDC Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32
    Public Declare Function GetClientRect Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByRef lpRect As RECT) As Int32
    Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Public Declare Function GetCursorPos Lib "user32.dll" ( _
      ByRef lpPoint As POINTAPI) As Int32

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure POINTAPI
        Public x As Int32
        Public y As Int32
    End Structure



    <StructLayout(LayoutKind.Sequential)> _
    Public Structure RECT
        Public Left As Int32
        Public Top As Int32
        Public Right As Int32
        Public Bottom As Int32
    End Structure

    Public Declare Function IsWindow Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32
    Public Declare Function IsWindowVisible Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32
    Public Declare Function ReleaseDC Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByVal hdc As Int32) As Int32

    Public Declare Function GetDC Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32

    Public Declare Function GetWindowRect Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByRef lpRect As RECT) As Int32
    Public Declare Function BitBlt Lib "gdi32.dll" ( _
      ByVal hDestDC As Int32, _
      ByVal x As Int32, _
      ByVal y As Int32, _
      ByVal nWidth As Int32, _
      ByVal nHeight As Int32, _
      ByVal hSrcDC As Int32, _
      ByVal xSrc As Int32, _
      ByVal ySrc As Int32, _
      ByVal dwRop As Int32) As Int32
    Public Const SRCCOPY As Int32 = &HCC0020

    Public Declare Function GetParent Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32
    Public Const GW_CHILD As Int32 = 5
    Public Const GW_HWNDFIRST As Int32 = 0
    Public Const GW_HWNDLAST As Int32 = 1
    Public Const GW_HWNDNEXT As Int32 = 2
    Public Const GW_HWNDPREV As Int32 = 3
    Public Const GW_OWNER As Int32 = 4

    Public Declare Function GetForegroundWindow Lib "user32.dll" () As Int32

    Public Declare Function ShowWindow Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByVal nCmdShow As Int32) As Int32
    Public Const SW_HIDE As Int32 = 0
    Public Const SW_MAXIMIZE As Int32 = 3
    Public Const SW_MINIMIZE As Int32 = 6
    Public Const SW_NORMAL As Int32 = 1
    Public Const SW_SHOW As Int32 = 5
    Public Const SW_SHOWMINIMIZED = 2
    Public Const SW_RESTORE As Int32 = 9

    Public Declare Function DeleteDC Lib "gdi32.dll" ( _
      ByVal hdc As Int32) As Int32

    Public Declare Function ShowCursor Lib "user32.dll" ( _
      ByVal bShow As Int32) As Int32

    Public Declare Function SetWindowPlacement Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByRef lpwndpl As WINDOWPLACEMENT) As Int32


    Public Declare Function SetForegroundWindow Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32
    Public Declare Function BringWindowToTop Lib "user32.dll" ( _
      ByVal hwnd As Int32) As Int32
    Public Declare Sub Sleep Lib "kernel32.dll" ( _
            ByVal dwMilliseconds As Int32)

    Public Declare Function GetWindowPlacement Lib "user32.dll" ( _
      ByVal hwnd As Int32, _
      ByRef lpwndpl As WINDOWPLACEMENT) As Int32
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure WINDOWPLACEMENT
        Public Length As Int32
        Public flags As Int32
        Public showCmd As Int32
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECT
    End Structure

    'Private Declare Function CreateBitmapIndirect Lib "gdi32.dll" (ByRef lpBitmap As Bitmap) As Long

    Public Const SRCINVERT As Long = &H660046
    Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
    Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

End Module
