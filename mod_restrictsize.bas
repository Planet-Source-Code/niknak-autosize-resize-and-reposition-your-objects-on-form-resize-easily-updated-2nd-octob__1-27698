Attribute VB_Name = "mod_restrictsize"
Option Explicit

'***********************************
'GLOBAL API DECLARATIONS
'***********************************
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'***********************************
'PRIVATE VARIABLES
'***********************************
    Private startupheight As Long
    Private startupwidth As Long

'***********************************
'GLOBAL VARIABLES
'***********************************
    Private defWindowProc As Long
    Private minX As Long
    Private minY As Long
    Private maxX As Long
    Private maxY As Long

'***********************************
'GLOBAL CONSTANTS
'***********************************
    Private Const WM_GETMINMAXINFO As Long = &H24
    Private Const GWL_WNDPROC = (-4)

'***********************************
'TYPE DECLARATIONS
'***********************************
    'GLOBAL
    Private Type POINTAPI
        x As Long
        y As Long
    End Type

    'PRIVATE
    Private Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
    End Type

'***********************************
'USER INTERFACE
'***********************************
    '-------------------------------
    'RESTRICT FORM
    '-------------------------------
    Public Sub restrictform(resrictform As Form)
        Dim startupwidth As Long
        Dim startupheight As Long
        With resrictform
            startupwidth = .width \ Screen.TwipsPerPixelX
            startupheight = .height \ Screen.TwipsPerPixelY
            minX = startupwidth
            minY = startupheight
            maxX = Screen.width \ Screen.TwipsPerPixelX
            maxY = Screen.height \ Screen.TwipsPerPixelY
            SubClass .hwnd
        End With
    End Sub

    '-------------------------------
    'UNRESTRICT FORM
    '-------------------------------
    Public Sub unrestrictform(restrictform As Form)
        UnSubClass restrictform.hwnd
    End Sub

'***********************************
'PRIVATE SUBS
'***********************************
    '-------------------------------
    'START SUBCLASSING
    '-------------------------------
    Private Sub SubClass(hwnd As Long)
        On Error Resume Next
        defWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    End Sub
    
    '-------------------------------
    'END SUBCLASSING
    '-------------------------------
    Private Sub UnSubClass(hwnd As Long)
        If defWindowProc Then
            SetWindowLong hwnd, GWL_WNDPROC, defWindowProc
            defWindowProc = 0
        End If
    End Sub

'***********************************
'WINDOW RESIZING PROCEDURE
'***********************************
    Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Select Case uMsg
            Case WM_GETMINMAXINFO
                Dim MMI As MINMAXINFO
                CopyMemory MMI, ByVal lParam, LenB(MMI)
                With MMI
                    .ptMinTrackSize.x = minX
                    .ptMinTrackSize.y = minY
                    .ptMaxTrackSize.x = maxX
                    .ptMaxTrackSize.y = maxY
                End With
                CopyMemory ByVal lParam, MMI, LenB(MMI)
                WindowProc = 0
            Case Else
                WindowProc = CallWindowProc(defWindowProc, hwnd, uMsg, wParam, lParam)
        End Select
    End Function
