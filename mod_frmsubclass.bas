Attribute VB_Name = "mod_frmsubclass"
Option Explicit

'***********************************
'GLOBAL API DECLARATIONS
'***********************************
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'***********************************
'PRIVATE VARIABLES
'***********************************

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
    Public Const GWL_WNDPROC As Long = (-4)
    Public Const WM_GETMINMAXINFO As Long = &H24

'***********************************
'TYPE DECLARATIONS
'***********************************
    'GLOBAL
    Public Type POINTAPI
        x As Long
        y As Long
    End Type

    'PRIVATE
    Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
    End Type

'***********************************
'USER INTERFACE
'***********************************
    Public Sub restrictform(resrictform As Form)
        Dim startupwidth As Long
        Dim startupheight As Long
        startupwidth = resrictform.Width \ Screen.TwipsPerPixelX
        startupheight = resrictform.Height \ Screen.TwipsPerPixelY
        minX = startupwidth
        minY = startupheight
        maxX = Screen.Width \ Screen.TwipsPerPixelX
        maxY = Screen.Height \ Screen.TwipsPerPixelY
        SubClass resrictform.hwnd
    End Sub

'***********************************
'PUBLIC SUB CLASSING/UN SUB CLASSING SUBROUTINES
'***********************************
    Public Sub SubClass(hwnd As Long)
        On Error Resume Next
        defWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    End Sub

    Public Sub UnSubClass(hwnd As Long)
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
