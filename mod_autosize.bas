Attribute VB_Name = "mod_autosize"
Option Explicit

'************************************
'PRIVATE CONSTANTS
'************************************
    '------------------------------------------
    'AXIS INDEX TABLE
    '------------------------------------------
    Global Const size_axis_x = 1
    Global Const size_axis_y = 2

'************************************
'PRIVATE DATA TYPES
'************************************
    '------------------------------------------
    'QUICK REPLACEMENT FOR AN OBJECT
    'DATA TYPE, HOLDS ALL NECESSARY
    'DATA
    '------------------------------------------
    Public Type objectinfo
        top As Long
        left As Long
        height As Long
        width As Long
        tag As String
    End Type

'************************************
'PRIVATE DATA TYPES
'************************************
    '------------------------------------------
    'AUTOSIZE SUB
    'OBJECT TAG CONTROLS
    'STRETCHH - STRETHCHES HORIZONTALY
    'STRETCHV - STRETCHES VERTICALY
    'STRETCHALL - STRETCHES BOTH HORIZONTALY AND VERTICALY
    'MOVEH - MOVES THE OBJECT HORIZONTALY
    'MOVEV - MOVES THE OBJECT VERTICALY
    'MOVEALL - MOVES THE OBJECT BOTH HORIZONTALY AND VERTICALY
    'STRETCHVMOVEH - STRETCHES VERTICALY AND MOVES THE OBJECT HORIZONTALLY
    'STRETCHHMOVEV - STRETCHES HORIZONTALLY AND MOVES THE OBJECT VERTICALY
    '------------------------------------------
    Public Sub Autosizeform(ByRef sizeobjects() As objectinfo, ByRef effectobjects() As Object, ByRef firstwidth As Long, ByRef firstheight As Long, ByRef noofobjects As Integer, sizeform As Form, Optional forcereset As Boolean, Optional axis As Integer)
        If sizeform.WindowState = vbMinimized Then Exit Sub
        Dim getobject As Object             'CURRENT OBJECT BEING RETRIEVED
        Dim setobject As Integer            'CURRENT OBJECT BEING RESCALED
        Dim restricted As Boolean           'TRUE IF AXIS VARIABLE IS USED
        
        '------------------------------------------
        'GET THE AXIS VARIABLE AND SET RESTRICTED
        'TO TRUE IF IT IS ABOVE 0.  THIS WILL CAUSE
        'ONLY SELECTED AXIS'S TO BE SCALES, HANDY
        'TO STOP THE OBJECTS BEING SCALED BELOW 0
        If axis > 0 Then restricted = True
        
        '------------------------------------------
        'THIS BIT RESETS ALL THE SAVED INFORMATION
        'IF NO PREVIOUSLY FOUND OBJECTS WERE FOUND
        'AND IF THE FORCERESET FLAG IS HIGH
        If noofobjects = 0 Or forcereset Then
            For Each getobject In sizeform
                If getobject.tag <> "" Then
                    '-------------------------------------------
                    'INCREASE THE NUMBER OF FOUND OBJECTS
                    noofobjects = noofobjects + 1
                    '-------------------------------------------
                    'REDIM AND SAVE THE OBJECT POSITIONS TO
                    'THE NEW NUMBER OF OBJECTS FOUND
                    ReDim Preserve sizeobjects(noofobjects)
                    sizeobjects(noofobjects).top = getobject.top
                    sizeobjects(noofobjects).left = getobject.left
                    sizeobjects(noofobjects).width = getobject.width
                    sizeobjects(noofobjects).height = getobject.height
                    sizeobjects(noofobjects).tag = getobject.tag
                    '-------------------------------------------
                    'REDIM THE EFFECT OBJECTS AND SET THE
                    'NEW FOUND OBJECT, USING PRESERVE TO KEEP
                    'ALL EXISTING OBJECTS IN THE ARRAY
                    ReDim Preserve effectobjects(noofobjects)
                    Set effectobjects(noofobjects) = getobject
                    '-------------------------------------------
                End If
            Next
            '-------------------------------------------
            'GET THE ORIGINAL HEIGHT AND WIDTH OF THE
            'FORM BEING AUTOSIZED
            firstheight = sizeform.height
            firstwidth = sizeform.width
            '-------------------------------------------
        End If
        '------------------------------------------
        
        '------------------------------------------
        'THIS BIT LOOKS AT THE TAG PART OF THE SAVED
        'OBJECTS AND SCALES THEM AS NECESSARY
        'NO SCALING ALGORYTHMS ARE ACTUALLY USED,
        'THE ROUTINE WORKS BE KEEPING THE RIGHTHAND
        'GAP THE SAME
        If noofobjects >= 1 Then
            For setobject = 1 To noofobjects
                If sizeobjects(setobject).tag <> "" Then effectobjects(setobject).Visible = False
                Select Case sizeobjects(setobject).tag
                    Case "STRETCHH"
                        If restricted = True And axis <> size_axis_x Then GoTo nextobject
                        effectobjects(setobject).width = sizeform.width - (firstwidth - (sizeobjects(setobject).left + sizeobjects(setobject).width)) - sizeobjects(setobject).left
                    Case "STRETCHV"
                        If restricted = True And axis <> size_axis_y Then GoTo nextobject
                        effectobjects(setobject).height = sizeform.height - (firstheight - (sizeobjects(setobject).top + sizeobjects(setobject).height)) - sizeobjects(setobject).top
                    Case "STRETCHALL"
                        If Not restricted Then
                            effectobjects(setobject).width = sizeform.width - (firstwidth - (sizeobjects(setobject).left + sizeobjects(setobject).width)) - sizeobjects(setobject).left
                            effectobjects(setobject).height = sizeform.height - (firstheight - (sizeobjects(setobject).top + sizeobjects(setobject).height)) - sizeobjects(setobject).top
                        Else
                            If axis = size_axis_x Then
                                effectobjects(setobject).width = sizeform.width - (firstwidth - (sizeobjects(setobject).left + sizeobjects(setobject).width)) - sizeobjects(setobject).left
                            Else
                                effectobjects(setobject).height = sizeform.height - (firstheight - (sizeobjects(setobject).top + sizeobjects(setobject).height)) - sizeobjects(setobject).top
                            End If
                        End If
                    Case "MOVEH"
                        If restricted = True And axis <> size_axis_x Then GoTo nextobject
                        effectobjects(setobject).left = sizeform.width - (firstwidth - sizeobjects(setobject).left)
                    Case "MOVEV"
                        If restricted = True And axis <> size_axis_y Then GoTo nextobject
                        effectobjects(setobject).top = sizeform.height - (firstheight - sizeobjects(setobject).top)
                    Case "MOVEALL"
                        If Not restricted Then
                            effectobjects(setobject).left = sizeform.width - (firstwidth - sizeobjects(setobject).left)
                            effectobjects(setobject).top = sizeform.height - (firstheight - sizeobjects(setobject).top)
                        Else
                            If axis = size_axis_x Then
                                effectobjects(setobject).left = sizeform.width - (firstwidth - sizeobjects(setobject).left)
                            Else
                                effectobjects(setobject).top = sizeform.height - (firstheight - sizeobjects(setobject).top)
                            End If
                        End If
                    Case "STRETCHVMOVEH"
                        If Not restricted Then
                            effectobjects(setobject).height = sizeform.height - (firstheight - (sizeobjects(setobject).top + sizeobjects(setobject).height)) - sizeobjects(setobject).top
                            effectobjects(setobject).left = sizeform.width - (firstwidth - sizeobjects(setobject).left)
                        Else
                            If axis = size_axis_x Then
                                effectobjects(setobject).left = sizeform.width - (firstwidth - sizeobjects(setobject).left)
                            Else
                                effectobjects(setobject).height = sizeform.height - (firstheight - (sizeobjects(setobject).top + sizeobjects(setobject).height)) - sizeobjects(setobject).top
                            End If
                        End If
                    Case "STRETCHHMOVEV"
                        If Not restricted Then
                            effectobjects(setobject).width = sizeform.width - (firstwidth - (sizeobjects(setobject).left + sizeobjects(setobject).width)) - sizeobjects(setobject).left
                            effectobjects(setobject).top = sizeform.height - (firstheight - sizeobjects(setobject).top)
                        Else
                            If axis = size_axis_x Then
                                effectobjects(setobject).width = sizeform.width - (firstwidth - (sizeobjects(setobject).left + sizeobjects(setobject).width)) - sizeobjects(setobject).left
                            Else
                                effectobjects(setobject).top = sizeform.height - (firstheight - sizeobjects(setobject).top)
                            End If
                        End If
                End Select
nextobject:
                effectobjects(setobject).Visible = True
            Next setobject
        End If
        '------------------------------------------
    End Sub

