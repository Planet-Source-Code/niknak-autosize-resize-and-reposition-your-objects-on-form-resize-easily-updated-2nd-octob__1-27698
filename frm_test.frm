VERSION 5.00
Begin VB.Form frm_test 
   Caption         =   "Autosize - How could it get any easier??"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra_information 
      Caption         =   "Useful Information"
      Height          =   1575
      Left            =   0
      TabIndex        =   9
      Tag             =   "STRETCHH"
      Top             =   0
      Width           =   8775
      Begin VB.Label lbl_info1 
         Alignment       =   2  'Center
         Caption         =   $"frm_test.frx":0000
         Height          =   615
         Left            =   1380
         TabIndex        =   11
         Tag             =   "STRETCHH"
         Top             =   180
         Width           =   7275
      End
      Begin VB.Label lbl_info2 
         Alignment       =   2  'Center
         Caption         =   $"frm_test.frx":0109
         Height          =   615
         Left            =   1380
         TabIndex        =   10
         Tag             =   "STRETCHH"
         Top             =   780
         Width           =   7275
      End
      Begin VB.Image img_icon 
         Height          =   480
         Left            =   420
         Picture         =   "frm_test.frx":01BE
         Top             =   540
         Width           =   480
      End
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "NO TAG"
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1620
      Width           =   1395
   End
   Begin VB.ListBox lst_example 
      Height          =   3165
      IntegralHeight  =   0   'False
      ItemData        =   "frm_test.frx":0A88
      Left            =   0
      List            =   "frm_test.frx":0A8F
      TabIndex        =   7
      Tag             =   "STRETCHV"
      Top             =   2220
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   1440
      TabIndex        =   6
      Tag             =   "STRETCHH"
      Text            =   "STRETCHH"
      Top             =   1620
      Width           =   5895
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "MOVEH"
      Height          =   555
      Index           =   1
      Left            =   7380
      TabIndex        =   5
      Tag             =   "MOVEH"
      Top             =   1620
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      Height          =   3195
      Left            =   2040
      TabIndex        =   4
      Tag             =   "STRETCHALL"
      Text            =   "STRETCHALL"
      Top             =   2220
      Width           =   4695
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "MOVEALL"
      Height          =   555
      Index           =   2
      Left            =   7380
      TabIndex        =   3
      Tag             =   "MOVEALL"
      Top             =   5460
      Width           =   1395
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "MOVEV"
      Height          =   555
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Tag             =   "MOVEV"
      Top             =   5460
      Width           =   1395
   End
   Begin VB.ListBox List1 
      Height          =   3165
      IntegralHeight  =   0   'False
      ItemData        =   "frm_test.frx":0A9D
      Left            =   6780
      List            =   "frm_test.frx":0AA4
      TabIndex        =   1
      Tag             =   "STRETCHVMOVEH"
      Top             =   2220
      Width           =   1995
   End
   Begin VB.TextBox Text3 
      Height          =   555
      Left            =   1440
      TabIndex        =   0
      Tag             =   "STRETCHHMOVEV"
      Text            =   "STRETCHHMOVEV"
      Top             =   5460
      Width           =   5895
   End
End
Attribute VB_Name = "frm_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***********************************
'PRIVATE VARIABLES
'***********************************
    '------------------------------------------
    'FORM OBJECT DATA
    '------------------------------------------
    Dim frm_test_objects() As objectinfo
    Dim frm_test_effectobjects() As Object
    Private frm_test_noofobjects As Integer
    Private frm_test_startwidth As Long
    Private frm_test_startheight As Long

'***********************************
'FORM EVENTS
'***********************************
    '------------------------------------------
    'RESIZE EVENT
    '------------------------------------------
    Private Sub Form_Load()
        frm_test_startwidth = Me.width
        frm_test_startheight = Me.height
    End Sub
    
    '------------------------------------------
    'RESIZE EVENT
    '------------------------------------------
    Private Sub Form_Resize()
        '------------------------------------------
        'YOU WILL NOTICE THAT THIS RESIZE FUNCTION
        'IS SLIGHTLY MORE COMPLEX, THIS IS ONLY
        'TO STOP THE APPLICATION FROM RESIZING
        'OBJECTS BELOW THEIR BOUNDRIES.  BASICALLY
        'WHEN THE OBJECT IS WIDE ENOUGH, THE WIDTH
        'IS EFFECTED, BUT WHEN ITS NOT, ITS IGNORED
        'THE SAME GOES FOR THE HEIGHT.
        If Me.width >= frm_test_startwidth And Me.height >= frm_test_startheight Then
            Autosizeform frm_test_objects, frm_test_effectobjects, frm_test_startwidth, frm_test_startheight, frm_test_noofobjects, Me
        ElseIf Me.width >= frm_test_startwidth And Me.height <= frm_test_startheight Then
            Autosizeform frm_test_objects, frm_test_effectobjects, frm_test_startwidth, frm_test_startheight, frm_test_noofobjects, Me, , size_axis_x
        ElseIf Me.height >= frm_test_startheight And Me.width <= frm_test_startwidth Then
            Autosizeform frm_test_objects, frm_test_effectobjects, frm_test_startwidth, frm_test_startheight, frm_test_noofobjects, Me, , size_axis_y
        End If
        '------------------------------------------
    End Sub
