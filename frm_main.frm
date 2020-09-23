VERSION 5.00
Begin VB.Form frm_main 
   Caption         =   "Autosize - How could it get any easier??"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_more 
      Caption         =   "More!!!"
      Height          =   555
      Left            =   7440
      TabIndex        =   12
      Tag             =   "MOVEALL"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      Height          =   555
      Left            =   1500
      TabIndex        =   11
      Tag             =   "STRETCHHMOVEV"
      Text            =   "STRETCHHMOVEV"
      Top             =   5520
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   3165
      IntegralHeight  =   0   'False
      ItemData        =   "frm_main.frx":08CA
      Left            =   6840
      List            =   "frm_main.frx":08D1
      TabIndex        =   10
      Tag             =   "STRETCHVMOVEH"
      Top             =   2280
      Width           =   1995
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "MOVEV"
      Height          =   555
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Tag             =   "MOVEV"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "MOVEALL"
      Height          =   555
      Index           =   2
      Left            =   6000
      TabIndex        =   8
      Tag             =   "MOVEALL"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      Height          =   3195
      Left            =   2100
      TabIndex        =   7
      Tag             =   "STRETCHALL"
      Text            =   "STRETCHALL"
      Top             =   2280
      Width           =   4695
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "MOVEH"
      Height          =   555
      Index           =   1
      Left            =   7440
      TabIndex        =   6
      Tag             =   "MOVEH"
      Top             =   1680
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   1500
      TabIndex        =   5
      Tag             =   "STRETCHH"
      Text            =   "STRETCHH"
      Top             =   1680
      Width           =   5895
   End
   Begin VB.ListBox lst_example 
      Height          =   3165
      IntegralHeight  =   0   'False
      ItemData        =   "frm_main.frx":08E4
      Left            =   60
      List            =   "frm_main.frx":08EB
      TabIndex        =   4
      Tag             =   "STRETCHV"
      Top             =   2280
      Width           =   1995
   End
   Begin VB.CommandButton cmd_example 
      Caption         =   "NO TAG"
      Height          =   555
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Frame fra_information 
      Caption         =   "Useful Information"
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Tag             =   "STRETCHH"
      Top             =   60
      Width           =   8775
      Begin VB.Image img_icon 
         Height          =   480
         Left            =   420
         Picture         =   "frm_main.frx":08F9
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lbl_info2 
         Alignment       =   2  'Center
         Caption         =   $"frm_main.frx":11C3
         Height          =   615
         Left            =   1380
         TabIndex        =   2
         Tag             =   "STRETCHH"
         Top             =   780
         Width           =   7275
      End
      Begin VB.Label lbl_info1 
         Alignment       =   2  'Center
         Caption         =   $"frm_main.frx":12AE
         Height          =   615
         Left            =   1380
         TabIndex        =   1
         Tag             =   "STRETCHH"
         Top             =   180
         Width           =   7275
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'NAME       :   AUTOSIZE
'PROGRAMMER :   NICK PATEMAN
'EMAIL      :   np24@blueyonder.co.uk
'----------------------------------------------------------
'THANKYOU FOR DOWNLOADING AUTOSIZE.  THIS IS THE FIRST
'RELEASE AND AS YET SEEMS TO DO ITS JOB PERFECTLY. THOUGH
'IM SURE SOME OF YOU WILL CORRECT ME ON THAT.  ANYWAYS
'PRESS THAT PLAY BUTTON AND HAVE A LOOK, IM GLAD THAT I
'HAVE DONE THIS, ITS GOING TO BE USED IN ALL MY APPS NOW.
'OH YEAH, YOU WILL FIND PROBLEMS WITH THE LISTBOX BEING
'RESIZED AS THIS HAS SOME KIND OF STRANGE FEATURE CALLED
'INTEGRAL HEIGHT, SO YOU HAVE TO SET IT TO FALSE WHEN YOU
'INCLUDE ONE ON YOUR FORM, THE ONE ON THIS FORM HAS ALREADY
'BEEN MODIFIED.  ANYWAYS, VOTE FOR THIS BECAUSE ITS ABOUT
'TIME SOMEONE DONE ONE LIKE THIS.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

'***********************************
'PRIVATE VARIABLES
'***********************************
    '------------------------------------------
    'FORM OBJECT DATA
    '------------------------------------------
    Dim frm_main_objects() As objectinfo
    Dim frm_main_effectobjects() As Object
    Private frm_main_noofobjects As Integer
    Private frm_main_startwidth As Long
    Private frm_main_startheight As Long

'***********************************
'FORM EVENTS
'***********************************
    '------------------------------------------
    'LOAD EVENT
    '------------------------------------------
    Private Sub Form_Load()
        frm_main_startwidth = Me.width
        frm_main_startheight = Me.height
        restrictform Me
    End Sub
    
    '------------------------------------------
    'UNLOAD EVENT
    '------------------------------------------
    Private Sub Form_Unload(Cancel As Integer)
        unrestrictform Me
    End Sub
    
    '------------------------------------------
    'RESIZE EVENT
    '------------------------------------------
    Private Sub Form_Resize()
        'THIS IS THE ONLY LINE OF CODE YOU NEED WRITE
        'ALL THE REST IS HARD CODED ONTO THE OBJECTS
        'USING THE TAG PROPERTY.  THE SAVED TAGS CAN
        'ALSO BE RESET AT RUNTIME SIMPLY BY USING
        'Autosizeform Me, True
        Autosizeform frm_main_objects, frm_main_effectobjects, frm_main_startwidth, frm_main_startheight, frm_main_noofobjects, Me
    End Sub
    
'***********************************
'COMMAND BUTTONS
'***********************************
    '------------------------------------------
    'SHOW UNSUBCLASSED FORM
    '------------------------------------------
    Private Sub cmd_more_Click()
        frm_test.Show
    End Sub
