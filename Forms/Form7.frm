VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   LinkTopic       =   "Form7"
   ScaleHeight     =   4515
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   6105
      TabIndex        =   1
      Top             =   720
      Width           =   6135
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   3015
         Left            =   2760
         TabIndex        =   3
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CommandButton BtnNext 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Set Splitter1 = New Splitter
    Splitter1.New_ False, Me, Panel1, "Splitter1", Check1, Check2
    Splitter1.LeftTopPos = Check1.Width 'important: set the start-position of the Splitter
    Splitter1.BorderStyle = bsXPStyl    'bsXPStyl: for cool-looking a Command-button will be
                                        'created using themeing and animation of the button!
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L As Single
    Dim T As Single: T = BtnNext.Top + BtnNext.Height + brdr
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then
        Panel1.Move L, T, W, H
    End If
End Sub

Private Sub Splitter1_OnMove(Sender As Splitter)
    'e.g. for resizing/moving controls outside the splitter-panel
    'not needed here
End Sub

Private Sub BtnNext_Click()
    Form8.Show
End Sub




