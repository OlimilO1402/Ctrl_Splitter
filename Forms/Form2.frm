VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnNext 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   2400
         ScaleHeight     =   3585
         ScaleWidth      =   3465
         TabIndex        =   3
         Top             =   0
         Width           =   3495
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Label3"
            Height          =   1695
            Left            =   0
            TabIndex        =   5
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Label2"
            Height          =   1695
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Label1"
         Height          =   3495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents SplitterH As Splitter
Attribute SplitterH.VB_VarHelpID = -1
Private WithEvents SplitterV As Splitter
Attribute SplitterV.VB_VarHelpID = -1

Private Sub Form_Load()

    Set SplitterH = MNew.Splitter(False, Me, Panel1, "Splitter1", Label1, Panel2)
    With SplitterH
        .BorderStyle = bsXPStyl     'bsXPStyle: we borrow the cool-look of a Command-button to use themeing and animation
        .LeftTopPos = Label1.Width  'important: set the start-position of the Splitter
    End With

    Set SplitterV = MNew.Splitter(False, Me, Panel2, "SplitterV", Label2, Label3)
    With SplitterV
        .IsHorizontal = False
        .BorderStyle = bsXPStyl     'bsXPStyle: we borrow the cool-look of a Command-button to use themeing and animation
        .LeftTopPos = Label2.Height * Screen.TwipsPerPixelY 'important: set the start-position of the Splitter
    End With

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

Private Sub SplitterH_OnMove(Sender As Splitter)
    'e.g. for resizing/moving controls outside the splitter-panel
    'not needed here
End Sub

Private Sub SplitterV_OnMove(Sender As Splitter)
    'e.g. for resizing/moving controls outside the splitter-panel
    'not needed here
End Sub

Private Sub BtnNext_Click()
    Form3.Show
End Sub


