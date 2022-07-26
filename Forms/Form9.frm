VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   LinkTopic       =   "Form9"
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
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   2040
         ScaleHeight     =   3585
         ScaleWidth      =   3945
         TabIndex        =   2
         Top             =   0
         Width           =   3975
         Begin VB.PictureBox Picture1 
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1515
            ScaleWidth      =   3555
            TabIndex        =   5
            Top             =   0
            Width           =   3615
         End
         Begin VB.FileListBox File1 
            Height          =   1845
            Left            =   0
            TabIndex        =   4
            Top             =   1680
            Width           =   3735
         End
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
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents SplitterH As Splitter
Attribute SplitterH.VB_VarHelpID = -1
Private WithEvents SplitterV As Splitter
Attribute SplitterV.VB_VarHelpID = -1

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
    Set SplitterH = MNew.Splitter(False, Me, Panel1, "SplitterH", Dir1, Panel2)
    With SplitterH
        .LeftTopPos = Dir1.Width  'important: set the start-position of the Splitter
        .BorderStyle = bsXPStyl    'bsXPStyl:  we borrow the cool-look of a Command-button to use themeing and animation
    End With
    
    Set SplitterV = MNew.Splitter(False, Me, Panel2, "SplitterV", Picture1, File1)
    With SplitterV
        .LeftTopPos = Picture1.Height 'important: set the start-position of the Splitter
        .BorderStyle = bsXPStyl    'bsXPStyl:  we borrow the cool-look of a Command-button to use themeing and animation
        .IsHorizontal = False
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
    '
End Sub

Private Sub BtnNext_Click()
    Form10.Show
End Sub


