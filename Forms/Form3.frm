VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   LinkTopic       =   "Form3"
   ScaleHeight     =   4575
   ScaleWidth      =   6255
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
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   2520
         ScaleHeight     =   3585
         ScaleWidth      =   3465
         TabIndex        =   3
         Top             =   0
         Width           =   3495
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFC0C0&
            Height          =   1695
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   5
            Text            =   "Form3.frx":0000
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0FFC0&
            Height          =   1695
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   4
            Text            =   "Form3.frx":0006
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   3495
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   2
         Text            =   "Form3.frx":000C
         Top             =   0
         Width           =   2415
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
Attribute VB_Name = "Form3"
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
    
    Set SplitterH = MNew.Splitter(False, Me, Panel1, "SplitterH", Text1, Panel2)
    With SplitterH
        .BorderStyle = bsXPStyl     'bsXPStyle: we borrow the cool-look of a Command-button to use themeing and animation
        .LeftTopPos = Text1.Width  'important: set the start-position of the Splitter
    End With
    
    Set SplitterV = MNew.Splitter(False, Me, Panel2, "SplitterV", Text2, Text3)
    With SplitterV
        .IsHorizontal = False
        .BorderStyle = bsXPStyl     'bsXPStyle: we borrow the cool-look of a Command-button to use themeing and animation
        .LeftTopPos = Text2.Height * Screen.TwipsPerPixelY / 3 'important: set the start-position of the Splitter
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
    Form4.Show
End Sub



