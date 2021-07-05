VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
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
      ScaleWidth      =   5985
      TabIndex        =   1
      Top             =   780
      Width           =   6015
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   1920
         ScaleHeight     =   3585
         ScaleWidth      =   3945
         TabIndex        =   3
         Top             =   0
         Width           =   3975
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0E0FF&
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1515
            ScaleWidth      =   3795
            TabIndex        =   5
            Top             =   0
            Width           =   3855
         End
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Height          =   1815
            Left            =   0
            ScaleHeight     =   1755
            ScaleWidth      =   3795
            TabIndex        =   4
            Top             =   1680
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Height          =   3615
         Left            =   0
         ScaleHeight     =   3555
         ScaleWidth      =   1755
         TabIndex        =   2
         Top             =   0
         Width           =   1815
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
Attribute VB_Name = "Form1"
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
    
    Set SplitterH = MNew.Splitter(False, Me, Panel1, "SplitterH", Picture1, Panel2)
    With SplitterH
        .LeftTopPos = Picture1.Width    'important: set the start-position of the Splitter
        .BorderStyle = bsXPStyl         'bsXPStyl: for cool-looking a Command-button will be
                                        'created using themeing and animation of the button!
    End With
    Set SplitterV = MNew.Splitter(False, Me, Panel2, "SplitterV", Picture2, Picture3)
    With SplitterV
        .LeftTopPos = 1575 'Picture2.Height 'important: set the start-position of the Splitter
        .BorderStyle = bsXPStyl         'bsXPStyl: for cool-looking a Command-button will be
                                        'created using themeing and animation of the button!
        .IsHorizontal = False
    End With
    
    Picture1.Print "Picture1"
    Picture2.Print "Picture2"
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
    Form2.Show
End Sub
