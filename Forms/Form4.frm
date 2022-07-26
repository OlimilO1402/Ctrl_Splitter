VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   LinkTopic       =   "Form4"
   ScaleHeight     =   4575
   ScaleWidth      =   6255
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
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   2040
         ScaleHeight     =   3465
         ScaleWidth      =   3945
         TabIndex        =   3
         Top             =   120
         Width           =   3975
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Frame2"
            Height          =   1575
            Left            =   0
            TabIndex        =   5
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H008080FF&
            Caption         =   "Frame2"
            Height          =   1695
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Frame1"
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1
Private WithEvents Splitter2 As Splitter
Attribute Splitter2.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Set Splitter1 = MNew.Splitter(False, Me, Panel1, "Splitter1", Frame1, Panel2)
    With Splitter1
        .BorderStyle = bsXPStyl    'bsXPStyl: for cool-looking a Command-button will be
        .LeftTopPos = Frame1.Width 'important: set the start-position of the Splitter
    End With                       'created using themeing and animation of the button!
        
    Set Splitter2 = MNew.Splitter(False, Me, Panel2, "Splitter2", Frame2, Frame3)
    With Splitter2
        .IsHorizontal = False
        .BorderStyle = bsXPStyl    'bsXPStyl: for cool-looking a Command-button will be
        '.LeftTopPos = Frame1.Width 'important: set the start-position of the Splitter
        .LeftTopPos = Frame2.Height * Screen.TwipsPerPixelY
    End With                       'created using themeing and animation of the button!
    
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
    Form5.Show
End Sub

