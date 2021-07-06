VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15255
   LinkTopic       =   "Form10"
   ScaleHeight     =   9255
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   0
      ScaleHeight     =   8385
      ScaleWidth      =   15105
      TabIndex        =   1
      Top             =   720
      Width           =   15135
      Begin VB.ListBox List1 
         Height          =   8250
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4215
      End
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   8295
         Left            =   4320
         ScaleHeight     =   8265
         ScaleWidth      =   10545
         TabIndex        =   2
         Top             =   0
         Width           =   10575
         Begin VB.TextBox Text1 
            Height          =   2055
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   7
            Text            =   "Form10.frx":0000
            Top             =   6120
            Width           =   9735
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   1695
            Left            =   0
            TabIndex        =   6
            Top             =   4200
            Width           =   9735
            ExtentX         =   17171
            ExtentY         =   2990
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   1935
            Left            =   0
            TabIndex        =   5
            Top             =   2160
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   3413
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"Form10.frx":0006
         End
         Begin VB.PictureBox Picture1 
            Height          =   1935
            Left            =   0
            ScaleHeight     =   1875
            ScaleWidth      =   10035
            TabIndex        =   4
            Top             =   0
            Width           =   10095
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
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1

Private Sub Form_Load()
    Set Splitter1 = New Splitter
    Splitter1.New_ False, Me, Panel1, "Splitter1", List1, Panel2
    With Splitter1
        .LeftTopPos = List1.Width
        .BorderStyle = bsXPStyl
    End With

End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    T = Panel1.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then
        Panel1.Move l, T, W, H
    End If
End Sub

