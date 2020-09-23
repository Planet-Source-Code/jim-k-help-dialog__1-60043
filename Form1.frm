VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterative Help dialog box"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSlide 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   210
      TabIndex        =   8
      ToolTipText     =   "Click here for help..."
      Top             =   360
      Width           =   210
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   2280
      ScaleHeight     =   930
      ScaleWidth      =   3060
      TabIndex        =   1
      Top             =   360
      Width           =   3060
      Begin VB.PictureBox picLeft 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   930
         Left            =   0
         Picture         =   "Form1.frx":0370
         ScaleHeight     =   930
         ScaleWidth      =   150
         TabIndex        =   6
         Top             =   0
         Width           =   150
      End
      Begin VB.PictureBox picMain 
         BorderStyle     =   0  'None
         Height          =   930
         Left            =   -1680
         Picture         =   "Form1.frx":0720
         ScaleHeight     =   930
         ScaleWidth      =   3585
         TabIndex        =   2
         Top             =   0
         Width           =   3585
         Begin VB.Label lblOk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            ForeColor       =   &H00AC9887&
            Height          =   195
            Left            =   3165
            TabIndex        =   5
            ToolTipText     =   "Ok.. Hide help"
            Top             =   570
            Width           =   225
         End
         Begin VB.Label lblHelpTxt 
            BackStyle       =   0  'Transparent
            Caption         =   "This GUI Help dialog tells you this chkBox enables/disables the textbox"
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   2760
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   1680
   End
   Begin VB.CheckBox Check1 
      Caption         =   "What's this checkbox for"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' By: Jim K on April, 05

Private Sub Check1_Click()
    Text1.Enabled = Check1.Value
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    picMain.Left = -picMain.Width
    picBack.Width = 0
    picSlide.Left = picBack.Left + picBack.Width
    picBack.Width = picSlide.Width
    picBack.Height = picSlide.Height
    picLeft.Visible = False
    Text1.Enabled = Check1.Value
End Sub

Private Sub lblOk_Click()
    Form_Resize
    picSlide.Visible = True
End Sub

Private Sub picSlide_Click()
    If picSlide.Left > 3585 Then
        tmrSlide.Enabled = False
        picLeft.Visible = False
    Else
        tmrSlide.Enabled = True
    End If
End Sub

Private Sub tmrSlide_Timer()
    picBack.Height = 930
    picMain.Visible = True
    picLeft.Visible = True
    lblOk.Visible = False
    picMain.Left = picMain.Left + 60
    picBack.Width = picMain.Left + picMain.Width
    picSlide.Left = picBack.Left + picBack.Width - 10
    picLeft.Left = 0
    If picMain.Left > 0 Then
        tmrSlide.Enabled = False
        picSlide.Visible = False
        picLeft.Left = 0
        picMain.Left = 0
        lblOk.Visible = True
    End If
End Sub
