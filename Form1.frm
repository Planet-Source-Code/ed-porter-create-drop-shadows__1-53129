VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtWidth 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5040
      TabIndex        =   0
      Text            =   "10"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drop Shadow"
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1080
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width "
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim flgS As Boolean


Private Sub Command1_Click()
'Create Drop Shadow

LockWindowUpdate Me.hWnd 'prevent flicker

'Create drop shadow effect - object, shadow width (in pixels)
DrawDropShadow pic1, CLng(txtWidth.Text)

'Refresh .....
pic1.Refresh
LockWindowUpdate 0

Command1.Enabled = False
txtWidth.Enabled = False
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
'Clear drop shadow - remove from pic1

With pic1
    .Width = .Width - CLng(txtWidth.Text)
    .Height = .Height - CLng(txtWidth.Text)
    .Refresh
End With

Command1.Enabled = True
Command2.Enabled = False

With txtWidth
    .Enabled = True
    .SetFocus
End With
    
End Sub


Private Sub txtWidth_GotFocus()

With txtWidth
    .SelStart = 0
    .SelLength = 10
End With

End Sub


Private Sub txtWidth_LostFocus()

'Prevent distortion errors due to false input
With txtWidth
    If .Text <> "" Then
        If CLng(.Text) < 4 Then .Text = Trim(Str(4))
        If CLng(.Text) > 14 Then .Text = Trim(Str(14))
    End If
End With

End Sub
