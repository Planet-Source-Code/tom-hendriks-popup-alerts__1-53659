VERSION 5.00
Begin VB.Form frmAlert6 
   BorderStyle     =   0  'None
   ClientHeight    =   1755
   ClientLeft      =   12495
   ClientTop       =   11520
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   1755
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   720
   End
   Begin VB.Image imgPointer 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMain.frx":10206
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label lblURL 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblLink 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image imgX 
      Height          =   240
      Left            =   2520
      Picture         =   "frmMain.frx":10510
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgXNormal 
      Height          =   240
      Left            =   960
      Picture         =   "frmMain.frx":10852
      Top             =   2520
      Width           =   225
   End
   Begin VB.Image imgXOver 
      Height          =   225
      Left            =   720
      Picture         =   "frmMain.frx":10B94
      Top             =   2520
      Width           =   210
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   2295
   End
   Begin VB.Label lblAlert 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your text here"
      Height          =   1095
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmAlert6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgX.Picture = imgXNormal.Picture
End Sub





Private Sub imgX_Click()
Unload Me
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgX.Picture = imgXOver.Picture
End Sub

Private Sub lblAlert_Click()
If lblLink.Caption = "True" Then
Dim lSuccess As Long

lSuccess = ShellExecute(Me.hwnd, "Open", lblURL.Caption, 0, 0, 10)

End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgX.Picture = imgXNormal.Picture
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 30
If Me.Top = 9330 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Me.Top = Me.Top + 30
If Me.Top = 11520 Then
Timer3.Enabled = False
Unload Me
End If
End Sub
