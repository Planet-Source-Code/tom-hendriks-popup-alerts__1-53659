VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public sTitle As String
Public sMessage As String
Public Link As Boolean
Public sUrl As String


Public Sub MSN6()

'Newest MSN style popup

'New Alert
Dim Alert As New frmAlert6

'Set the Title and Message
Alert.lblTitle.Caption = sTitle
Alert.lblAlert.Caption = sMessage

'If link option is used then use "url hand" icon
'and tell the alert where to link to.

If Link = True Then
Alert.lblAlert.MouseIcon = Alert.imgPointer.Picture
Alert.lblURL.Caption = sUrl
Alert.lblLink.Caption = "True"
End If

'Display the alert

Alert.Show





End Sub

Public Sub FlatStyle()

'New Alert by ME
'Doesn't support links

'New Alert
Dim Alert As New frmNewAlert

'Set the Title and Message
Alert.lblTitle.Caption = sTitle
Alert.lblAlert.Caption = sMessage


'Display the alert

Alert.Show





End Sub


Public Sub MSNOld()

'New Alert
Dim Alert As New frmAlertOld

'Set Message (no title used here)
Alert.lblMessage.Caption = sMessage

'If link option is used then use "url hand" icon
'and tell the alert where to link to.
If Link = True Then
Alert.lblMessage.MouseIcon = Alert.imgPointer.Picture
Alert.lblURL.Caption = sUrl
Alert.lblLink.Caption = "True"

End If

'Display the alert
Alert.Show
End Sub
