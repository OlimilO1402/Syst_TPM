VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "TPM active?"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer

Private Sub Command1_Click()
    If Not IsAdmin Then
        If MsgBox("You must run the applicatiOn as administrator!" & vbCrLf & "Check anyway?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Dim tpm As New TrustPM
    If tpm.Handle = 0 Then MsgBox tpm.Error_ToStr
End Sub

Public Function IsAdmin() As Boolean
    IsAdmin = CBool(IsUserAnAdmin)
End Function

