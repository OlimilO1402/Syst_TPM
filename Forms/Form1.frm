VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Text            =   "Form1.frx":1782
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton BtnIsTPMActive 
      Caption         =   "Is Trusted-Platform-Module active?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function AppRunsAsAdmin Lib "shell32" Alias "#680" () As Integer

Private Sub BtnIsTPMActive_Click()
    If Not IsAdmin Then
        If MsgBox("You must run the application as administrator!" & vbCrLf & "Check anyway?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Dim tpm As New TrustPM
    If Not tpm.CheckContext1 Then MsgBox "Check TPM-Context Version 1.2: " & vbCrLf & tpm.Error_ToStr
    If Not tpm.CheckContext2 Then MsgBox "Check TPM-Context Version 2.0: " & vbCrLf & tpm.Error_ToStr
    
    If Not tpm.CheckDeviceInfo Then MsgBox "Check Device Info: " & vbCrLf & tpm.Error_ToStr
    Text1.Text = "Check Device Info: " & vbCrLf & tpm.DeviceInfo_ToStr
End Sub

Public Function IsAdmin() As Boolean
    IsAdmin = CBool(AppRunsAsAdmin)
End Function

