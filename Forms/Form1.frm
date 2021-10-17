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
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton BtnIsTPMActive 
      Caption         =   "Check Trusted-Platform-Module"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function AppRunsAsAdmin Lib "shell32" Alias "#680" () As Integer

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & _
           App.FileDescription, vbOKOnly Or vbInformation
End Sub

Private Sub BtnIsTPMActive_Click()
    If Not IsAdmin Then
        If MsgBox("You must run the application as administrator!" & vbCrLf & "Check anyway?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    Dim tpm As New TrustPM
    Dim s As String: s = "Check TPM-Context Version "
    MsgBox s & "1.2: " & vbCrLf & IIf(tpm.CheckContext1, "Found!", tpm.Error_ToStr)
    MsgBox s & "2.0: " & vbCrLf & IIf(tpm.CheckContext2, "Found!", tpm.Error_ToStr)
    
    If Not tpm.CheckDeviceInfo Then MsgBox "Check Device Info: " & vbCrLf & tpm.Error_ToStr
    Text1.Text = "Check Device Info: " & vbCrLf & tpm.DeviceInfo_ToStr
End Sub

Public Function IsAdmin() As Boolean
    IsAdmin = CBool(AppRunsAsAdmin)
End Function

Private Sub Form_Resize()
    Dim L As Single, T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub
