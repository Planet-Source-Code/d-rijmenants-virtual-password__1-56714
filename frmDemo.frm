VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Demo by D. Rijmenants"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Login Demo"
      Height          =   1935
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton cmdEnter 
         Caption         =   "&Enter Password"
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Click the button below to enter the password and check if it's valid."
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Password Demo"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Password"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Click the button below to change the default 'TEST' password into your own."
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"frmDemo.frx":0000
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
' for this demo we start with using "TEST" as secret password
strSecretCode = "TEST"
End Sub

Private Sub cmdEnter_Click()
frmEnter.Show (vbModal)
If strEnteredCode = "" Then Exit Sub
If strEnteredCode = strSecretCode Then
    MsgBox "You have entered a valid password.", vbInformation
    Else
    MsgBox "Sorry, no access: wrong Password!", vbCritical
    End If

End Sub

Private Sub cmdNew_Click()
Dim strTmp As String
Dim ret As Integer
frmEnter.Show (vbModal)
If strEnteredCode = "" Then Exit Sub
strTmp = strEnteredCode
'confirm the new code
ret = MsgBox("Please confirm the password.", vbOKCancel + vbExclamation)
If ret = vbCancel Then Exit Sub
frmEnter.Show (vbModal)
If strEnteredCode = strTmp Then
    'change code into new code
    strSecretCode = strEnteredCode
    MsgBox "Your password has been changed!", vbInformation
    Else
    MsgBox "The password and the confirmation do not match! Please try again.", vbCritical
    End If
End Sub


