VERSION 5.00
Begin VB.Form frmEnter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   2400
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   960
   End
   Begin VB.Label lblMove 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "move mouse over this field until the characters appear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   240
      TabIndex        =   39
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Index           =   39
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   1750
      Width           =   300
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   38
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Index           =   37
      Left            =   900
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   1750
      Width           =   775
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Index           =   36
      Left            =   575
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   1750
      Width           =   300
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   35
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   34
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   33
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   32
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   31
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   30
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   29
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   28
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   27
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   26
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   25
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   24
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   23
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   22
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   21
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   20
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   19
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   18
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   17
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   16
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   15
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   14
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   13
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   12
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   11
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************
'
' VIRTUAL PASSWORD
'
' This program is used to enter passwords without using
' key or mouseclick's. By using this concept it is
' impossible to steal passwords by logging keys, mouse
' position (each time the alfabet has other positions),
' by screencapture or hooking textboxes with passwordchars
'
' When the Password Form appears you move the mouse
' across the square to generate a random sequenced
' alfabet. Once the alfabet appears you can select
' characters by pointing them more than one second.
' If you need a pause or are searching for the next
' character, don't move after your last character,
' or keep moving across the square without stopping,
' or move the mouse outside the alfabet. Use CE to
' restart, CANCEL to cancel, OK to validate the code
' and A... to change from upper to lower case.
'
'
' For those who think someone may peek over theire shoulder,
' you could change the alfabet colors in such a way that the
' contrast is very poor. Safer, but harder to read (but then
' again, thats the purpose). In reality you'll notice that
' it's almost impossible for others to follow the movements
' and validation of the chars from a distance.
'
'
'
' D. Rijmenants 2004
'
'***************************************************************

Private intLastChar As Integer
Private strCode As String
Private strChar(35) As String
Private ScramKey() As Byte
Private strSeed As String
Private blnUpperCase As Boolean
Private blnSilentMode As Boolean
Private blnScramFlag As Boolean

Option Explicit

Private Sub ValidateCode()

'***************************************************************
' This is the place to validate your password
' In this demo case we transfer the password to
' the public value strEnteredCode to test it
'***************************************************************

strEnteredCode = strCode

End Sub

Private Sub Form_Activate()

'***************************************************************
'select silentmode (beep/nobeep after each entered char)
blnSilentMode = False
'select if started with uppercase/lowercase
blnUpperCase = True
'***************************************************************

Me.Timer1.Enabled = False
intLastChar = -1
strCode = ""
strSeed = ""
Me.lblMove.Caption = vbCrLf & "move mouse over this field until the characters appear"
blnScramFlag = False
Call ChangeCase
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'hold when moving over form
intLastChar = -1
End Sub

Private Sub lblChar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'detect if other label is moved over
If Index <> intLastChar Then
    Me.Timer1.Interval = Int((2000) * Rnd + 1000)
    Me.Timer1.Enabled = False
    Me.Timer1.Enabled = True
    intLastChar = Index
    End If
End Sub


Private Sub lblMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'collect rnd values for seeding the scramble routine
Static XP As Single
Static YP As Single
Static Skip As Integer
intLastChar = -1
If X = XP And Y = YP Then Exit Sub
XP = X: YP = Y
Skip = Skip + 1
If Skip < 3 Then Skip = Skip + 1: Exit Sub
Skip = 0
strSeed = strSeed & Chr((X Xor Y) And 255)
If Len(strSeed) > 34 Then
    'seed ready
    Me.lblMove.Visible = False
    Call ScrambleChars
    End If
End Sub

Private Sub Timer1_Timer()
'Timer tick when more than 1 second over a label
Me.Timer1.Enabled = False
If intLastChar = -1 Then Exit Sub 'don't tick if moved over form
Select Case intLastChar
Case 36
    'erase (CE)
    strCode = ""
    Call ScrambleChars
Case 37
    'cancel entering
    Call ClearAll
    Me.Hide
Case 38
    'validate code
    Call ValidateCode
    Call ClearAll
    Me.Hide
Case 39
    'change upper/lower case
    blnUpperCase = Not blnUpperCase
    Call ChangeCase
    Call ScrambleChars
Case Else
    'add char
    strCode = strCode & Me.lblChar(intLastChar).Caption
    Call ScrambleChars
End Select
If Not blnSilentMode Then Beep
End Sub

Private Sub ScrambleChars()
' fill the square with random sequenced alfabet an figures
Dim i As Integer
Dim j As Integer
Dim tmp As Integer
Dim aKey As String
Dim k
'setup seed array
ScramKey() = StrConv(strSeed, vbFromUnicode)
'init scram the first time
If blnScramFlag = False Then
    blnScramFlag = True
    For i = 0 To 35
        strChar(i) = i
    Next i
End If
'scramble sequence
For i = 0 To 35
    j = (j + strChar(i) + ScramKey(i Mod Len(strSeed))) Mod 34
    tmp = strChar(i)
    strChar(i) = strChar(j)
    strChar(j) = tmp
Next
'transfer chars to labels
For i = 0 To 35
    Me.lblChar(i).Visible = True
    j = strChar(i)
    If j < 26 Then
        'set letters
        If blnUpperCase = True Then
            Me.lblChar(i) = Chr(Asc("A") + j)
            Else
            Me.lblChar(i) = Chr(Asc("a") + j)
            End If
        Else
        'set figures
        Me.lblChar(i) = Trim(Str(j - 26))
    End If
Next i
End Sub

Private Sub ClearAll()
'erase all chars
Dim i As Integer
For i = 0 To 35
    Me.lblChar(i).Caption = ""
    Me.lblChar(i).Visible = False
Next i
Me.lblMove.Visible = True
strCode = ""
End Sub

Private Sub ChangeCase()
'change lower/upper cases labels
Dim i As Integer
For i = 0 To 35
    If blnUpperCase = True Then
        lblChar(i).Caption = UCase(lblChar(i).Caption)
        Else
        lblChar(i).Caption = LCase(lblChar(i).Caption)
    End If
Next i
If blnUpperCase = True Then
    lblChar(39).Caption = "a.."
    Else
    lblChar(39).Caption = "A.."
    End If
End Sub
