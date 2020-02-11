VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.dmFlatButton cmdOK 
      Height          =   360
      Left            =   4245
      TabIndex        =   3
      Top             =   1620
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   635
      TextColorHover  =   8388608
      Text            =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblHyper 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit the code tips website"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   735
      TabIndex        =   5
      Top             =   1605
      Width           =   2220
   End
   Begin VB.Label lblDeveloper 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by DreamVB"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   4
      Top             =   1140
      Width           =   1725
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.2"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4500
      TabIndex        =   2
      Top             =   375
      Width           =   420
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free collection of C/C++ tips and tricks"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   885
      TabIndex        =   1
      Top             =   825
      Width           =   4200
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM C++ Code Reader"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1125
      TabIndex        =   0
      Top             =   315
      Width           =   3225
   End
   Begin VB.Image ImgLogo 
      Height          =   720
      Left            =   270
      Picture         =   "frmabout.frx":0000
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Call Unload(frmabout)
End Sub

Private Sub Form_Load()
    Set frmabout.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub

Private Sub lblHyper_Click()
'    Call OpenUrl("https://dreamvb.000webhostapp.com/CodeTips/")
End Sub
