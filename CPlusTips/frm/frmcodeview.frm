VERSION 5.00
Begin VB.Form frmcodeview 
   Caption         =   "Codeview"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7350
   Icon            =   "frmcodeview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pButtons 
      Align           =   2  'Align Bottom
      BackColor       =   &H00BCBCBC&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   490
      TabIndex        =   2
      Top             =   3855
      Width           =   7350
      Begin Project1.dmFlatButton cmdCopy 
         Height          =   420
         Left            =   210
         TabIndex        =   3
         Top             =   30
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   741
         BackGround      =   12369084
         BackgroundHover =   13619151
         TextColorHover  =   0
         BorderHover     =   10724259
         Text            =   "Copy"
         MouseDownBackground=   14211288
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.dmFlatButton cmdClose 
         Height          =   420
         Left            =   6465
         TabIndex        =   4
         Top             =   30
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   741
         BackGround      =   12369084
         BackgroundHover =   13619151
         TextColorHover  =   0
         BorderHover     =   10724259
         Text            =   "Close"
         MouseDownBackground=   14211288
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Project1.dStatusbar dStatusbar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   4350
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   582
      BarStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripStyle       =   1
   End
   Begin Project1.dEditor dEditor1 
      Height          =   3660
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   6456
      Locked          =   -1  'True
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmcodeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Call Unload(frmcodeview)
End Sub

Private Sub cmdCopy_Click()
    Call dEditor1.SelectAll
    Call dEditor1.Copy
End Sub

Private Sub Form_Load()
    'Place sourcecode into editor.
    dEditor1.Text = SourceCode
End Sub

Private Sub Form_Resize()
On Error Resume Next
    dEditor1.Width = frmcodeview.ScaleWidth - dEditor1.Left - 8
    dEditor1.Height = (frmcodeview.ScaleHeight - pButtons.ScaleHeight - dEditor1.Top) - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmcodeview = Nothing
End Sub

Private Sub pButtons_Resize()
On Error Resume Next
    cmdClose.Left = (pButtons.ScaleWidth - cmdClose.Width) - 8
End Sub
