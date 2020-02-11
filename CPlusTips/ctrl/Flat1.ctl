VERSION 5.00
Begin VB.UserControl Flat 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   Palette         =   "Flat1.ctx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   1410
   ToolboxBitmap   =   "Flat1.ctx":0312
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   1185
      X2              =   1185
      Y1              =   285
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   165
      X2              =   165
      Y1              =   300
      Y2              =   1350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   1230
      X2              =   105
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   105
      X2              =   1230
      Y1              =   150
      Y2              =   150
   End
End
Attribute VB_Name = "Flat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DM Flat ActiveX Control For Visual Basic
' Writen and designed by Ben Jones
' Email DreamVb@outlook.com
' Copyright © 2002 Ben Jones
Option Explicit

Private Sub UserControl_Resize()
    Line1(0).X1 = 0
    Line1(0).Y1 = 0
    Line1(0).Y2 = 0
    Line1(1).X1 = 0
    Line1(2).X1 = 0
    Line1(2).X2 = 0
    Line1(2).Y1 = 0
    Line1(3).Y1 = 0
    
    Line1(3).X1 = UserControl.Width - 8
    Line1(3).X2 = UserControl.Width - 8
    Line1(3).Y2 = UserControl.Height
    Line1(2).Y2 = UserControl.Height
    Line1(0).X2 = UserControl.Width
    Line1(1).Y1 = UserControl.Height - 8
    Line1(1).Y2 = UserControl.Height - 8
    Line1(1).X2 = UserControl.Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    BorderColor = Line1(0).BorderColor
    BorderColor = Line1(1).BorderColor
    BorderColor = Line1(2).BorderColor
    BorderColor = Line1(3).BorderColor
End Property

Public Property Let BorderColor(ByVal New_ForeColor As OLE_COLOR)
    Line1(0).BorderColor() = New_ForeColor
    Line1(1).BorderColor() = New_ForeColor
    Line1(2).BorderColor() = New_ForeColor
    Line1(3).BorderColor() = New_ForeColor
    PropertyChanged "BorderColor"
End Property

'Load property values from storage

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Line1(0).BorderColor = PropBag.ReadProperty("BorderColor", &H80000012)
    Line1(1).BorderColor = PropBag.ReadProperty("BorderColor", &H80000012)
    Line1(2).BorderColor = PropBag.ReadProperty("BorderColor", &H80000012)
    Line1(3).BorderColor = PropBag.ReadProperty("BorderColor", &H80000012)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderColor", Line1(0).BorderColor, &H80000012)
    Call PropBag.WriteProperty("BorderColor", Line1(1).BorderColor, &H80000012)
    Call PropBag.WriteProperty("BorderColor", Line1(2).BorderColor, &H80000012)
    Call PropBag.WriteProperty("BorderColor", Line1(3).BorderColor, &H80000012)
End Sub

