VERSION 5.00
Begin VB.UserControl ucLM 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   70
End
Attribute VB_Name = "ucLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintNumber As Integer
Private mbytState As Byte
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Sub ChangeState()

    'Toggle between 0=Ok, 1=Wounded and 2=Dead
    mbytState = (mbytState + 1) Mod 3
    DrawChar

End Sub

Private Sub DrawChar()

    Cls
    DrawWidth = 1
    'Draw the triangle
    Line (2, 68)-(68, 68)
    Line (2, 68)-(35, 2)
    Line (35, 2)-(68, 68)
    DrawWidth = 3
    If mbytState = 1 Then 'Wounded
        Line (0, 0)-(70, 70), RGB(255, 0, 0)
    ElseIf mbytState = 2 Then 'Dead
        Line (0, 0)-(70, 70), RGB(255, 0, 0)
        Line (70, 0)-(0, 70), RGB(255, 0, 0)
    End If
    CurrentX = (ScaleWidth / 2) - (TextWidth(CStr(mintNumber)) / 2)
    CurrentY = ScaleHeight - (ScaleHeight / 3) - (TextHeight(CStr(mintNumber)) / 2)
    Print CStr(mintNumber)
    Refresh

End Sub

Public Property Get State() As String

    If mbytState = 0 Then
        State = "Ready to kill"
    ElseIf mbytState = 1 Then
        State = "Wounded"
    Else
        State = "Dead"
    End If

End Property

Public Property Get Number() As Integer

    Number = mintNumber

End Property

Public Property Let Number(ByVal iNewValue As Integer)

    mintNumber = iNewValue
    DrawChar

End Property

Private Sub UserControl_Initialize()

    mintNumber = 0
    mbytState = 0
    DrawChar

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub UserControl_Resize()

    UserControl.Height = 1050
    UserControl.Width = 1050

End Sub
