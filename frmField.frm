VERSION 5.00
Begin VB.Form frmField 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D&D battle field"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmField.frx":0000
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1012
   Begin VB.ComboBox cboColor 
      Height          =   315
      ItemData        =   "frmField.frx":030A
      Left            =   13680
      List            =   "frmField.frx":0326
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   90
      Width           =   1365
   End
   Begin VB.CommandButton cmdRemoveLines 
      Caption         =   "Remove Lines"
      Height          =   375
      Left            =   10980
      TabIndex        =   10
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemoveMonsters 
      Caption         =   "Remove Monsters"
      Height          =   375
      Left            =   9090
      TabIndex        =   9
      Top             =   45
      Width           =   1455
   End
   Begin DnDBattleField.ucMM ucMM 
      Height          =   525
      Index           =   0
      Left            =   10665
      TabIndex        =   8
      Top             =   630
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin DnDBattleField.ucLM ucLM 
      Height          =   1050
      Index           =   0
      Left            =   11430
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1852
   End
   Begin DnDBattleField.ucLPC ucLPC 
      Height          =   1050
      Index           =   0
      Left            =   9180
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1852
   End
   Begin DnDBattleField.ucMPC ucMPC 
      Height          =   525
      Index           =   0
      Left            =   8415
      TabIndex        =   5
      Top             =   630
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.CommandButton cmdRemovePCs 
      Caption         =   "Remove PCs"
      Height          =   375
      Left            =   3285
      TabIndex        =   4
      Top             =   45
      Width           =   1410
   End
   Begin VB.CommandButton cmdAddLM 
      Caption         =   "Add Large Monster"
      Height          =   375
      Left            =   7155
      TabIndex        =   3
      Top             =   45
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddMM 
      Caption         =   "Add Medium Monster"
      Height          =   375
      Left            =   5130
      TabIndex        =   2
      Top             =   45
      Width           =   1905
   End
   Begin VB.CommandButton cmdAddLPC 
      Caption         =   "Add Large PC"
      Height          =   375
      Left            =   1710
      TabIndex        =   1
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddMPC 
      Caption         =   "Add Medium PC"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1500
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Line color"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12825
      TabIndex        =   12
      Top             =   135
      Width           =   825
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintNbPC As Integer
Private mintNbMPC As Integer
Private mintNbLPC As Integer
Private mintNbM As Integer
Private mintNbMM As Integer
Private mintNbLM As Integer
Private mintMouseDownMPCIndex As Integer
Private mintMouseDownLPCIndex As Integer
Private mintMouseDownLMIndex As Integer
Private mintMouseDownMMIndex As Integer
Private mintLastX As Integer
Private mintLastY As Integer
Private mlngLineColor As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Sub cboColor_Click()

    Select Case cboColor.Text
    Case "Green"
        mlngLineColor = RGB(0, 170, 0)
    Case "Black"
        mlngLineColor = RGB(0, 0, 0)
    Case "Blue"
        mlngLineColor = RGB(0, 0, 220)
    Case "Red"
        mlngLineColor = RGB(220, 0, 0)
    Case "Yellow"
        mlngLineColor = RGB(255, 255, 0)
    Case "Orange"
        mlngLineColor = RGB(255, 150, 0)
    Case "Violet"
        mlngLineColor = RGB(150, 0, 150)
    Case "Brown"
        mlngLineColor = RGB(200, 100, 0)
    End Select

End Sub

Private Sub cmdAddLM_Click()

    mintNbLM = mintNbLM + 1
    mintNbM = mintNbM + 1
    Load ucLM(mintNbLM)
    ucLM(mintNbLM).Number = mintNbM
    ucLM(mintNbLM).Left = 8
    ucLM(mintNbLM).Top = 32
    ucLM(mintNbLM).Visible = True

End Sub

Private Sub cmdAddLPC_Click()

    mintNbLPC = mintNbLPC + 1
    mintNbPC = mintNbPC + 1
    Load ucLPC(mintNbLPC)
    ucLPC(mintNbLPC).Number = mintNbPC
    ucLPC(mintNbLPC).Left = 8
    ucLPC(mintNbLPC).Top = 32
    ucLPC(mintNbLPC).Visible = True

End Sub

Private Sub cmdAddMM_Click()

    mintNbMM = mintNbMM + 1
    mintNbM = mintNbM + 1
    Load ucMM(mintNbMM)
    ucMM(mintNbMM).Number = mintNbM
    ucMM(mintNbMM).Left = 8
    ucMM(mintNbMM).Top = 32
    ucMM(mintNbMM).Visible = True

End Sub

Private Sub cmdAddMPC_Click()

    mintNbMPC = mintNbMPC + 1
    mintNbPC = mintNbPC + 1
    Load ucMPC(mintNbMPC)
    ucMPC(mintNbMPC).Number = mintNbPC
    ucMPC(mintNbMPC).Left = 8
    ucMPC(mintNbMPC).Top = 32
    ucMPC(mintNbMPC).Visible = True

End Sub

Private Sub cmdRemoveLines_Click()

    Me.Cls

End Sub

Private Sub cmdRemovePCs_Click()

    RemovePCs

End Sub

Private Sub cmdRemoveMonsters_Click()

    RemoveMonsters

End Sub

Private Sub DrawGrid(ByVal intEspacement As Integer)

Dim i As Integer

    For i = 6 To 1006 Step intEspacement
        Line (i, 30)-(i, 710), RGB(0, 0, 0)
    Next i
    For i = 30 To 710 Step intEspacement
        Line (6, i)-(1006, i), RGB(0, 0, 0)
    Next i

End Sub

Private Sub Form_Load()

    Me.AutoRedraw = True
    DrawGrid 40
    Me.Picture = Me.Image
    DrawWidth = 5
    mintNbM = 0
    mintNbLM = 0
    mintNbMM = 0
    mintMouseDownLMIndex = 0
    mintMouseDownMMIndex = 0
    mintNbLPC = 0
    mintNbMPC = 0
    mintNbPC = 0
    mintMouseDownMPCIndex = 0
    mintMouseDownLPCIndex = 0
    Me.Top = 0
    Me.Left = 0
    cboColor.Text = "Green"
    mlngLineColor = RGB(0, 170, 0)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Me.MousePointer = 99
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        Line (mintLastX, mintLastY)-(x, y), mlngLineColor
    ElseIf Button = vbRightButton Then
        PaintPicture Me.Picture, x - 13, y - 13, 26, 26, x - 13, y - 13, 26, 26, vbSrcCopy
        Me.Refresh
    End If
    mintLastX = x
    mintLastY = y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Me.MousePointer = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    RemoveMonsters
    RemovePCs

End Sub

Private Sub RemoveMonsters()

Dim i As Integer

    For i = 1 To mintNbMM
        Unload ucMM(i)
    Next i
    For i = 1 To mintNbLM
        Unload ucLM(i)
    Next i
    mintNbM = 0
    mintNbLM = 0
    mintNbMM = 0
    mintMouseDownLMIndex = 0
    mintMouseDownMMIndex = 0

End Sub

Private Sub RemovePCs()

Dim i As Integer

    For i = 1 To mintNbMPC
        Unload ucMPC(i)
    Next i
    For i = 1 To mintNbLPC
        Unload ucLPC(i)
    Next i
    mintNbLPC = 0
    mintNbMPC = 0
    mintNbPC = 0
    mintMouseDownMPCIndex = 0
    mintMouseDownLPCIndex = 0

End Sub

Private Sub ucLPC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ucLPC(Index).ChangeState
    Else
        mintMouseDownLPCIndex = Index
    End If

End Sub

Private Sub ucLPC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim pa As POINTAPI

    If mintMouseDownLPCIndex = Index And Index <> 0 Then
        GetCursorPos pa
        ucLPC(Index).Left = pa.x - (Me.Left / Screen.TwipsPerPixelX) - (ucLPC(Index).Width / 2)
        ucLPC(Index).Top = pa.y - (Me.Top / Screen.TwipsPerPixelY) - (ucLPC(Index).Height / 2) - 25
    End If

End Sub

Private Sub ucLPC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If mintMouseDownLPCIndex = Index And Index <> 0 Then
        mintMouseDownLPCIndex = 0
    End If

End Sub

Private Sub ucLM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ucLM(Index).ChangeState
    Else
        mintMouseDownLMIndex = Index
    End If

End Sub

Private Sub ucLM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim pa As POINTAPI

    If mintMouseDownLMIndex = Index And Index <> 0 Then
        GetCursorPos pa
        ucLM(Index).Left = pa.x - (Me.Left / Screen.TwipsPerPixelX) - (ucLM(Index).Width / 2)
        ucLM(Index).Top = pa.y - (Me.Top / Screen.TwipsPerPixelY) - (ucLM(Index).Height / 2) - 25
    End If

End Sub

Private Sub ucLM_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If mintMouseDownLMIndex = Index And Index <> 0 Then
        mintMouseDownLMIndex = 0
    End If

End Sub

Private Sub ucMM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ucMM(Index).ChangerState
    Else
        mintMouseDownMMIndex = Index
    End If

End Sub

Private Sub ucMM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim pa As POINTAPI

    If mintMouseDownMMIndex = Index And Index <> 0 Then
        GetCursorPos pa
        ucMM(Index).Left = pa.x - (Me.Left / Screen.TwipsPerPixelX) - (ucMM(Index).Width / 2)
        ucMM(Index).Top = pa.y - (Me.Top / Screen.TwipsPerPixelY) - (ucMM(Index).Height / 2) - 25
    End If

End Sub

Private Sub ucMM_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If mintMouseDownMMIndex = Index And Index <> 0 Then
        mintMouseDownMMIndex = 0
    End If

End Sub

Private Sub ucMPC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ucMPC(Index).ChangerState
    Else
        mintMouseDownMPCIndex = Index
    End If

End Sub

Private Sub ucMPC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim pa As POINTAPI

    If mintMouseDownMPCIndex = Index And Index <> 0 Then
        GetCursorPos pa
        ucMPC(Index).Left = pa.x - (Me.Left / Screen.TwipsPerPixelX) - (ucMPC(Index).Width / 2)
        ucMPC(Index).Top = pa.y - (Me.Top / Screen.TwipsPerPixelY) - (ucMPC(Index).Height / 2) - 25
    End If

End Sub

Private Sub ucMPC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If mintMouseDownMPCIndex = Index And Index <> 0 Then
        mintMouseDownMPCIndex = 0
    End If

End Sub
