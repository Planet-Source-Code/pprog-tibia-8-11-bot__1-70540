VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tibia Uane Bot - by Hdo01"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   2430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Always on top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   2175
      Begin VB.Timer Timer3 
         Interval        =   200
         Left            =   1680
         Top             =   600
      End
      Begin VB.Timer Timer2 
         Interval        =   3000
         Left            =   1680
         Top             =   120
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   34
         Text            =   "150"
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "300"
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "MANA Below"
         Height          =   210
         Left            =   420
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "HP below"
         Height          =   255
         Left            =   420
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Joystic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   2175
      Begin VB.Line Line29 
         BorderColor     =   &H00FFFFFF&
         X1              =   360
         X2              =   1800
         Y1              =   1680
         Y2              =   480
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00FFFFFF&
         X1              =   360
         X2              =   1800
         Y1              =   480
         Y2              =   1680
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00FFFFFF&
         X1              =   1080
         X2              =   1080
         Y1              =   240
         Y2              =   1920
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   1920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   960
         Width           =   255
      End
      Begin VB.Line Line2 
         X1              =   1080
         X2              =   840
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line3 
         X1              =   1080
         X2              =   1320
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line5 
         X1              =   840
         X2              =   1080
         Y1              =   1200
         Y2              =   1920
      End
      Begin VB.Line Line6 
         X1              =   1320
         X2              =   1080
         Y1              =   1200
         Y2              =   1920
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   735
         Left            =   720
         Shape           =   3  'Circle
         Top             =   720
         Width           =   735
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   1920
         Y1              =   840
         Y2              =   1080
      End
      Begin VB.Line Line7 
         X1              =   1200
         X2              =   1920
         Y1              =   1320
         Y2              =   1080
      End
      Begin VB.Line Line8 
         X1              =   960
         X2              =   240
         Y1              =   840
         Y2              =   1080
      End
      Begin VB.Line Line9 
         X1              =   960
         X2              =   240
         Y1              =   1320
         Y2              =   1080
      End
      Begin VB.Line Line10 
         X1              =   1080
         X2              =   1800
         Y1              =   840
         Y2              =   480
      End
      Begin VB.Line Line11 
         X1              =   1800
         X2              =   1320
         Y1              =   480
         Y2              =   1080
      End
      Begin VB.Line Line12 
         X1              =   1080
         X2              =   360
         Y1              =   840
         Y2              =   480
      End
      Begin VB.Line Line13 
         X1              =   840
         X2              =   360
         Y1              =   1080
         Y2              =   480
      End
      Begin VB.Line Line14 
         X1              =   840
         X2              =   360
         Y1              =   1080
         Y2              =   1680
      End
      Begin VB.Line Line15 
         X1              =   360
         X2              =   1080
         Y1              =   1680
         Y2              =   1320
      End
      Begin VB.Line Line16 
         X1              =   1080
         X2              =   1800
         Y1              =   1320
         Y2              =   1680
      End
      Begin VB.Line Line17 
         X1              =   1320
         X2              =   1800
         Y1              =   1080
         Y2              =   1680
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   19
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00FFFFFF&
         X1              =   840
         X2              =   960
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00FFFFFF&
         X1              =   1320
         X2              =   1200
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00FFFFFF&
         X1              =   840
         X2              =   960
         Y1              =   960
         Y2              =   600
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00FFFFFF&
         X1              =   1320
         X2              =   1200
         Y1              =   960
         Y2              =   600
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00FFFFFF&
         X1              =   960
         X2              =   600
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00FFFFFF&
         X1              =   600
         X2              =   960
         Y1              =   1200
         Y2              =   1320
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00FFFFFF&
         X1              =   1560
         X2              =   1200
         Y1              =   960
         Y2              =   840
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00FFFFFF&
         X1              =   1560
         X2              =   1200
         Y1              =   1200
         Y2              =   1320
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   600
   End
   Begin VB.Line Line31 
      X1              =   2400
      X2              =   2400
      Y1              =   0
      Y2              =   6240
   End
   Begin VB.Label Label31 
      Caption         =   "[iNFO]"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1800
      TabIndex        =   37
      Top             =   5880
      Width           =   495
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2400
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   35
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "---"
      Height          =   210
      Left            =   2160
      TabIndex        =   14
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "---"
      Height          =   210
      Left            =   2160
      TabIndex        =   13
      Top             =   720
      Width           =   180
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "---"
      Height          =   210
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   180
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -240
      X2              =   2400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   11
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Soul:"
      Height          =   210
      Left            =   450
      TabIndex        =   10
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Level:"
      Height          =   210
      Left            =   370
      TabIndex        =   9
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Exp:"
      Height          =   210
      Left            =   495
      TabIndex        =   8
      Top             =   360
      Width           =   315
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cap:"
      Height          =   210
      Left            =   465
      TabIndex        =   7
      Top             =   1320
      Width           =   330
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "MANA:"
      Height          =   210
      Left            =   285
      TabIndex        =   6
      Top             =   960
      Width           =   510
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "HP:"
      Height          =   210
      Left            =   570
      TabIndex        =   5
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' http://www.hdo01.yoyo.pl
'
' hdo01@o2.pl
'
' gg: 937303
'
'
' by Hdo01
'


Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub SendPacket Lib "packet.dll" (ByVal ProcessID As Long, ByRef Packet() As Byte, Optional ByVal Encrypt As Byte = True, Optional ByVal SafeArray As Byte = True)

Private Const PLAYER_HP_MAX = &H613B68
Private Const PLAYER_MANA_MAX = &H613B4C
Private Const PLAYER_LVLP = &H613B58
Private Const PLAYER_SOUL = &H613B48
Private Const PLAYER_LEVEL = &H613B60
Private Const PLAYER_EXP = &H613B64
Private Const PLAYER_HP = &H613B6C
Private Const PLAYER_MANA = &H613B50
Private Const PLAYER_CAP = &H613B40
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Function SetWinPos(lHWnd As Long, bPos As Boolean) As Boolean
Dim lWinPos As Long
Dim l As Long

Select Case bPos
 Case False
  WinPos = HWND_NOTOPMOST
 Case True
  lWinPos = HWND_TOPMOST
End Select
If SetWindowPos(lHWnd, lWinPos, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE) Then
 SetWinPos = True
End If
End Function

Function CzytajPamiec(Adres As Long)
On Error Resume Next
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
Dim valbuffer As Long
Dim test As Long
hwnd = FindWindow("tibiaclient", vbNullString)
If (hwnd = 0) Then Exit Function
GetWindowThreadProcessId hwnd, pid
pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
test = ReadProcessMemory(pHandle, Adres, valbuffer, 4, 0&)
CzytajPamiec = valbuffer
CloseHandle pHandle
End Function


Private Sub Check3_Click()
If Check3.Value = 1 Then SetWinPos Me.hwnd, True
End Sub

Private Sub Form_Load()
MsgBox "Change 'packet.dll_' on 'packet.dll'!"
MsgBox "Tibia Uane Bot" & vbNewLine & "It is only samle to show how to read/write tibia packets... Enjoy it!" & vbNewLine & "Tibia char: Uane of Aurea" & vbNewLine & "mail: hdo01@o2.pl" & vbNewLine & "gg: 937303" & vbNewLine & vbNewLine & " by Hdo01"
MsgBox "Click on Joystick to move your player in Tibia"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call StopAlarm
End Sub

Private Function StartAlarm()
alarm = "1"
mciSendString "open " & "alarm.wav" & " type MPEGVideo alias CURRFILE", 0&, 0, 0
mciSendString "play CURRFILE repeat", 0&, 0, 0

End Function
Private Function StopAlarm()
alarm = "0"
mciSendString "close CURRFILE", 0&, 0, 0
End Function


Private Sub Label16_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H65
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.BorderStyle = 1
End Sub

Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.BorderStyle = 0
End Sub

Private Sub Label17_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H6A
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.BorderStyle = 1
End Sub

Private Sub Label17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.BorderStyle = 0
End Sub

Private Sub Label18_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H66
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.BorderStyle = 1
End Sub

Private Sub Label18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.BorderStyle = 0
End Sub

Private Sub Label19_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H6B
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.BorderStyle = 1
End Sub

Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.BorderStyle = 0
End Sub

Private Sub Label20_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H67
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.BorderStyle = 1
End Sub

Private Sub Label20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.BorderStyle = 0
End Sub

Private Sub Label21_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H6C
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label21.BorderStyle = 1
End Sub

Private Sub Label21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label21.BorderStyle = 0
End Sub

Private Sub Label22_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H68
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.BorderStyle = 1
End Sub

Private Sub Label22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.BorderStyle = 0
End Sub

Private Sub Label23_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H6D
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.BorderStyle = 1
End Sub

Private Sub Label23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.BorderStyle = 0
End Sub

Private Sub Label24_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H6F
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label24.BorderStyle = 1
End Sub

Private Sub Label24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label24.BorderStyle = 0
End Sub

Private Sub Label25_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H71
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label25.BorderStyle = 1

End Sub

Private Sub Label25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label25.BorderStyle = 0
End Sub

Private Sub Label26_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H72
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.BorderStyle = 1

End Sub

Private Sub Label26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.BorderStyle = 0
End Sub

Private Sub Label27_Click()
Dim LowerVal As Long
Dim UpperVal As Long
Dim ProcessID As Long
Dim PacketBuffer(2) As Byte
GetWindowThreadProcessId FindWindow("tibiaclient", vbNullString), ProcessID
PacketBuffer(0) = &H10
PacketBuffer(1) = &H0
PacketBuffer(2) = &H70
SendPacket ProcessID, PacketBuffer
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label27.BorderStyle = 1

End Sub

Private Sub Label27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label27.BorderStyle = 0
End Sub

Private Sub Label31_Click()
MsgBox "Copyright © 2008 by Hdo01"
End Sub

Private Sub Text1_Change()
If Not IsNumeric(Text1.Text) Then Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
Call sprawdz
End Sub
Private Function sprawdz()
Label5.Caption = CzytajPamiec(PLAYER_LEVEL)
Label1.Caption = CzytajPamiec(PLAYER_EXP)
Label2.Caption = CzytajPamiec(PLAYER_HP)
Label3.Caption = CzytajPamiec(PLAYER_MANA)
Label4.Caption = CzytajPamiec(PLAYER_CAP)
Label12.Caption = CzytajPamiec(PLAYER_SOUL)
Label13.Caption = (100 - CzytajPamiec(PLAYER_LVLP)) & "% do " & (CzytajPamiec(PLAYER_LEVEL) + 1) & "lvlu"
Label14.Caption = "(max. " & CzytajPamiec(PLAYER_HP_MAX) & ")"
Label15.Caption = "(max. " & CzytajPamiec(PLAYER_MANA_MAX) & ")"
Label30.Caption = "GPS: (x:" & CzytajPamiec(&H61E9C8) & ",y:" & CzytajPamiec(&H61E9C4) & ",z:" & CzytajPamiec(&H61E9C0) & ")"
End Function

Private Sub Timer2_Timer()
' Need work on this function

'If Check1.Value = 1 Then If Label2.Caption < Text1.Text Then  Call StartAlarm: Exit Sub
'Call StopAlarm
End Sub

Private Sub Timer3_Timer()
' Need work on this function

'If Check2.Value = 1 Then If Label3.Caption < Text2.Text Then Call StartAlarm: Exit Sub
'Call StopAlarm
End Sub

Private Sub Timer4_Timer()

End Sub
