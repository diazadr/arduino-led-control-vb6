VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton tombolOn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton tombolOff 
      BackColor       =   &H000000FF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5280
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "KONTROL LED MELALUI ARDUINO DAN VISUAL BASIC"
      BeginProperty Font 
         Name            =   "Helvetica Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   8775
   End
   Begin VB.Image Image2 
      Height          =   960
      Index           =   0
      Left            =   9120
      Picture         =   "Arduino LED Control using VB.frx":0000
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "Arduino LED Control using VB.frx":1456B
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   960
      Index           =   1
      Left            =   7920
      Picture         =   "Arduino LED Control using VB.frx":1EE39
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      FillColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   1320
      Top             =   720
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tombolOn_Click()
    MSComm1.Output = "1"
End Sub

Private Sub tombolOff_Click()
    MSComm1.Output = "0"
End Sub

Private Sub Form_Load()
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
        MSComm1.Output = "0"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
End Sub



