VERSION 5.00
Begin VB.Form frmGetIn 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Say the password"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2640
      Top             =   120
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   2280
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblStop 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   1200
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   120
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmGetIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tm As Date, Recording As Boolean

Private Sub ResetLabel()
lblTime.Caption = CDate(Time - Time)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Recording = False And KeyAscii = 13 Then
    lblRec_Click
ElseIf Recording = True And KeyAscii = 13 Then
    lblStop_Click
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRec.ForeColor = &HFF&
lblStop.ForeColor = &HFFFF&
lblOK.ForeColor = &HFFFF&
End Sub

Private Sub lblOK_Click()
Unload Me
End Sub

Private Sub lblOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = &H80FF&
End Sub

Private Sub lblRec_Click()
RecordWave
ResetLabel
Tm = Time: tmrTime.Enabled = True
Recording = True: Me.SetFocus
End Sub

Private Sub lblRec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRec.ForeColor = &H80&
End Sub

Private Sub lblStop_Click()
StopWave
TmpPath = App.Path & "\Tmp.wav"
SaveWave TmpPath
ResetLabel: tmrTime.Enabled = False
Recording = False: Me.SetFocus
End Sub

Private Sub Form_Load()
ResetLabel
End Sub

Private Sub lblStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStop.ForeColor = &H80FF&
End Sub

Private Sub tmrTime_Timer()
lblTime.Caption = CDate(Time - Tm)
End Sub
