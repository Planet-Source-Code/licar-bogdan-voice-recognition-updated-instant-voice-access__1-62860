VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecord 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Recorder"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3720
      Top             =   120
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   1440
      Top             =   1440
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   2640
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Save"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   240
      Top             =   1440
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
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Record your default sound, save it in a safe place, then change the DefaultPath (in cmdLoad_Click) with this one."
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
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
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tm As Date, Recording As Boolean

Private Sub ResetLabel()
lblTime.Caption = CDate(Time - Time)
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

Private Sub lblSave_Click()
CD1.Filter = "Wav file (*.wav)|*.wav"
CD1.Flags = &H2 Or &H400
CD1.ShowSave

SaveWave CD1.FileName
End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSave.ForeColor = &H80FF&
End Sub

Private Sub lblStop_Click()
StopWave
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
lblSave.ForeColor = &HFFFF&
End Sub

