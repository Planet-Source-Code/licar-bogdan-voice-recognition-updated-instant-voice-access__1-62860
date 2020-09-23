VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voice recognition"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10620
   Begin VB.OptionButton optIsolate 
      Caption         =   "Auto-Isolate"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   3600
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton optManually 
      Caption         =   "Choose wave length manually"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Go"
      ForeColor       =   &H8000000E&
      Height          =   2775
      Left            =   9000
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton cmdRec 
         Caption         =   "Record Default Sound"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdSpeak 
         Caption         =   "Get In"
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H8000000A&
         Caption         =   "Load Wave"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdMatch 
         Caption         =   "Start Matching Process"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   1560
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Waves"
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9015
      Begin VB.PictureBox picWav1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3075
         ScaleWidth      =   4275
         TabIndex        =   7
         Top             =   240
         Width           =   4335
         Begin VB.Label Label1 
            BackColor       =   &H80000007&
            Caption         =   "Original sound"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblLength1 
            BackColor       =   &H80000007&
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3360
            TabIndex        =   9
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox picWav2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000012&
         Height          =   3135
         Left            =   4560
         ScaleHeight     =   3075
         ScaleWidth      =   4275
         TabIndex        =   6
         Top             =   240
         Width           =   4335
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            Caption         =   "Verifying sound"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblPlay 
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PLAY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblLength2 
            BackColor       =   &H80000007&
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3360
            TabIndex        =   10
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000B&
      Caption         =   "Choose wave length manually"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   2280
      Width           =   2655
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
      EndProperty
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load a wavyeah"
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Help"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   10200
      TabIndex        =   16
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                       *** Voice Recognition Program ***
'
'Description: It basically reads byte by byte two wave files and compares them like
'             two sequences of integers, leaving a small range of error and giving
'             the final result of the matching process. The two sounds must have
'             pretty much the same tone of voice, in order to match. Otherwise it will
'             give a negative result. I think there are still some bugs and it needs
'             plenty of optimizations. NO EXTERNAL ENGINES (dlls, ocxs etc).
'             You can quickly bypass the program by saying the password, or by loading
'             a pre-recorded wave file.
'
'Purpose:     Password substitution method. I wanted to copy spy movies. It actually
'             works, even if obviously it isn't unfailable.
'
'Note:        I didn't put a command button to load the first wave file too, because
'             this program is projected as a password substituter, so such a routine
'             would make no sense.
'
'Author:      Licar Bogdan (c).

Option Explicit

Dim values1() As Double, values2() As Double
Dim Path1 As String, Y1() As Double, Y2() As Double
Dim DefaultPath As String

'Try it with the two waves that I've included (in honour of Robert Redford). Sorry for my horrible voice.
Private Sub DrawWave(picWav As PictureBox, Path As String, values() As Double, _
Y() As Double, Optional Col = vbYellow, Optional AllWave As Boolean = True)

Dim i As Long, Buff As Long, Xrate As Double, Yrate As Double
Dim LastX As Double, LastY As Double, Min As Double, Max As Double
Dim CurrX As Double, j As Long, imin As Long, imax As Long

'Not using other variables just for lack of will:
'values(0)=the startpoint of the important part
'values(1)=the endpoint        "    "
'values(3)=number of low peaks
'values(4)=number of high peaks

On Error Resume Next
        
        'Reads it and registers all values in an array
        i = 44 'Set i To 44, since the wave sample is begin at Byte 44.
        Open Path For Random As #1
        Do
            Get #1, i, Buff
            i = i + 1: ReDim Preserve values(i)
            If Buff < Min Then Min = Buff
            If Buff > Max Then Max = Buff
            values(i) = Buff
        Loop Until EOF(1)
        Close #1
        
    With picWav
        'Operations for the drawing of the wave file
        Xrate = (.ScaleWidth / i)
        Yrate = (Max - Min) / (.ScaleTop)
        LastY = 0
        Max = (Max / Yrate)
        Min = (Min / Yrate)

        'Change the values of the *.wav in cartesian coordinates
        For j = 44 To UBound(values)
            ReDim Preserve Y(j - 43)
            Y(j - 43) = (values(j) / Yrate)
            
            'Count peaks above/below a certain value
            If Y(j - 43) > (.ScaleTop / 3.5) Then values(4) = values(4) + 1
            If Y(j - 43) < (-.ScaleTop / 3.5) Then values(3) = values(3) + 1
        Next j

        'Loops for isolating the important part of the wave file
        If (AllWave = False And optIsolate.Value = True) Or (AllWave = False And Path = DefaultPath) Then
            For j = 1 To UBound(Y) - 1          'The beginning of the "talking" part
                If (Abs(Y(j) - Y(j + 1))) > 100 Then values(0) = j: Exit For
            Next j
            For j = UBound(Y) To 1 Step -1      'The end of the "talking"
                If (Abs(Y(j) - Y(j - 1))) > 100 Then values(1) = j: Exit For
            Next j
            imin = values(0): imax = values(1)
        ElseIf AllWave = False And optManually.Value = True And Path <> DefaultPath Then
            imin = values(0): imax = values(1)  'Values chosen by the user
        ElseIf AllWave = True Then
            imin = 1: imax = UBound(Y)
        End If
        
        'Activate these lines to have the "important" part fit on the picture
        'Xrate = (.ScaleWidth / (imax - imin))
        'Yrate = (Max - Min) / (.ScaleTop)
    
        'The drawing part
        picWav.Line (0, 0)-(.ScaleWidth, 0), Col
        For i = imin To imax
            CurrX = CurrX + Xrate
            picWav.Line (LastX, LastY)-(CurrX, Y(i)), Col
            LastX = CurrX
            LastY = .CurrentY
            If CurrX > .ScaleWidth Then Exit For
            DoEvents
        Next i
    
    End With
End Sub

Private Sub cmdSpeak_Click()
frmGetIn.Show vbNormalFocus
Path1 = TmpPath
cmdLoad_Click
End Sub

Private Sub Form_Load()
SetTheScale picWav1, 0, 1000, -500, 500
SetTheScale picWav2, 0, 1000, -500, 500

DefaultPath = App.Path & "\voice.wav"
End Sub

Private Sub cmdLoad_Click()
On Error Resume Next

Screen.MousePointer = vbHourglass
picWav1.Cls: picWav2.Cls: picWav1.Left = 120: picWav2.Left = 4560
lblLength1.Caption = "": lblLength2.Caption = ""
StBar.Panels(2).Text = "Drawing wave files..."
lblPlay.Visible = False: Label1.Visible = False: Label2.Visible = False
    
    If Path1 = "" Then
    With CommonDialog1
        .CancelError = True
        .Filter = "Wave files (*.wav)|*.wav"
        .ShowOpen
        Path1 = .FileName
    End With
    StBar.Panels(1).Text = Path1
    End If
    
    'Resets the arrays
    DeleteArray values1: DeleteArray values2
    DeleteArray Y1: DeleteArray Y2

    'Draws the default recorded wave file and the one used for the matching process
    DrawWave picWav1, DefaultPath, values1, Y1
    DrawWave picWav2, Path1, values2, Y2, vbRed
    lblLength1.Caption = Format(UBound(values1) / 691, "#.#0") & " sec"
    lblLength2.Caption = Format(UBound(values2) / 691, "#.#0") & " sec"

    lblPlay.Visible = True: Label1.Visible = True: Label2.Visible = True

    StBar.Panels(2).Text = "To verify, start matching process."
    cmdMatch.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMatch_Click()
Dim i As Double, NumTries As Integer

lblLength1.Caption = ""
StBar.Panels(2).Text = "Checking equivalence..."
Label1.Visible = False

    'Moves the waves towards each other
    For i = 1 To 50000
        If picWav2.Left > picWav1.Left Then
            If i Mod 1000 = 0 Then picWav2.Left = picWav2.Left - 60: picWav1.Left = picWav1.Left + 60
        End If
        DoEvents
    Next i
    
    'Redraw the waves, isolating the important part
    picWav1.Cls
    DrawWave picWav1, DefaultPath, values1, Y1, vbYellow, False
    DrawWave picWav1, Path1, values2, Y2, vbRed, False
    
    StBar.Panels(2).Text = PointToPoint_Comparison(Y1, Y2) & "% match"
    
    'Activate these lines to allow a limitate number of tries
    'If NumTries <= 3 Then
    If AccessGranted = True Then
        
        MsgBox "                   ACCESS GRANTED." & vbCrLf & vbCrLf _
            & "Heuristic Analysis: Waves Match." & vbCrLf _
            & "Point-to-Point Percentage Result: " & PointToPoint_Comparison(Y1, Y2) & "% match." & vbCrLf _
            & "Wave Statistic-Sectioning: " & Statistic_Comparison(Y1, Y2) & "% match.", vbInformation
    Else
    
        MsgBox "                   ACCESS DENIED.            " & vbCrLf & vbCrLf _
            & "Heuristic Analysys: Waves Do Not Match." & vbCrLf _
            & "Point-to-Point Percentage Result: " & PointToPoint_Comparison(Y1, Y2) & "% match." & vbCrLf _
            & "Wave Statistic-Sectioning: " & Statistic_Comparison(Y1, Y2) & "% match.", vbCritical
    'NumTries = NumTries + 1
    End If
    'Else
        'MsgBox "You have no more tries left.", vbCritical
    'End If
    
    'Kill TmpPath
    cmdSpeak.SetFocus
    Path1 = ""
End Sub

Function AccessGranted() As Boolean
    'Result based on the number of high/low peaks and on statistics of the whole "important part".
    'For now it has a failure probability = 12%. These error ranges may be varied.
    'If you find a better way to verify waves, please let me know.
    On Error Resume Next
    If Abs(UBound(values1) - UBound(values2)) > 300 Then AccessGranted = False: Exit Function
    If (Abs(values1(3) - values2(3)) < 10) And (Abs(values1(4) - values2(4)) < 10) And _
    (Abs(ArithmeticMean(Y1, values1(0), values1(1)) - ArithmeticMean(Y2, values2(0), values2(1))) < 6) And _
    (Abs(StandardDeviation(Y1, values1(0), values1(1)) - StandardDeviation(Y2, values2(0), values2(1))) < 10) Then
        AccessGranted = True
    Else
        AccessGranted = False
    End If
End Function

Private Sub cmdRec_Click()
frmRecord.Show vbNormalFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ResetWave
Kill App.Path & "\Tmp.wav"
End Sub

Private Sub Label3_Click()
MsgBox "Record your default sound (any length; better maximum 5 secs), then save the file" & vbCrLf _
        & "in a safe place (i.e. \system32) and with a safe name (i.e. sys.wav)." & vbCrLf _
        & "Finally, change the DefaultPath (in Form_Load) with the path you've recorded. " & vbCrLf _
        & "You can either record instantly a wave or to load a wave, in order to compare it." & vbCrLf _
        & "It can be easily integrated in your programs, consisting in a password substituter.", vbInformation

End Sub

Private Sub lblPlay_Click()
sndPlaySound Path1, 1
End Sub

Private Sub picWav2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If the "Choose the wave length manually" is checked
Dim i As Integer
If optManually.Value = True Then
    If values2(0) = 0 Then
        picWav2.Line (X, picWav2.ScaleTop)-(X, -picWav2.ScaleTop), vbBlue
        values2(0) = Int(X / (picWav2.ScaleWidth / UBound(values2))) - 1
        For i = 0 To X Step 15
            picWav2.Line (i, 500)-(i, -500), vbGreen
        Next i
    ElseIf values2(1) = 0 Then
        picWav2.Line (X, picWav2.ScaleTop)-(X, -picWav2.ScaleTop), vbBlue
        values2(1) = Int(X / (picWav2.ScaleWidth / UBound(values2))) - 1
        For i = X To 1000 Step 15
            picWav2.Line (i, 500)-(i, -500), vbGreen
        Next i
    End If
End If
End Sub

Function PointToPoint_Comparison(vals1() As Double, vals2() As Double) As Double
Dim j As Long, i As Long, Same As Long, Same2 As Long, ErrRange As Integer
On Error Resume Next
ErrRange = 10
'Compares each value of the default sound with the value of the sound used
'for the matching process. It leaves a small range of error. With a greater number
'than 10, there are more chances the sounds to match, but they could also be different,
'giving a high percentage of matching though.
'This is a not so highly efficient method.

For j = 1 To (values1(1) - values1(0))
    If j = values2(1) Then Exit For

    For i = -2 To 2
        If (Abs(vals1(values1(0) + j) - vals2(values2(0) + j + i))) < ErrRange Then Same = Same + 1: Exit For
    Next i

Next j

PointToPoint_Comparison = Format((Same * 100) / (values1(1) - values1(0)), "#.##")
End Function

Function Statistic_Comparison(vals1() As Double, vals2() As Double) As Double
Dim i As Long, j As Long, Same2 As Long, ErrRange As Integer, v1() As Double, v2() As Double, ArrSize As Integer
ArrSize = 20
ReDim v1(ArrSize): ReDim v2(ArrSize)
On Error Resume Next

'This could be done also by dividing the wave in totally separate parts and analyze them
For j = 1 To (values1(1) - values1(0))
    If (j + values2(0)) > values2(1) Then Exit For
    
    For i = 1 To ArrSize
        v1(i) = vals1(values1(0) + j + i): v2(i) = vals2(values2(0) + j + i)
    Next i
    
    If (Abs(ArithmeticMean(v1, LBound(v1), UBound(v1)) - _
        ArithmeticMean(v2, LBound(v2), UBound(v2)))) < 10 And _
        (Abs(StandardDeviation(v1, LBound(v1), UBound(v1)) - _
        StandardDeviation(v2, LBound(v2), UBound(v2)))) < 20 Then Same2 = Same2 + 1
    
Next j
'I think it's better than the point-to-point technique.
Statistic_Comparison = Format((Same2 * 100) / Round((values1(1) - values1(0))), "#.##")
End Function

Function SetTheScale(pic As Object, XMin As Single, XMax As Single, YMin As Single, YMax As Single)
pic.ScaleLeft = XMin
pic.ScaleTop = YMax
pic.ScaleWidth = XMax - XMin
pic.ScaleHeight = -(YMax - YMin)
End Function

Function ArithmeticMean(vals() As Double, ByVal StartPoint As Long, ByVal EndPoint As Long) As Single
Dim i As Long, result As Single
On Error Resume Next
    For i = StartPoint To EndPoint
        result = result + vals(i)
    Next i
ArithmeticMean = result / (EndPoint - StartPoint)
End Function

Function GeometricMean(vals() As Double, ByVal StartPoint As Long, ByVal EndPoint As Long)
Dim i As Long, result As Single
On Error Resume Next
    result = 1
    For i = StartPoint + 10 To EndPoint
        If vals(i) = 0 Then GoTo Nxt
        result = result * vals(i)
Nxt:
    Next i
GeometricMean = Abs(result ^ (1 / (EndPoint - StartPoint)))
End Function

Function StandardDeviation(vals() As Double, ByVal StartPoint As Long, ByVal EndPoint As Long) As Double
Dim i As Long, Am As Single, result As Double
On Error Resume Next
Am = ArithmeticMean(vals(), StartPoint, EndPoint)
    For i = StartPoint To EndPoint
        result = result + ((vals(i) - Am) ^ 2)
    Next i
StandardDeviation = Format(Sqr(result / (EndPoint - StartPoint)), "#.##")
End Function

Private Sub DeleteArray(arr() As Double)
Dim i As Double
On Error GoTo ErrorHandler
    For i = 0 To UBound(arr)
        arr(i) = 0
    Next i
ErrorHandler:
End Sub
'School sucks.
