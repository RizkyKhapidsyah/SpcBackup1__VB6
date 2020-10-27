VERSION 5.00
Begin VB.Form SpcMain 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Statistical Process Control Charting Program"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton StartButton 
      Caption         =   "Initialize"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton EnterButton 
      Caption         =   "Enter Reading"
      Height          =   495
      Left            =   8160
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox HighLimitTextBox 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6960
      TabIndex        =   9
      Text            =   " 10.0"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Text            =   " 0.0"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox AverageTextBox 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox LastTextBox 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Range"
      Height          =   375
      Left            =   7680
      TabIndex        =   32
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "% of Tol. used"
      Height          =   375
      Left            =   7680
      TabIndex        =   31
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reading #"
      Height          =   255
      Left            =   7680
      TabIndex        =   30
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please Enter High and Low limits before the start, to set the X-Bar and ""R"" increments"
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   720
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X B a r + R C h a r t"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   615
   End
   Begin VB.Label qlable15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Highest Reading"
      Height          =   255
      Left            =   7680
      TabIndex        =   27
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label qlable14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lowest Reading"
      Height          =   255
      Left            =   7680
      TabIndex        =   26
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label qlable13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label qlable12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label qlable11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label qlable0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label qlable1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label qlable2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label qlable3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label qlable4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label qlable5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label qlable6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label qlable7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label qlable8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label qlable9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label qlable10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   " "
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.Label HighLimitTextBoxLable 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "High Limit"
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LowLimitTextBoxLable 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Low Limit"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label AverageTextBoxLable 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Aver. Reading"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LastTextBoxLable 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Reading"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Text1Lable 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter#Here"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "SpcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public proline1
Public proline2
Public proline3
Public middleman
Public middleman1
Public middleman2
Public middleman3
Public amplifier1
Public amplifier2
Public julxa
Public julxb
Public josxa
Public josxb
Public indexit1
Public indexit2
Public indexit3
Public indexit4
Public recorda
Public pace
Public hisline1y
Public hisline2y
Public hisline3y
Public hisline4y
Public hisline5y
Public hisline6y
Public hisline7y
Public hisline8y
Public hisline9y
Public hisline10y
Public hisline11y
Public verx1
Public verx2
Public overalldif
Public tinyman
Public rightHigh
Public rightLow







Private Sub EnterButton_Click()
middleman2 = middleman1
middleman = Text1.Text
middle1man = Text2.Text
If Text1.Text = "" Then middleman = 0
middleman = middleman - middle1man
'If middleman > 10 Then middleman = 10
'If middleman < 0 Then middleman = 0
tinyman = overalldif * 0.1
middleman1 = middleman / tinyman
If middleman1 < 0 Then middleman1 = 0
If middleman1 > 10 Then middleman1 = 10
Call mainchartset
Call rangechartset

Call averiges
Call hischecker
Call Historgram
Call BottomRightBox
LastTextBox.Text = Text1.Text
Text1.Text = ""
indexit1 = indexit1 + indexit3
indexit2 = indexit2 + indexit4
Label2.Caption = ""
Text1.SetFocus
End Sub

Private Sub BottomRightBox()
If rightLow > Text1.Text Then rightLow = Text1.Text
If rightHigh < Text1.Text Then rightHigh = Text1.Text
qlable14.Caption = "Lowest Read " & rightLow
qlable15.Caption = "Highest Read " & rightHigh
Label4.Caption = (((rightHigh - rightLow) / overalldif) * 100) & "% of Tol. Used"
'range now

If rightRange < middleman3 Then rightRange = middleman3

Label5.Caption = "Last Range " & (rightRange * tinyman)
End Sub
Private Sub mainchartset()

julxa = proline1 - (middleman1 * amplifier1)
julxb = proline1 - (middleman2 * amplifier1)

Circle (indexit1, julxa), 35
Line (indexit1, julxa)-((indexit1 - indexit3), julxb)

End Sub
Private Sub rangechartset()
If middleman1 > middleman2 Then middleman3 = middleman1 - middleman2
If middleman2 > middleman1 Then middleman3 = middleman2 - middleman1
josxb = josxa
If middleman3 < 0 Then middleman3 = 10
If middleman3 > 10 Then middleman3 = 10

josxa = proline2 - (middleman3 * amplifier2)
Circle (indexit2, josxa), 30
Line (indexit2, josxa)-((indexit2 - indexit4), josxb)

End Sub
Private Sub averiges()
recorda = recorda + Text1.Text
pace = pace + 1
Label3.Caption = "Reading #" & pace
If pace > 24 Then EnterButton.Visible = False

AverageTextBox.Text = (recorda / pace)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If pace < 25 Then EnterButton.Value = True
End If
End Sub

Private Sub StartButton_Click()
qxcurrent = 1300
qycurrent = 1000
qleftxcurrent = 1335
qrightxcurrent = 9280
qmovingycurrent = 1040
qcolorindexspace = 10
HowWide = 8000
HowHigh = 5000
proline1 = 5950
proline2 = 8050
amplifier1 = 490
amplifier2 = 180
indexit1 = 1650
indexit2 = 1500
indexit3 = 320
indexit4 = 140
recorda = 0
josxa = 8000
proline3 = 8070
hisline1y = 8050
hisline2y = 8050
hisline3y = 8050
hisline4y = 8050
hisline5y = 8050
hisline6y = 8050
hisline7y = 8050
hisline8y = 8050
hisline9y = 8050
hisline10y = 8050
hisline11y = 8050
'>>>>>>>>>>>>>>>><<<<<<<<<<<<<<
Call verified1
middleman1 = 5
rightHigh = qlable1.Caption
rightLow = qlable9.Caption
Line (qxcurrent, qycurrent)-((qxcurrent + HowWide), qycurrent)
Line (qxcurrent, (qycurrent + HowHigh))-((qxcurrent + HowWide), (qycurrent + HowHigh))
Line (qxcurrent, (qycurrent + HowHigh))-(qxcurrent, qycurrent)
Line ((qxcurrent + HowWide), (qycurrent + HowHigh))-((qxcurrent + HowWide), qycurrent)


SpcMain.ForeColor = &HFF&

Do While qmovingycurrent < 5980
    Line (qleftxcurrent, qmovingycurrent)-(qrightxcurrent, qmovingycurrent)
    qmovingycurrent = qmovingycurrent + qcolorindexspace
    If qmovingycurrent > 1240 Then
        If qmovingycurrent < 2275 Then SpcMain.ForeColor = &HFFFF&
        'Yellow
    End If
    If qmovingycurrent > 2275 Then
        If qmovingycurrent < 4745 Then SpcMain.ForeColor = &HC000&
        'green
    End If
    If qmovingycurrent > 4745 Then
        If qmovingycurrent < 5780 Then SpcMain.ForeColor = &HFFFF&
        'yellow
    End If
    If qmovingycurrent > 5780 Then
        If qmovingycurrent < 5980 Then SpcMain.ForeColor = &HFF&
        'Red
    End If
Loop
qmovingycurrent = qmovingycurrent + 100
qrightxcurrent = 4800
Do While qmovingycurrent < 8100
    Line (qleftxcurrent, qmovingycurrent)-(qrightxcurrent, qmovingycurrent)
    qmovingycurrent = qmovingycurrent + qcolorindexspace
    If qmovingycurrent > 6500 Then
        If qmovingycurrent < 7500 Then SpcMain.ForeColor = &HFFFF&
        'Yellow
    End If
    If qmovingycurrent > 7500 Then
        If qmovingycurrent < 8100 Then SpcMain.ForeColor = &HC000&
        'green
    End If
    
Loop
'>>>>>>>>>>>>>>>>>>>>>>>>
SpcMain.ForeColor = &H0&
qmovingycurrent = qmovingycurrent + 20
qleftxcurrent = qleftxcurrent - 35
qrightxcurrent = qrightxcurrent + 20
qstoppedycurrent = 6040
Line (qleftxcurrent, qmovingycurrent)-(qrightxcurrent, qmovingycurrent)
Line (qleftxcurrent, qstoppedycurrent)-(qrightxcurrent, qstoppedycurrent)
Line (qleftxcurrent, qstoppedycurrent)-(qleftxcurrent, qmovingycurrent)
Line (qrightxcurrent, qstoppedycurrent)-(qrightxcurrent, qmovingycurrent)
Text1.SetFocus
End Sub
Private Sub verified1()

verx1 = Text2.Text 'lowlimit
verx2 = HighLimitTextBox.Text  'highlimit
overalldif = verx2 - verx1
qlable0.Caption = verx1
qlable1.Caption = verx1 + (overalldif * 0.1)
qlable2.Caption = verx1 + (overalldif * 0.2)
qlable3.Caption = verx1 + (overalldif * 0.3)
qlable4.Caption = verx1 + (overalldif * 0.4)
qlable5.Caption = verx1 + (overalldif * 0.5)
qlable6.Caption = verx1 + (overalldif * 0.6)
qlable7.Caption = verx1 + (overalldif * 0.7)
qlable8.Caption = verx1 + (overalldif * 0.8)
qlable9.Caption = verx1 + (overalldif * 0.9)
qlable10.Caption = verx2
qlable11.Caption = overalldif
qlable12.Caption = overalldif / 2
Label2.Caption = "Ready for New Data to be Entered"



End Sub

Private Sub hischecker()
If middleman1 > -1 Then
    If middleman1 < 0.5 Then hisline1y = hisline1y - 200
End If
If middleman1 > 0.4 Then
    If middleman1 < 1.2 Then hisline2y = hisline2y - 200
End If
If middleman1 > 1.1 Then
    If middleman1 < 2.2 Then hisline3y = hisline3y - 200
End If
If middleman1 > 2.1 Then
    If middleman1 < 3.2 Then hisline4y = hisline4y - 200
End If
If middleman1 > 3.1 Then
    If middleman1 < 4.2 Then hisline5y = hisline5y - 200
End If
If middleman1 > 4.1 Then
    If middleman1 < 5.2 Then hisline6y = hisline6y - 200
End If
If middleman1 > 5.1 Then
    If middleman1 < 6.2 Then hisline7y = hisline7y - 200
End If
If middleman1 > 6.1 Then
    If middleman1 < 7.2 Then hisline8y = hisline8y - 200
End If
If middleman1 > 7.1 Then
    If middleman1 < 8.2 Then hisline9y = hisline9y - 200
End If
If middleman1 > 8.1 Then
    If middleman1 < 9.6 Then hisline10y = hisline10y - 200
End If
If middleman1 > 9.5 Then
    If middleman1 < 11.1 Then hisline11y = hisline11y - 200
End If

End Sub
Private Sub Historgram()
glexa = 4870
glexb = 6045
glexc = 7500
glexd = 8120
glexe = 4890
glexf = 7480
glexg = 6065
'>>>>>>>.

'..>>>>>>>>
hline1x = 5100
hline2x = 5300
hline3x = 5500
hline4x = 5700
hline5x = 5900
hline6x = 6100
hline7x = 6300
hline8x = 6500
hline9x = 6700
hline10x = 6900
hline11x = 7100
'.>>>>>>>>>>>>>>
SpcMain.ForeColor = &H808080
Do While glexg < 8100
    Line (glexe, glexg)-(glexf, glexg)
    glexg = glexg + 2
Loop
SpcMain.ForeColor = &H0&
Line (glexa, glexb)-(glexc, glexb)
Line (glexa, glexd)-(glexc, glexd)
Line (glexa, glexb)-(glexa, glexd)
Line (glexc, glexb)-(glexc, glexd)
'>>>>>>>>
Line ((glexa + 50), (glexd - 50))-((glexc - 50), (glexd - 50))
'>>>>>>>>>>.
SpcMain.ForeColor = &HFF&
'red
Do While hline1x < 5250
Line (hline1x, proline3)-(hline1x, hisline1y)
hline1x = hline1x + 2
Loop

SpcMain.ForeColor = &HFFFF&
'yellow
Do While hline2x < 5450
Line (hline2x, proline3)-(hline2x, hisline2y)
hline2x = hline2x + 2
Loop


Do While hline3x < 5650
Line (hline3x, proline3)-(hline3x, hisline3y)
hline3x = hline3x + 2
Loop


SpcMain.ForeColor = &HC000&
'green
Do While hline4x < 5850
Line (hline4x, proline3)-(hline4x, hisline4y)
hline4x = hline4x + 2
Loop


Do While hline5x < 6050
Line (hline5x, proline3)-(hline5x, hisline5y)
hline5x = hline5x + 2
Loop


Do While hline6x < 6250
Line (hline6x, proline3)-(hline6x, hisline6y)
hline6x = hline6x + 2
Loop


Do While hline7x < 6450
Line (hline7x, proline3)-(hline7x, hisline7y)
hline7x = hline7x + 2
Loop


Do While hline8x < 6650
Line (hline8x, proline3)-(hline8x, hisline8y)
hline8x = hline8x + 2
Loop

SpcMain.ForeColor = &HFFFF&
'yellow
Do While hline9x < 6850
Line (hline9x, proline3)-(hline9x, hisline9y)
hline9x = hline9x + 2
Loop


Do While hline10x < 7050
Line (hline10x, proline3)-(hline10x, hisline10y)
hline10x = hline10x + 2
Loop

SpcMain.ForeColor = &HFF&
'red
Do While hline11x < 7250
Line (hline11x, proline3)-(hline11x, hisline11y)
hline11x = hline11x + 2
Loop
SpcMain.ForeColor = &H0&
'black


End Sub
