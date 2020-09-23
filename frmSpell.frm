VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmSpell 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Spell Checker"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   FillColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScrollSpeed 
      Height          =   135
      Left            =   4560
      Max             =   200
      Min             =   50
      TabIndex        =   36
      Top             =   1080
      Value           =   100
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   5640
      TabIndex        =   35
      Top             =   2880
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4680
      TabIndex        =   24
      Text            =   "Voice Type"
      Top             =   720
      Width           =   1935
   End
   Begin VB.OptionButton optTakeTest 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Read to Take Test ? Then Check me"
      Height          =   255
      Left            =   8040
      TabIndex        =   20
      Top             =   600
      Width           =   3495
   End
   Begin VB.OptionButton optReview 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Review Before Exam"
      Height          =   255
      Left            =   8040
      TabIndex        =   19
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H0000C0C0&
      Caption         =   "Close Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtPad 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "frmSpell.frx":0000
      Top             =   2880
      Width           =   4815
   End
   Begin VB.OptionButton optLevel1 
      BackColor       =   &H00FF8080&
      Caption         =   "Level5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin VB.OptionButton optLevel1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Level4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton optLevel1 
      BackColor       =   &H00FF8080&
      Caption         =   "Level3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.OptionButton optLevel1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Level2"
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
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton optLevel1 
      BackColor       =   &H00FF8080&
      Caption         =   "Level1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FF8080&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox cmbIncorrect 
      Height          =   315
      Left            =   5160
      TabIndex        =   7
      Text            =   "In Correct Word"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox cmbcorrect 
      Height          =   315
      Left            =   8880
      TabIndex        =   6
      Text            =   "Correct Word"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton optRndRect 
      BackColor       =   &H00C0C000&
      Caption         =   "Rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Want to see Rectangular Application ?"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.OptionButton optEllip 
      BackColor       =   &H00FFFF00&
      Caption         =   "Elliptical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      ToolTipText     =   "Want to see Elliptical Application ? Check Borders to change further."
      Top             =   6000
      Width           =   1215
   End
   Begin VB.OptionButton optNormal 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      ToolTipText     =   "Back To Noraml ?"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdspell 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Spell Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoadText 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Load Text"
      Height          =   495
      Left            =   4200
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdTestWord 
      BackColor       =   &H0080FF80&
      Caption         =   "Click For Test Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblResultInCorrect 
      BackColor       =   &H008080FF&
      Caption         =   "Incorrect Words"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   40
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblResultCorrect 
      BackColor       =   &H0080FFFF&
      Caption         =   "Correct Words"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   39
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblResult1 
      BackColor       =   &H0080FF80&
      Caption         =   "No of Attempt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   38
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblWindow 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Optional : View Different Window Style"
      Height          =   375
      Left            =   4920
      TabIndex        =   37
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label lblStep41 
      BackColor       =   &H00C0FFFF&
      Caption         =   "An input box will show up. Type the test word in the Box, then click ""OK"". To continue, repeat Step 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   34
      Top             =   5400
      Width           =   3855
   End
   Begin VB.Label lblStep4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Step  4."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   33
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label lblStep31 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmSpell.frx":0006
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   32
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label lblstep21 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Select Test Level. Click on Option button."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   31
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblstep11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "You may choose voice type from combo box. To test voice click on Click For Test  Word. Set the speed with Volume control Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   30
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Step   3.  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   29
      Top             =   3600
      Width           =   870
   End
   Begin VB.Label lblStep2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Step 2."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   28
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lblStep1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Step 1. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   27
      Top             =   360
      Width           =   690
   End
   Begin VB.Label lblWhatTodo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Instruction : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblVoiceSpeed 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5160
      TabIndex        =   25
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblInstructor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Instructor"
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblCurentWord 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Chosen Word"
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblUsedWord 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Word Already Used"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   1080
      Width           =   2055
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   10080
      Top             =   960
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label lblIncorrectscore 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Incorrect Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblCorrectScore 
      BackColor       =   &H00FFFF00&
      Caption         =   "Correct Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblCorrect 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Correct Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Search, WordBuffer, DefaultBuffer, TestBuffer  ' Declare variables to collect word in cmdSpell
Dim searchComma, WordBufferComma
' Genie speaks
Dim strspeakGenie As String
'Load text file
Dim strLevel As String
'Use Microsof word 8.0
Private WithEvents RS1 As ADODB.Recordset
Attribute RS1.VB_VarHelpID = -1
Private WithEvents RS2 As ADODB.Recordset
Attribute RS2.VB_VarHelpID = -1
'
Private curAgent As IAgentCtlCharacter
'Get information about operating system
Dim osInfo As clsSysInfo
'Class loading Genie
Dim osLoadacs As clsLoadAcs
'Playing with Shape
'Setting windows' handle
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'if users already tried this word
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'string to Hold the search word in a combo box
Private Const CB_FINDSTRING = &H14C
Dim blnUsedWord As Boolean
'Use Text to speech to say Test Word
Dim voiceText As TextToSpeech
'Variables to show result in List box
Dim intListCorrect As Integer
Dim intListIncorrect As Integer
'
Dim intOptNum As Integer

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdLoadText_Click()
On Error Resume Next
txtPad.Text = ""
Dim X As String
Dim n As Variant
Dim Filename As String
optLevel1_Click (strLevel)
'Filename = App.Path & "\wordbook1.txt"
Filename = App.Path & strLevel
Open Filename For Input As #1
'LOF for the length of the file
'x will hold all the characters
n = LOF(1)
X = Input(n, #1)

Close #1
txtPad.Text = X
txtPad.SetFocus
cmdspell.Enabled = True
End Sub
Private Sub cmdPrinter_Click()
cdlPrinter.Flags = cdlPDPrintSetup
cdlPrinter.PrinterDefault = True
cdlPrinter.ShowPrinter
End Sub

Private Sub cmdRefresh_Click()
If MsgBox("It will Erase all data. Do you want to Refresh.", vbYesNo) = vbYes Then
Unload Me
frmSpell.Show
List1.Refresh
End If
End Sub

Private Sub cmdspell_Click()
'On Error Resume Next
   Dim intLength As Integer
   Dim intComma As Integer
   Dim strWord As String
      ' Get search string from user.
   Search = InputBox("Enter text to be found:")
   If Search = "" Then Exit Sub
   'search if word alreay tried
   Load_genie
   Search = LCase(Search)
   Mid(Search, 1, 1) = UCase(Mid(Search, 1, 1))
   Word_Change ' Checks if the word already exists in combobox.
   'This protects the score board from duplcate words treated as incorrect word.
   'After writting correct word, the correct will disappear form the text pad
   If blnUsedWord = True Then ' Checks if word exists in combobox
   Call curAgent.Speak(" Hey Hey Hay !, didn't you tried this word already ?")
   Call curAgent.Hide 'Genie will hide after
   blnUsedWord = False
   Exit Sub
   End If
   'Add a comma to the string to look for complete word "May,"
   'rather than looking for any letter "M" or "Ma"
   Search = Search & ","
   'Call AlreadyEnterned
   WordBuffer = InStr(txtPad.Text, Search)   ' Find string in text.
   If WordBuffer Then   ' If found,
      'txtPad.SetFocus
      txtPad.SelStart = WordBuffer - 1
    ' set selection start and
      txtPad.SelLength = Len(Search)   ' set selection length.
    'Let us remove comma before showing the word, it looks ugly in combobox
        intLength = Len(Search)
        intComma = InStr(Search, ",")
        strspeakGenie = Left(Search, intComma - 1)
        lblCorrect.Caption = strspeakGenie
    'Genie speaks the typed word
       Call curAgent.Speak("Good Job, You Spelled word " & strspeakGenie & "  correctly. Keep it up")
    'Add the correct word with a stop "." sign after the word.
    'That helps of eliminate typing same word over agian.
    'See procedure Word_Change
       cmbcorrect.AddItem strspeakGenie & "."
       cmbcorrect.Text = cmbcorrect.List(0)
       txtPad.SelText = ""
    'Post correct Word
       CorrectWord
    'Empty seltext
    'LoadTempRS
      strspeakGenie = ""
   Else
      Call curAgent.Speak("Oops, You know what I did not ask this word. Sorry, wrong answer. Good Luck")    ' Notify user.
      cmbIncorrect.AddItem Search
      cmbIncorrect.Text = cmbIncorrect.List(0)
      InCorrectWord
      End If
cmdLoadText.BackColor = vbRed
cmdLoadText.Caption = "X"
cmdLoadText.Enabled = False
cmdRefresh.Enabled = True
cmdTestWord.SetFocus
End Sub
Private Sub Load_genie()
On Error Resume Next
Dim strchar1 As String
Set osInfo = New clsSysInfo
Set osLoadacs = New clsLoadAcs
With osInfo
.OSSys_Type
frmSpell.Caption = .strOS
End With
osLoadacs.Load_genie
Call Agent1.Characters.Load("Genie", strchar)
Set curAgent = Agent1.Characters("Genie")
Call curAgent.Show
'bGenieshow = True
End Sub
Private Sub cmdTestWord_Click()
On Error Resume Next
'Genie_testWord
'Load_genie
Dim iComma As Integer
Dim iLength As Integer
Dim strTestWord As String
If InStr(txtPad.Text, ",") > 0 Then
iLength = Len(txtPad.Text)
iComma = InStr(txtPad.Text, ",")
strTestWord = Left(txtPad.Text, iComma - 1)
lblCurentWord.Caption = strTestWord
End If
'Call curAgent.Speak(strTestWord)
'Set voiceText = New TextToSpeech
voiceText.Select (Combo1.ListIndex + 1)
voiceText.Speed = HScrollSpeed.Value
voiceText.Speak ("spell the word ")
voiceText.Speak strTestWord
lblCurentWord.Visible = False
cmdspell.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
cmdRefresh.Enabled = False
cmdspell.Enabled = False
txtPad.Visible = False
Load_genie
'
Dim intVoiceType As Integer
Dim intCountEngine As Integer
'Early binding and initializing Text to Speech
Set voiceText = New TextToSpeech
'Get info about all the types
intVoiceType = voiceText.CountEngines
 'Use a loop to load in Combo, save space
For intCountEngine = 1 To intVoiceType
Combo1.AddItem voiceText.ModeName(intCountEngine), intCountEngine - 1
Combo1.ListIndex = 5
Next
'Start with a voice type
voiceText.Select (Combo1.ListIndex + 5)
'Show initial set up of speed of speech
lblVoiceSpeed.Caption = HScrollSpeed.Value
End Sub

Private Sub lblInstructor_Click()
On Error Resume Next
lblCurentWord.Visible = True
End Sub

Private Sub optLevel1_Click(Index As Integer)

txtPad.Text = ""
Select Case Index
Case 0
strLevel = "\wordbook1.txt"
intOptNum = 1
Case 1
strLevel = "\wordbook2.txt"
intOptNum = 2
Case 2
strLevel = "\wordbook3.txt"
intOptNum = 3
Case 3
strLevel = "\wordbook4.txt"
intOptNum = 4
Case 4
strLevel = "\wordbook5.txt"
intOptNum = 5
End Select
cmdLoadText_Click
cmdLoadText.Visible = False
List1.AddItem "Level  " & intOptNum
End Sub
Private Sub optNormal_Click()
SetWindowRgn hwnd, CreateRectRgn(0, 0, 800, 500), True
End Sub
Private Sub optEllip_Click()
SetWindowRgn hwnd, CreateEllipticRgn(0, 0, 800, 560), True
End Sub

Private Sub optReview_Click()
On Error Resume Next
txtPad.Visible = True
List1.Visible = False
End Sub

Private Sub optRndRect_Click()
SetWindowRgn hwnd, CreateRoundRectRgn(20, 20, 800, 800, 20, 20), True
End Sub
Private Sub InCorrectWord()
Static intDontKnow As Integer
intDontKnow = intDontKnow + 1
lblIncorrectscore.Caption = "InCorrect Or Not Known  " & intDontKnow
intListIncorrect = intDontKnow
show_result
End Sub
Private Sub CorrectWord()
Static intCorrect As Integer
intCorrect = intCorrect + 1
lblCorrectScore.Caption = "Correct Answer  " & intCorrect
intListCorrect = intCorrect
show_result
End Sub
Private Sub LoadTempRS()
'create temporary recordset
Set RS1 = New ADODB.Recordset
    With RS1.Fields
    .Append "ID", adSmallInt
    .Append "WordUsed", adBSTR
    End With
    RS1.Open
'Polulate recordset
Dim i As Integer
For i = 0 To cmbcorrect.ListCount - 1
cmbcorrect.ListIndex = i
RS1.AddNew
RS1!ID = i
RS1![wordused] = cmbcorrect.Text
lblUsedWord.Caption = cmbcorrect.Text
Next i
'FindWord
End Sub
Private Sub FindWord()
Dim strFindWord As String
RS1.Filter = strspeakGenie
If RS1.Filter = adFilterConflictingRecords Then
RS1.MoveFirst
Do Until RS1.EOF
RS1.Find strspeakGenie
MsgBox "Word fond"
Loop
Else
MsgBox "OK to move"
End If
End Sub
Private Sub Word_Change()
On Error Resume Next
Dim strsearch As String
Dim lngList As Long
strsearch = Search
'I had a bug here. World Can or Canoe both word can't be used in the same text pad.
'To fix this,Post the word with "." in combobox.
'Search a word upto a stop "."  sign in combobox.
strsearch = strsearch & "."
With cmbcorrect
    If strsearch = "" Then
    .ListIndex = 0
    Else
lngList = SendMessage(cmbcorrect.hwnd, CB_FINDSTRING, -1, ByVal strsearch)
If lngList >= 0 Then
.ListIndex = lngList
lblUsedWord.Caption = "Confirmed  " & strsearch
blnUsedWord = True
strsearch = ""
End If
End If
End With
End Sub
 
Private Sub optTakeTest_Click()
On Error Resume Next
txtPad.Visible = False
List1.Visible = True
End Sub
Private Sub HScrollSpeed_Change()
lblVoiceSpeed.Caption = HScrollSpeed.Value
End Sub
Private Sub show_result()
Static intAttemp As Integer
intAttemp = intAttemp + 1
If intListIncorrect = 0 Then
List1.AddItem ">   " & intAttemp & Space$(15) & intListCorrect & Space$(15) & intListIncorrect
Else
List1.AddItem ">   " & intAttemp & Space$(15) & intListCorrect & Space$(15) & intListIncorrect
End If
End Sub
