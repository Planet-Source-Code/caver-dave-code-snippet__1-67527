VERSION 5.00
Begin VB.Form frmLister 
   Caption         =   "CODE DIRECTORY LISTER"
   ClientHeight    =   9315
   ClientLeft      =   2790
   ClientTop       =   1500
   ClientWidth     =   11490
   Icon            =   "frmLister.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   11490
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4440
      Top             =   8760
   End
   Begin VB.CommandButton Command5 
      Caption         =   "COPY DISPLAYED CODE"
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
      Left            =   5040
      TabIndex        =   7
      Top             =   8760
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   7860
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD NEW SNIPPET"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OPEN CODE"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Text            =   "SELECT CATEGORY"
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&XIT"
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
      Left            =   10440
      TabIndex        =   2
      Top             =   8760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   7935
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmLister.frx":0CCA
      Top             =   720
      Width           =   8295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   7680
      Width           =   1215
   End
End
Attribute VB_Name = "frmLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Dim Filehandle As Integer
 Dim lne$
Filehandle = FreeFile

List1.Clear

If Combo1.Text = "Classes" Then
Open App.Path & "\" & Combo1.Text & "\Classes.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Forms" Then
Open App.Path & "\" & Combo1.Text & "\Forms.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Full Projects" Then
Open App.Path & "\" & Combo1.Text & "\Projects.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Functions" Then
Open App.Path & "\" & Combo1.Text & "\Functions.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Miscellaneous" Then
Open App.Path & "\" & Combo1.Text & "\Miscellaneous.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Modules" Then
Open App.Path & "\" & Combo1.Text & "\Modules.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Property Pages" Then
Open App.Path & "\" & Combo1.Text & "\Property_Pages.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "Tips, Tricks And Hints" Then
Open App.Path & "\" & Combo1.Text & "\Tips_Tricks.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
ElseIf Combo1.Text = "User Controls" Then
Open App.Path & "\" & Combo1.Text & "\User_Controls.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
List1.AddItem lne$
Loop
Close #Filehandle
End If
Text1.Text = ""
End Sub

Private Sub Command3_Click()
frmAddNewCode.Show
End Sub

Private Sub Command4_Click()
Clipboard.Clear
   Clipboard.SetText Text1.SelText
End Sub

Private Sub Command5_Click()
Text1.SetFocus
Timer1.Enabled = True
Timer1.Interval = 250
End Sub

Private Sub Form_Load()
Dim Filehandle As Integer
 Dim lne$
Filehandle = FreeFile
Open App.Path & "\Folders.dat" For Input As #Filehandle
Do While Not EOF(Filehandle)
Line Input #Filehandle, lne$
Combo1.AddItem lne$
Loop
Close #Filehandle
End Sub

Private Sub List1_Click()
Dim Filehandle As Integer
Dim data As String
Dim FileLength
Dim var1

Filehandle = FreeFile
Open App.Path & "\" & Combo1.Text & "\" & List1.Text & ".dat" For Input As #Filehandle
FileLength = LOF(Filehandle)
var1 = Input(FileLength, #Filehandle)
Text1.Text = var1
End Sub

Private Sub Text1_GotFocus()
'highlights all the text in the text box on cursor entry
  Call SelectMe(Me.ActiveControl)
End Sub

Private Sub SelectMe(This As TextBox)
'select all text
This.SelStart = 0
This.SelLength = Len(This.Text)
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval = 250 Then
Timer1.Enabled = False
Call Command4_Click
Text1.SelLength = 0
End If
End Sub
