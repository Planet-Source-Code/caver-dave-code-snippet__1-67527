VERSION 5.00
Begin VB.Form frmAddNewCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ADD NEW CODE"
   ClientHeight    =   9825
   ClientLeft      =   3630
   ClientTop       =   705
   ClientWidth     =   9570
   Icon            =   "frmAddNewCode.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   9570
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE CODE"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   2175
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
      Left            =   8520
      TabIndex        =   3
      Top             =   9240
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   7095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2040
      Width           =   9375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "ENTER SNIPPET NAME IN HERE"
      Top             =   1080
      Width           =   8055
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
      TabIndex        =   0
      Text            =   "SELECT CATEGORY"
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "CODE TITLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "CODE WINDOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "frmAddNewCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim Filehandle As Integer
Dim Msa As String
Dim data As String
Filehandle = FreeFile
Msa = Text1.Text
If Combo1.Text = "Classes" Then
Open App.Path & "\" & Combo1.Text & "\Classes.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Forms" Then
Open App.Path & "\" & Combo1.Text & "\Forms.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Full Projects" Then
Open App.Path & "\" & Combo1.Text & "\Projects.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Functions" Then
Open App.Path & "\" & Combo1.Text & "\Functions.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Miscellaneous" Then
Open App.Path & "\" & Combo1.Text & "\Miscellaneous.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Modules" Then
Open App.Path & "\" & Combo1.Text & "\Modules.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Property Pages" Then
Open App.Path & "\" & Combo1.Text & "\Property_Pages.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "Tips, Tricks And Hints" Then
Open App.Path & "\" & Combo1.Text & "\Tips_Tricks.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
ElseIf Combo1.Text = "User Controls" Then
Open App.Path & "\" & Combo1.Text & "\User_Controls.dat" For Append As #Filehandle   ' Open file for output.
Print #Filehandle, Msa
Close #Filehandle
End If

data = Text2.Text
Open App.Path & "\" & Combo1.Text & "\" & Text1.Text & ".dat" For Append As #Filehandle
Print #Filehandle, data
Close #Filehandle





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
