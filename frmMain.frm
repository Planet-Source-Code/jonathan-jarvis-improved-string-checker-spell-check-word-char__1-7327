VERSION 5.00
Object = "{28BC48A4-13C5-11D4-93F0-38FD09C10000}#6.0#0"; "CHKSTR32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String Checker"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin CheckString.SCheck SCheck1 
      Left            =   360
      Top             =   360
      _ExtentX        =   1349
      _ExtentY        =   1349
   End
   Begin VB.CommandButton cmdWord 
      Caption         =   "Word Count"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdChar 
      Caption         =   "Char Count"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSpell 
      Caption         =   "Spell Check"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txttest 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChar_Click()
SCheck1.countchar (txttest.Text)
End Sub

Private Sub cmdSpell_Click()
txttest.Text = SCheck1.checkspell(txttest.Text)
End Sub

Private Sub cmdWord_Click()
SCheck1.countwords (txttest.Text)
End Sub

Private Sub Form_Load()
SCheck1.wordlink
SCheck1.showmessages (True)
End Sub
