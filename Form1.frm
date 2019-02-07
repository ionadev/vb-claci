VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdeq 
      Caption         =   "="
      Height          =   2775
      Left            =   7800
      TabIndex        =   18
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdpm 
      Caption         =   "+/-"
      Height          =   2175
      Left            =   7800
      TabIndex        =   17
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmddot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   16
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      Height          =   1095
      Left            =   2280
      TabIndex        =   15
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdc 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      Height          =   1095
      Left            =   720
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      Height          =   1095
      Left            =   2280
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      Height          =   1095
      Left            =   3840
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      Height          =   1095
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      Height          =   1095
      Left            =   3840
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      Height          =   1095
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdd 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   7
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdmul 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdmi 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdp 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      Height          =   1095
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   1095
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   1095
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtres 
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Double
Dim operand1 As Double, operand2 As Double
Dim cleardisplay As Boolean
Dim operator As String

Private Sub cmd0_Click()
txtres.Text = txtres.Text + "0"
End Sub

Private Sub cmd1_Click()
txtres.Text = txtres.Text + "1"
End Sub

Private Sub cmd2_Click()
txtres.Text = txtres.Text + "2"
End Sub

Private Sub cmd3_Click()
txtres.Text = txtres.Text + "3"
End Sub

Private Sub cmd4_Click()
txtres.Text = txtres.Text + "4"
End Sub

Private Sub cmd5_Click()
txtres.Text = txtres.Text + "5"
End Sub

Private Sub cmd6_Click()
txtres.Text = txtres.Text + "6"
End Sub

Private Sub cmd7_Click()
txtres.Text = txtres.Text + "7"
End Sub

Private Sub cmd8_Click()
txtres.Text = txtres.Text + "8"
operand2 = 8
End Sub

Private Sub cmd9_Click()
txtres.Text = txtres.Text + "9"
End Sub

Private Sub cmdc_Click()
txtres.Text = ""
End Sub

Private Sub cmdd_Click()
operand1 = Val(txtres.Text)
operator = "/"
txtres = ""
End Sub

Private Sub cmdeq_Click()
If IsNull(operand1) Then
MsgBox ("Please Give The First nput , Select Operation, and retry")
End If
operand2 = txtres.Text
If operator = "+" Then result = operand1 + operand2
If operator = "-" Then result = operand1 - operand2
If operator = "*" Then result = operand1 * operand2
End Sub

Private Sub cmdmi_Click()
operand1 = Val(txtres.Text)
operator = "-"
txtres = ""
End Sub

Private Sub cmdmul_Click()
operand1 = Val(txtres.Text)
operator = "*"
txtres = ""
End Sub

Private Sub cmdp_Click()
operand1 = Val(txtres.Text)
operator = "+"
txtres = ""
End Sub

