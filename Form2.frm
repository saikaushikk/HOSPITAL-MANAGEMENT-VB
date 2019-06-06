VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFF00&
   Caption         =   "WELCOME"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15690
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "APPOINTMENTS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   11880
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ADD NURSE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   8160
      TabIndex        =   7
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ADD DOCTOR"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   4200
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "VIEW DOCTORS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   4200
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GENERATE BILL"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   11880
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VIEW NURSES"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   8160
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ADD PATIENT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW PATIENTS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "  HOSPITAL MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Form6.Show
End Sub

Private Sub Command2_Click(Index As Integer)
Form7.Show
End Sub

Private Sub Command3_Click(Index As Integer)
Form8.Show
End Sub

Private Sub Command4_Click(Index As Integer)
Form9.Show
End Sub

Private Sub Command5_Click(Index As Integer)
Form4.Show
End Sub

Private Sub Command6_Click(Index As Integer)
Form3.Show
End Sub

Private Sub Command7_Click(Index As Integer)
Form5.Show
End Sub

Private Sub Command8_Click(Index As Integer)
Form10.Show
End Sub
