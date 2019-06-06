VERSION 5.00
Begin VB.Form Form1A 
   BackColor       =   &H00FFFF00&
   Caption         =   "REGISTER"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   LinkTopic       =   "Form3"
   ScaleHeight     =   4935
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "ALREADY REGISTERED?"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "USERNAME:"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "        REGISTER"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
 conn.Open ("Provider=MSDAORA.1;Password=Harshitha$99;User ID=system;Persist Security Info=True")
      
        rs.Open "insert into userlogin values('" & Text1.Text & "','" & Text2.Text & "')", conn
        MsgBox "Registered successfully"
Form1.Show
Form1A.Hide

End Sub


