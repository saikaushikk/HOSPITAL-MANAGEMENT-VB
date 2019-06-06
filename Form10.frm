VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00FFFF00&
   Caption         =   "APPOINTMENTS"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18330
   LinkTopic       =   "Form10"
   ScaleHeight     =   10260
   ScaleWidth      =   18330
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   2655
      Left            =   5880
      TabIndex        =   17
      Top             =   7440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2775
      Left            =   5880
      TabIndex        =   16
      Top             =   4200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LOAD"
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   8040
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   5880
      TabIndex        =   14
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   2520
      TabIndex        =   12
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2400
      TabIndex        =   8
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE APPOINTMENT ID"
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "<--TREATMENTS AVAILABLE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15000
      TabIndex        =   20
      Top             =   8520
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "<--DOCTORS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15120
      TabIndex        =   19
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<--  APPOINTMENTS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15000
      TabIndex        =   18
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "DOCTOR ID:"
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
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "TREATMENT ID:"
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
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "DATE OF APPOINTMENT:"
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
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "     TIME:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "PATIENT ID:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "                APPOINTMENTS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form10"
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
      
        rs.Open "insert into appointment values(seq1.nextval,'" & Text2.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "')", conn
        MsgBox "APPOINTMENT added successfully"

End Sub


Private Sub Command2_Click()
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Delete As String
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=MSDAORA.1;Password=Harshitha$99;User ID=system;Persist Security Info=True"
Delete = "delete from appointment where appointment_id= " & Text1.Text
conn.Execute (Delete)
MsgBox ("Deleted")
End Sub

Private Sub Command3_Click()
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim an As String
Set con = New ADODB.Connection
con.Open ("Provider=MSDAORA.1;Password=Harshitha$99;User ID=system;Persist Security Info=True")
cmd.ActiveConnection = con
con.CursorLocation = adUseClient
If rs.State = adStateClosed Then
rs.Open "select * from appview order by appointment_id", con, adOpenDynamic, adLockBatchOptimistic
End If
Set DataGrid1.DataSource = rs
End Sub


Private Sub Form_Load()
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
conn.CursorLocation = adUseClient
 conn.Open ("Provider=MSDAORA.1;Password=Harshitha$99;User ID=system;Persist Security Info=True")
      
        rs1.Open "select * from appview order by appointment_id", conn, adOpenDynamic, adLockBatchOptimistic
        Set DataGrid1.DataSource = rs1
        rs2.Open "select * from doctor", conn, adOpenDynamic, adLockBatchOptimistic
        Set DataGrid2.DataSource = rs2
        rs3.Open "select * from treatment", conn, adOpenDynamic, adLockBatchOptimistic
        Set DataGrid3.DataSource = rs3
End Sub


