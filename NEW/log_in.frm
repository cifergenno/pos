VERSION 5.00
Begin VB.Form log_in 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Log In"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   1
      Left            =   6000
      TabIndex        =   9
      Top             =   960
      Width           =   735
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         Index           =   1
         X1              =   0
         X2              =   720
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         Index           =   1
         X1              =   360
         X2              =   360
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         Index           =   1
         X1              =   720
         X2              =   0
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         Index           =   1
         X1              =   720
         X2              =   0
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1440
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   735
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         Index           =   0
         X1              =   720
         X2              =   0
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000C0&
         Index           =   0
         X1              =   720
         X2              =   0
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   120
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         Index           =   0
         X1              =   0
         X2              =   720
         Y1              =   120
         Y2              =   600
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "LOG IN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "USER LOG IN"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "Ric'z Cycle Parts and Accessories  Center"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Tagbilaran City, Bohol, Philippines"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7575
   End
End
Attribute VB_Name = "log_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ihap As Integer
Private Sub Command1_Click()
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim fld As ADODB.Field
Dim rs3 As ADODB.Recordset
Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Dim name As String
Dim pass As String
Dim idd As String


conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT ID FROM users where ID = '" & Text1.Text & "'", conn
rs2.Open "SELECT NAME FROM users where ID = '" & Text1.Text & "'", conn
rs3.Open "SELECT PASSWORD from users where ID = '" & Text1.Text & "'", conn
Do Until rs1.EOF

For Each fld In rs1.Fields
idd = fld.Value
Next

For Each fld In rs2.Fields
name = fld.Value
Next

For Each fld In rs3.Fields
pass = fld.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
Loop


If LCase(Text1.Text) = LCase(idd) Then
If LCase(text2.Text) = LCase(pass) Then
main.casher.Text = name
Unload Me
loader.Show

Else
MsgBox ("Password and User ID did not match!!!")
End If
Else
MsgBox ("Password and User ID did not match!!!")
End If

End Sub

Public Sub tuyok()


For X = 1 To 4

Next

End Sub

Private Sub Form_Load()
Line1(0).Visible = False
Line2(0).Visible = False
Line3(0).Visible = False
Line4(0).Visible = False
Line1(0).Visible = False
Line2(0).Visible = False
Line3(0).Visible = False
Line4(0).Visible = False
Timer1.Enabled = True
ihap = 1
End Sub

Private Sub Timer1_Timer()

'MsgBox ""


If ihap = 1 Then
Line1(0).Visible = True
Line2(0).Visible = False
Line3(0).Visible = False
Line4(0).Visible = False
Line1(1).Visible = False
Line2(1).Visible = False
Line3(1).Visible = False
Line4(1).Visible = True
End If

If ihap = 2 Then
Line1(0).Visible = False
Line2(0).Visible = True
Line3(0).Visible = False
Line4(0).Visible = False
Line1(1).Visible = False
Line2(1).Visible = False
Line3(1).Visible = True
Line4(1).Visible = False
End If

If ihap = 3 Then
Line1(0).Visible = False
Line2(0).Visible = False
Line3(0).Visible = True
Line4(0).Visible = False
Line1(1).Visible = False
Line2(1).Visible = True
Line3(1).Visible = False
Line4(1).Visible = False
End If

If ihap = 4 Then
Line1(0).Visible = False
Line2(0).Visible = False
Line3(0).Visible = False
Line4(0).Visible = True
Line1(1).Visible = True
Line2(1).Visible = False
Line3(1).Visible = False
Line4(1).Visible = False
End If

ihap = ihap + 1
If ihap = 5 Then
ihap = 1
End If
End Sub
