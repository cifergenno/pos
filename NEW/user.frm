VERSION 5.00
Begin VB.Form user 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "user.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   4800
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   3120
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "&Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&Edit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5160
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme Password"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "User List"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True

Dim conn As ADODB.Connection

Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

conn.Execute "INSERT INTO users (NAME, ID, PASSWORD)" _
& "values ('" & Text1.Text & "', '" & Text2.Text & "', '" & Text3.Text & "')"

Unload Me
main.Enabled = True
main.Show
End Sub

Private Sub Command4_Click()

Dim conn As ADODB.Connection
Set conn = New ADODB.Connection

conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open
conn.Execute "UPDATE users SET NAME = '" & Text1.Text & "' WHERE ID = '" & List2.List(List1.ListIndex) & "'"
conn.Execute "UPDATE users SET PASSWORD = '" & Text3.Text & "' WHERE ID = '" & List2.List(List1.ListIndex) & "'"
conn.Execute "UPDATE users SET ID = '" & Text2.Text & "' WHERE ID = '" & List2.List(List1.ListIndex) & "'"

Unload Me
main.Enabled = True
main.Show

End Sub

Private Sub Command5_Click()
Unload Me
main.Enabled = True
main.Show
End Sub

Private Sub Form_Load()

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim fld As ADODB.Field

Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset

conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open



rs1.Open "SELECT NAME FROM users", conn
rs2.Open "SELECT ID FROM users", conn
rs3.Open "SELECT PASSWORD FROM users", conn

Do Until rs1.EOF

For Each fld In rs1.Fields
List1.AddItem fld.Value
Next

For Each fld In rs2.Fields
List2.AddItem fld.Value
Next

For Each fld In rs3.Fields
List3.AddItem fld.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show

main.Enabled = True
main.item_code.SetFocus
Unload Me
End Sub

Private Sub List1_Click()
Text1.Text = List1.List(List1.ListIndex)
Text2.Text = List2.List(List1.ListIndex)
Text3.Text = List3.List(List1.ListIndex)
Command1.Enabled = True
Command3.Enabled = True
'Command4.Enabled = True
End Sub


Private Sub paminaw(keyhit)

If keyhit = 13 Then
End If

If keyhit = 27 Then
Unload Me
main.Enabled = True
main.Show
End If

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub
