VERSION 5.00
Begin VB.Form addclient 
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   Icon            =   "addclient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "addclient.frx":1C64C
   ScaleHeight     =   4665
   ScaleWidth      =   7650
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox num 
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
      Left            =   4560
      TabIndex        =   4
      Text            =   "num"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox card 
      Enabled         =   0   'False
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
      Left            =   1560
      TabIndex        =   3
      Text            =   "card"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox add 
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
      Left            =   1440
      TabIndex        =   1
      Text            =   "add"
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox ngalan 
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
      Left            =   1440
      TabIndex        =   0
      Text            =   "ngalan"
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Control Number:"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Number"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "addclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
'Check1.Value = 1
Else: Check1.Value = 1
End If
If Check1.Value = 1 Then
card.Enabled = True
card.Text = ""
Else
card.Enabled = False
card.Text = "unregistere"
End If
End Sub

Private Sub Command1_Click()
main.Enabled = True
Unload Me
main.Show
End Sub

Private Sub Command2_Click()


add_na
End Sub

Private Sub Form_Load()
ngalan.Text = ""
num.Text = ""
add.Text = ""
card.Text = "unregistered"
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim id As String
Dim looper As Integer
Dim xxx As Integer
Dim id2 As Integer
looper = 0

id = ""
Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs.Open "SELECT CUSTOMER_ID FROM customer", conn

If looper = 0 And rs.EOF = True Then
id = "0"
End If
Text1.Text = Val(id) + 1

Do Until rs.EOF


For Each fld In rs.Fields
id = fld.Value


Next
Text1.Text = Val(id) + 1

rs.MoveNext
looper = looper + 1
Loop

End Sub

Private Sub paminaw(keyhit)

 If keyhit = 27 Then
 main.Enabled = True
Unload Me
main.Show

 End If
 
If keyhit = 13 Then
add_na
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
main.Show
main.item_code.SetFocus
End Sub

Private Sub ngalan_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub


Private Sub add_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub



Private Sub card_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub



Private Sub num_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub add_na()
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection

conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open
conn.Execute "INSERT INTO customer (CARD_NUMBER, CUSTOMER_ID, NAME ,ADDRESS, NUMBER) values ('" & addclient.card.Text & "', '" & Text1.Text & "', '" & ngalan.Text & "', '" & add.Text & "','" & num.Text & "')"
'conn.Execute "INSERT INTO customer (CUSTOMER_ID, NAME ,ADDRESS, NUMBER, POINTS, CREDIT, BAL) values ('" & addclient.card.Text & "', '" & Text1.Text & "', '" & addclient.ngalan & "', '" & addclient.add & "','" & addclient.num & "' , ' ', ' ', ' ')"

main.Enabled = True
Unload Me
main.Show
End Sub


