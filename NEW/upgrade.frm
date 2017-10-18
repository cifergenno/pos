VERSION 5.00
Begin VB.Form upgrade 
   Caption         =   "Suki Upgrade Form"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
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
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Alter Changes"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox card 
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
      Left            =   2640
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   5895
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   5895
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
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
      Left            =   3480
      TabIndex        =   0
      Text            =   "Enter Control #"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Reference"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Update Customer/Upgrade To Suki Card Holder"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control Number"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "New Suki Card Number"
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
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2535
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
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   855
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
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   855
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
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "upgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Public rs1 As ADODB.Recordset
Public rs2 As ADODB.Recordset
Public rs3 As ADODB.Recordset
Public rs4 As ADODB.Recordset
Public fld1 As ADODB.Field

Private Sub add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
saveMe
End If
End Sub

Private Sub card_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
saveMe
End If


End Sub

Private Sub Command1_Click()

Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

conn.Execute "UPDATE customer SET NAME = '" & ngalan.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"
conn.Execute "UPDATE customer SET ADDRESS = '" & add.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"
conn.Execute "UPDATE customer SET NUMBER = '" & num.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"
conn.Execute "UPDATE customer SET CARD_NUMBER = '" & card.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"

MsgBox ngalan & "'s profile was updated!"
Unload Me
main.Show
main.Enabled = True
End Sub


Private Sub saveMe()
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

conn.Execute "UPDATE customer SET NAME = '" & ngalan.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"
conn.Execute "UPDATE customer SET ADDRESS = '" & add.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"
conn.Execute "UPDATE customer SET NUMBER = '" & num.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"
conn.Execute "UPDATE customer SET CARD_NUMBER = '" & card.Text & "' WHERE CUSTOMER_ID = " & "'" & Text1.Text & "'"

MsgBox ngalan & "'s profile was updated!"
Unload Me
main.Show
main.Enabled = True
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub ngalan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
saveMe
End If

End Sub

Private Sub num_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
saveMe
End If
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then

Unload Me
End If



If KeyCode = 13 Then

If ngalan.Text <> "" Then
saveMe
Exit Sub
End If



Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim fld1 As ADODB.Field

On Error GoTo sibat
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset

rs1.Open "SELECT NAME FROM customer where CUSTOMER_ID = '" & Text1.Text & "'", conn
rs2.Open "SELECT ADDRESS FROM customer where CUSTOMER_ID = '" & Text1.Text & "'", conn
rs3.Open "SELECT NUMBER FROM customer where CUSTOMER_ID = '" & Text1.Text & "'", conn
rs4.Open "SELECT CARD_NUMBER FROM customer where CUSTOMER_ID = '" & Text1.Text & "'", conn

'Do Until rs4.EOF

'rs4.MoveNext
'Loop

Do Until rs1.EOF
For Each fld1 In rs4.Fields
card.Text = fld1.Value
Next
For Each fld1 In rs1.Fields
ngalan.Text = fld1.Value
Next
For Each fld1 In rs2.Fields
add.Text = fld1.Value
Next
For Each fld1 In rs3.Fields
num.Text = fld1.Value
Next




rs4.MoveNext
rs1.MoveNext
rs2.MoveNext
rs3.MoveNext

Loop




End If


sibat:



Exit Sub
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
saveMe
End If
End Sub
