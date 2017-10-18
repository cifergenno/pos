VERSION 5.00
Begin VB.Form history2 
   Caption         =   "History Viewer"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   6480
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
      Begin VB.OptionButton Option5 
         Caption         =   "All Client"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "All Suki Card"
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
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
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
         Height          =   405
         Left            =   5640
         TabIndex        =   8
         Top             =   2160
         Width           =   1815
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
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   2160
         Width           =   1815
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
         Height          =   405
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Suki Card"
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
         Left            =   4800
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Credit"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Card Number:"
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
         Left            =   5640
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Contact:"
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
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "History Viewer"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "history2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

customer.Text8.Text = Text3
Command2.SetFocus

If Option5.Value = True Then
mga_taw.Show
End If

If Option2.Value = True Then
customer.Text8.Text = Text3.Text
customer.Show
customer.Text8.Text = Text3.Text
End If

If Option3.Value = True Then
suki.Text8.Text = Text4.Text
suki.Show
suki.Text8.Text = Text4.Text
End If

If Option1.Value = True Then
mga_utang.Show
End If

If Option4.Value = True Then
mga_suki.Show
End If

End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
main.Enabled = True
Unload Me
main.item_code.SetFocus
End Sub

Private Sub Option1_Click()
Command1.SetFocus
If KeyCode = 27 Then

Unload Me
End If
'Label2.Caption = "Start date:"
'Label3.Caption = "End Date:"
'Label4.Caption = ""
'Frame2.Visible = True
'Frame3.Visible = True
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Option2_Click()
Label2.Caption = "Name:"
Label3.Caption = "Contact:"
Label4.Caption = "Card Number:"
'Frame2.Visible = False
'Frame3.Visible = False
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Option3_Click()
Label2.Caption = "Name:"
Label3.Caption = "Contact:"
Label4.Caption = "Suki Card Number:"
'Frame2.Visible = False
'Frame3.Visible = False
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
Command1.SetFocus
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Option5_Click()
Command1.SetFocus
End Sub

Private Sub Text1_Click()
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If

If KeyCode = 13 Then

If Option1.Value = True Then
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

If Option2.Value = True Then
rs1.Open "SELECT NUMBER FROM customer where NAME = '" & Text1.Text & "'", conn
rs2.Open "SELECT CUSTOMER_ID FROM customer where NAME = '" & Text1.Text & "'", conn
Do Until rs1.EOF
For Each fld1 In rs1.Fields
Text2.Text = fld1.Value
Next
For Each fld1 In rs2.Fields
Text3.Text = fld1.Value
Next
rs1.MoveNext
rs2.MoveNext
Loop
End If

If Option3.Value = True Then
rs1.Open "SELECT NUMBER FROM customer where NAME = '" & Text1.Text & "'", conn
rs2.Open "SELECT CARD_NUMBER FROM customer where NAME = '" & Text1.Text & "'", conn
rs3.Open "SELECT CUSTOMER_ID FROM customer where NAME = '" & Text1.Text & "'", conn
Do Until rs1.EOF
For Each fld1 In rs1.Fields
Text2.Text = fld1.Value
Next
For Each fld1 In rs2.Fields
Text3.Text = fld1.Value
Next

For Each fld1 In rs3.Fields
Text4.Text = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
Loop

End If

Command1.SetFocus
sibat:

Exit Sub
End If

End Sub

Private Sub Text2_Click()
Text1.Text = ""
Text3.Text = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If

If Option1.Value = True Then
Exit Sub
End If



If KeyCode = 13 Then

Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim fld1 As ADODB.Field

On Error GoTo sibat
'Text12.Text = Text1.Text
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset

If Option2.Value = True Then
rs1.Open "SELECT NAME FROM customer where NUMBER = '" & Text2.Text & "'", conn
rs2.Open "SELECT CUSTOMER_ID FROM customer where NUMBER = '" & Text2.Text & "'", conn
Do Until rs1.EOF
For Each fld1 In rs1.Fields
Text1.Text = fld1.Value
Next
For Each fld1 In rs2.Fields
Text3.Text = fld1.Value
Next
rs1.MoveNext
rs2.MoveNext
Loop
End If

If Option3.Value = True Then
rs1.Open "SELECT NAME FROM customer where NUMBER = '" & Text2.Text & "'", conn
rs2.Open "SELECT CARD_NUMBER FROM customer where NUMBER = '" & Text2.Text & "'", conn
rs3.Open "SELECT CUSTOMER_ID FROM customer where NUMBER = '" & Text2.Text & "'", conn
Do Until rs1.EOF
For Each fld1 In rs1.Fields
Text1.Text = fld1.Value
Next
For Each fld1 In rs2.Fields
Text3.Text = fld1.Value
Next
For Each fld1 In rs3.Fields
Text4.Text = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
Loop

End If
Command1.SetFocus
sibat:
Command1.SetFocus
Exit Sub
End If

End Sub

Private Sub Text3_Change()
'Text2.Text = ""
'Text1.Text = ""
End Sub

Private Sub Text3_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
If KeyCode = 13 Then
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim fld1 As ADODB.Field

'On Error GoTo sibat
'Text12.Text = Text1.Text
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset

If Option2.Value = True Then
rs1.Open "SELECT NAME FROM customer where CUSTOMER_ID = '" & Text3.Text & "'", conn
rs2.Open "SELECT NUMBER FROM customer where CUSTOMER_ID = '" & Text3.Text & "'", conn
Do Until rs1.EOF
For Each fld1 In rs1.Fields
Text1.Text = fld1.Value
Next
For Each fld1 In rs2.Fields
Text2.Text = fld1.Value
Next
rs1.MoveNext
rs2.MoveNext
Loop
End If





If Option3.Value = True Then
rs1.Open "SELECT NAME FROM customer where CARD_NUMBER = '" & Text3.Text & "'", conn
rs2.Open "SELECT NUMBER FROM customer where CARD_NUMBER = '" & Text3.Text & "'", conn
rs3.Open "SELECT CUSTOMER_ID FROM customer where CARD_NUMBER = '" & Text3.Text & "'", conn


Do Until rs1.EOF

For Each fld1 In rs1.Fields
Text1.Text = fld1.Value
Next

For Each fld1 In rs2.Fields
Text2.Text = fld1.Value
Next

For Each fld1 In rs3.Fields
Text4.Text = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
Loop

End If


If Option1.Value = True Then
Exit Sub
End If




sibat:
Command1.SetFocus
Exit Sub
Command1.SetFocus
End If

End Sub
