VERSION 5.00
Begin VB.Form returnme 
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Lucida Sans"
      Size            =   48
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "return.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "All Item"
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
      Left            =   240
      TabIndex        =   24
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List8 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   5760
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List7 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   22
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3360
      TabIndex        =   21
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2640
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "return.frx":4AC5
      Left            =   360
      List            =   "return.frx":4ACC
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "return.frx":4ADA
      Left            =   960
      List            =   "return.frx":4AE1
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "return.frx":4AF0
      Left            =   1320
      List            =   "return.frx":4AF7
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
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
      Height          =   960
      Left            =   1440
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&Cancel / ESC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "&Return / Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox Text8 
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
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   5775
   End
   Begin VB.TextBox Text7 
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
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2175
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2175
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1935
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date_Sold"
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
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount     paid"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Served By:"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice #"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "returnme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
List1.Visible = False
Text8.Text = "All Items"
Text7.Text = "All Items"

End Sub

Private Sub Command1_Click()

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim fld1 As ADODB.Field
Set rs1 = New ADODB.Recordset

Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

If Check1.Value = 1 Then
conn.Execute "UPDATE sales SET DATE_SOLD = '   ----RETURNED----' " & " WHERE INVOICE = " & "'" & Text1.Text & "'"

Else


conn.Execute "UPDATE sales SET DATE_SOLD = '   ----RETURNED----' " & " WHERE INVOICE = " & "'" & Text1.Text & "' AND ITEM_CODE = '" & Text7.Text & "'"

rs1.Open "SELECT STOCK_ON_HAND FROM stock_info where ITEM_CODE = '" & Text7.Text & "'", conn

For Each fld1 In rs1.Fields
s_o_h = fld1.Value + Val(Text4.Text)
Next

'MsgBox s_o_h
conn.Execute "UPDATE stock_info SET STOCK_ON_HAND = '" & s_o_h & "' WHERE ITEM_CODE = " & "'" & Text7.Text & "'"

End If
Me.Enabled = False
Unload Me
main.Enabled = True
main.Show
'userpassword.Show
End Sub

Private Sub Command2_Click()
Unload Me
main.Enabled = True
main.Show
End Sub

Private Sub Form_Load()
'Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
main.Show
Unload Me
main.item_code.SetFocus
End Sub

Private Sub List1_Click()


worker

End Sub


Private Sub List1_DblClick()

worker
List1.Visible = False
If Text6.Text = "   ----RETURNED----" Then
Exit Sub
Else
'Command1.SetFocuss
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
List1.Visible = False
Command1.Enabled = True
Command1.SetFocus
End If


If KeyCode = 13 Then
If Text6.Text = "   ----RETURNED----" Then
Command1.Enabled = False
Text1.SetFocus
Exit Sub
End If
End If

End Sub

Private Sub Text1_Change()

List1.Visible = True
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear




'List1.AddItem "All Item"

'List5.AddItem "All Item"
'List6.AddItem "All Item"
'List7.AddItem "All Item"
'List8.AddItem "All Item"



Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text7.Text = ""
Text8.Text = ""






Dim conn As ADODB.Connection
Dim fld As ADODB.Field
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs7 As ADODB.Recordset
Dim rs8 As ADODB.Recordset
Dim na As String
Dim add As String
Dim card As String
Dim numb As String


Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
Set rs7 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset



conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT ITEM_CODE FROM sales where INVOICE like '" & Text1.Text & "%'", conn
rs2.Open "SELECT DESCRIPTION FROM sales where INVOICE like '" & Text1.Text & "%'", conn
rs3.Open "SELECT PCS FROM sales where INVOICE like '" & Text1.Text & "%'", conn
rs4.Open "SELECT DATE_SOLD FROM sales where INVOICE like '" & Text1.Text & "%'", conn
rs5.Open "SELECT INVOICE FROM sales where INVOICE = '" & Text1.Text & "'", conn
rs6.Open "SELECT CASHER FROM sales where INVOICE like '" & Text1.Text & "%'", conn
rs7.Open "SELECT DESCRIPTION FROM sales where INVOICE like '" & Text1.Text & "%'", conn
rs8.Open "SELECT TOTAL FROM sales where INVOICE like '" & Text1.Text & "%'", conn


Do Until rs1.EOF
On Error GoTo hi




For Each fld In rs5.Fields
List5.AddItem fld.Value
Next

For Each fld In rs6.Fields
List6.AddItem fld.Value
Next

For Each fld In rs7.Fields
List7.AddItem fld.Value
Next

For Each fld In rs8.Fields
List8.AddItem fld.Value
Next



For Each fld In rs1.Fields
na = fld.Value
'List2.AddItem fld.Value
Next

For Each fld In rs4.Fields
nn = fld.Value
'List3.AddItem fld.Value
Next


For Each fld In rs2.Fields
add = fld.Value
Next

For Each fld In rs3.Fields
add = add & ",  " & fld.Value
List4.AddItem fld.Value
Next


Dim vall() As String

For Each fld In rs4.Fields
vall = Split(fld.Value, " ")
add = add & ", " & Trim(vall(0))
Next



List1.AddItem add
List2.AddItem na
List3.AddItem nn


rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext




Loop

hi:
'MsgBox Err.Description
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Or KeyCode = 38 Then
List1.SetFocus
End If

If KeyCode = 13 Then
End If

If KeyCode = 27 Then
Unload Me
main.Enabled = True
main.Show
End If
End Sub


Private Sub worker()
Dim workKK() As String


Text7.Text = List2.List(List1.ListIndex)
Text4.Text = List4.List(List1.ListIndex)
Text6.Text = List3.List(List1.ListIndex)
workKK = Split(Text6.Text, " ")
'Text6.Text = workKK(0)
Text1.Text = List5.List(List1.ListIndex)
Text2.Text = List6.List(List1.ListIndex)
Text8.Text = List7.List(List1.ListIndex)
Text5.Text = List8.List(List1.ListIndex)

'MsgBox Text7.Text
Dim ret As Integer
Dim datos As String
datos = Text6.Text
ret = Val(Text1.Text)


Dim fld1 As ADODB.Field




Dim looper As Integer

Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs7 As ADODB.Recordset
Dim rs8 As ADODB.Recordset

Dim a1 As String
Dim ax As String
Dim a2 As String
Dim a3 As String
Dim a4 As String
Dim a5 As String
Dim a6 As String

Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset


rs1.Open "SELECT INVOICE FROM sales where INVOICE =  " & "'" & ret & "'" & " And DATE_SOLD = '" & datos & "'", conn
rs2.Open "SELECT CASHER FROM sales where INVOICE =  " & "'" & ret & "'" & " And DATE_SOLD = '" & datos & "'", conn
rs3.Open "SELECT DESCRIPTION FROM sales where INVOICE =  " & "'" & ret & "'" & " And DATE_SOLD = '" & datos & "'", conn
rs4.Open "SELECT ITEM_CODE FROM sales where INVOICE =  " & "'" & ret & "'" & " And DATE_SOLD = '" & datos & "'", conn
rs5.Open "SELECT PCS FROM sales where INVOICE =  " & "'" & ret & "'" & " And DATE_SOLD = '" & datos & "'", conn
rs6.Open "SELECT TOTAL FROM sales where INVOICE =  " & "'" & ret & "'" & " And DATE_SOLD = '" & datos & "'", conn


Do Until rs2.EOF
'search.Text = ""


For Each fld1 In rs2.Fields
Text2.Text = fld1.Value
Next

For Each fld1 In rs3.Fields
Text8.Text = fld1.Value
Next

For Each fld1 In rs4.Fields
Text7.Text = fld1.Value
Next

For Each fld1 In rs5.Fields
Text4.Text = fld1.Value
Next

For Each fld1 In rs6.Fields
Text5.Text = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
Loop



End Sub

