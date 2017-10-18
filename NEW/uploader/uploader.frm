VERSION 5.00
Begin VB.Form uploader 
   Caption         =   "RIC's uploader"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   Icon            =   "uploader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "NEW ITEM/F1"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
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
      Height          =   285
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "SAVE/ENTER"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "UPDATE/F3"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "ADD/F2"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text8 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox Text7 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text6 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text5 
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Text4 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3120
      Width           =   7215
   End
   Begin VB.TextBox Text2 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text3 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1680
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
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text9 
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Qty:"
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
      Left            =   7200
      TabIndex        =   21
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Date Received"
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
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Supplier Name:"
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
      Left            =   4320
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Category:"
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
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Item Code:"
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
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "RP"
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
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "CP"
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
      Left            =   720
      TabIndex        =   12
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Model /    Size"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "RIC'z CYCLE PART STOCK ENCODER"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "uploader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aa As Boolean
Public item As Double
Public counter_s As Integer
Public bbb As Boolean
Public qq As Boolean

Private Sub Command1_Click()
bbb = True
add_me

End Sub

Private Sub Command2_Click()
update_me
End Sub

Private Sub Command3_Click()
If Command1.Enabled = True And Command2.Enabled = True And Command4.Enabled = True Then
MsgBox "Please select a command."
Exit Sub
End If
Dim fld As ADODB.Field
Dim rs1 As ADODB.Recordset
Dim conn As ADODB.Connection

Set rs1 = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT CATEGORY FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn


If rs1.EOF = False And Command4.Enabled = fales Then
MsgBox "Item already exist."
Exit Sub
End If


If rs1.EOF And Command4.Enabled = True Then
MsgBox "Item does not exist."
Exit Sub
End If



save_me
End Sub

Private Sub Command4_Click()
qq = False
Text3.Text = ""
Text1.Text = ""
Text9.Text = ""
Text2.Text = ""
Text5.Text = ""
Text4.Text = ""
Text6.Text = ""
Text8.Text = ""
qq = True
new_me
End Sub

Private Sub Form_Load()
bbb = False
aa = False
qq = False
End Sub

Private Sub Text1_Change()

If Command4.Enabled = False Then
Exit Sub
End If

counter_s = 0


If Command2.Enabled = False Then

End If


Dim fld As ADODB.Field
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs7 As ADODB.Recordset
Dim rs9 As ADODB.Recordset
Dim rs8 As ADODB.Recordset

Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
Set rs7 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
Set rs9 = New ADODB.Recordset

conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT CATEGORY FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs3.Open "SELECT MODEL FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs4.Open "SELECT STOCK_ON_HAND from stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs6.Open "SELECT RP FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs7.Open "SELECT CP FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs8.Open "SELECT DATE_RECEIVED FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn
rs9.Open "SELECT ITEM_CODE FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn



If rs1.EOF Then

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

Text8.Text = ""
Text9.Text = ""
End If


Do Until rs1.EOF


If aa = True Then
For Each fld In rs9.Fields
ss = fld.Value
Next
rs9.MoveNext

If ss = Text1.Text Then
counter_s = counter_s + 1
End If
Exit Sub
End If


For Each fld In rs1.Fields
Text3.Text = fld.Value
Next

For Each fld In rs2.Fields
Text4.Text = fld.Value
Next

For Each fld In rs3.Fields
Text2.Text = fld.Value
Next

For Each fld In rs4.Fields
Text9.Text = fld.Value
item = Val(Text9.Text)
Next

For Each fld In rs5.Fields
Text5.Text = fld.Value
Next

For Each fld In rs6.Fields
Text8.Text = fld.Value
Next

For Each fld In rs7.Fields
Text6.Text = fld.Value
Next

If Command1.Enabled = False Then
GoTo sunod_na
End If

For Each fld In rs8.Fields
Text7.Text = fld.Value
Next

sunod_na:

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
rs6.MoveNext
rs7.MoveNext
rs8.MoveNext

item = Val(Text9.Text)
If Command2.Enabled = False Then
item = Val(Text9.Text)
End If


'If counter_s = 0 Then
'MsgBox "Item Code not found!"
'Text1.SetFocus
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Text8.Text = ""
'Text9.Text = ""
'
'Exit Sub
'End If

Loop


End Sub


Private Sub update_me()
Text3.Enabled = True

Text1.SetFocus
'Text9.Text = "# of arived astock"
Command2.Enabled = False
Command1.Enabled = True
Command4.Enabled = True

Text9.Enabled = False
Text9.Locked = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
'Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
aa = False
'Text1.SetFocus
End Sub

Private Sub add_me()

Text3.Enabled = True
Text9.Enabled = True
Command4.Enabled = False
Text9.SetFocus
'Text9.Text = "# of arived astock"
Command2.Enabled = True
Command1.Enabled = False
Command4.Enabled = True

Text9.Enabled = True
Text9.Locked = False
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = True
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
'Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
aa = False
Text1.SetFocus
Text7.Text = Format(Now, "mm/d/yyyy")



End Sub


Private Sub save_me()
'item = Val(Text9.Text)
Dim cp As Double
Dim rp As Double
cp = Val(Text6.Text)
rp = Val(Text8.Text)

If cp = 0 Or rp = 0 Or Text3.Text = "" Or cp >= rp Or Val(Text9.Text) = 0 Or Text1.Text = "" Then
Err.Description = "Please enter the right amount in peso."
GoTo savingErr
End If

On Error GoTo savingErr
Dim margin As Double
margin = (rp - cp) / (Val(Text6.Text))
margin = margin * 100
Dim marg_peso As Double
marg_peso = rp - cp

If Command1.Enabled = True And Command2.Enabled = True And Command4.Enabled = True Then
MsgBox "Please select a command."
Exit Sub
End If
Text1.SetFocus
If counter_s <> 0 Then
MsgBox "Item Code already used!!"
Exit Sub
End If

If counter_s <> 0 And Command2.Enabled = fales Then
MsgBox "Item Code not found!"
Text1.SetFocus
Exit Sub

End If

Dim fld As ADODB.Field
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open


If Command2.Enabled = False Then
'MsgBox item
'MsgBox ((Val(Text9.Text)) + item)
  conn.Execute "UPDATE stock_info SET STOCK_ON_HAND = " & "'" & Val(Text9.Text) & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET ITEM_CODE = " & "'" & Text1.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET MODEL = " & "'" & Text2.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET SUPPLIER_NAME = " & "'" & Text5.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET DESCRIPTION = " & "'" & Text4.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET CP = " & "'" & Text6.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET RP = " & "'" & Text8.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET CATEGORY = " & "'" & Text3.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 conn.Execute "UPDATE stock_info SET MARGIN = " & "'" & margin & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
  conn.Execute "UPDATE stock_info SET MARGIN_PESO = " & "'" & marg_peso & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
End If

If Command4.Enabled = False Then
conn.Execute "INSERT INTO stock_info (CATEGORY, MODEL,DESCRIPTION,ITEM_CODE, DATE_RECEIVED, SUPPLIER_NAME, CP, RP,STOCK_ON_HAND, MARGIN,MARGIN_PESO)" _
& "values ('" & Text3.Text & "', '" & Text2.Text & "', '" & Text4.Text & "', '" & Text1.Text & "', '" & Text7.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & Text8.Text & "', '" & Text9.Text & "', '" & margin & "', '" & marg_peso & "')"
End If


If Command1.Enabled = False Then
 conn.Execute "UPDATE stock_info SET STOCK_ON_HAND = " & "'" & Val(Text9.Text) + (item) & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
conn.Execute "UPDATE stock_info SET DATE_RECEIVED = " & "'" & Text7.Text & "'" & " WHERE ITEM_CODE = '" & Text1.Text & "'"
 End If
 
 
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
'Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

List1.Visible = False

savingErr:

If Err.Description = "" Then
MsgBox "Data successfully saved!"
Else
MsgBox "Please enter the right data needed."
Exit Sub
End If
End Sub


Private Sub new_me()

bbb = True

Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
aa = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
'Text7.Locked = False
Text8.Locked = False
Text9.Locked = False
Text7.Text = Format(Now, "mm/d/yyyy")
Text3.SetFocus

End Sub


Private Sub Text1_GotFocus()
If Command1.Enabled = fales Then
bbb = True
Else
bbb = False
End If


End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
If KeyCode = 112 Then
qq = False
Text3.Text = ""
Text1.Text = ""
Text9.Text = ""
Text2.Text = ""
Text5.Text = ""
Text4.Text = ""
Text6.Text = ""
Text8.Text = ""
qq = True
new_me
End If


End Sub

Private Sub Text2_Change()
If Command1.Enabled = False Or Command2.Enabled = False Then
Exit Sub
End If

If qq = False Then
Exit Sub
End If

List1.Visible = True
List1.Top = 2880
List1.Width = 2655
List1.Left = 1440
List1.Clear

On Error GoTo hh

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim fld As ADODB.Field

Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT MODEL FROM stock_info where MODEL like '" & UCase(Text2.Text) & "%'", conn
looper = 0
Do Until rs1.EOF

For Each fld In rs1.Fields
List1.AddItem fld.Value

If List1.List(looper) = fld.Value Then
GoTo hh
End If
Next

rs1.MoveNext
looper = looper + 1
Loop


hh:


End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then
Text2.Text = List1.List(0)
End If

If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
End Sub

Private Sub Text2_LostFocus()
List1.Visible = False
End Sub

Private Sub Text3_Change()
If Command1.Enabled = False Or Command2.Enabled = False Then
Exit Sub
End If




If qq = False Then
Exit Sub
End If



Dim looper As Integer
List1.Visible = True
List1.Top = 2160
List1.Width = 2655
List1.Left = 1440
List1.Clear

On Error GoTo hh

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim fld As ADODB.Field

Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT CATEGORY FROM stock_info where CATEGORY like '" & UCase(Text3.Text) & "%'", conn
looper = 0
Do Until rs1.EOF

For Each fld In rs1.Fields
List1.AddItem fld.Value

If List1.List(looper) = fld.Value Then
GoTo hh
End If
Next

rs1.MoveNext
looper = looper + 1
Loop


hh:
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then

Text3.Text = List1.List(0)
End If
If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If

End Sub

Private Sub Text3_LostFocus()
List1.Clear
List1.Visible = False
End Sub


Private Sub Text4_Change()

If Command1.Enabled = False Or Command2.Enabled = False Then
Exit Sub
End If


If qq = False Then
Exit Sub
End If



List1.Visible = True
List1.Top = 3600
List1.Width = 7215
List1.Left = 1440
List1.Clear


On Error GoTo hh

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim fld As ADODB.Field

Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT DESCRIPTION FROM stock_info where DESCRIPTION like '" & UCase(Text4.Text) & "%'", conn
looper = 0
Do Until rs1.EOF

For Each fld In rs1.Fields
List1.AddItem fld.Value

If List1.List(looper) = fld.Value Then
GoTo hh
End If
Next

rs1.MoveNext
looper = looper + 1
Loop


hh:



End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 16 Then
Text4.Text = List1.List(0)
End If

If KeyCode = 113 Then

add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
End Sub

Private Sub Text4_LostFocus()
List1.Visible = False
End Sub

Private Sub Text5_Change()

If Command1.Enabled = False Or Command2.Enabled = False Then
Exit Sub
End If

If qq = False Then
Exit Sub
End If


List1.Visible = True
List1.Top = 2880
List1.Width = 3255
List1.Left = 5400
List1.Clear


On Error GoTo hh

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim fld As ADODB.Field

Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT SUPPLIER_NAME FROM stock_info where SUPPLIER_NAME like '" & UCase(Text5.Text) & "%'", conn
looper = 0
Do Until rs1.EOF

For Each fld In rs1.Fields
List1.AddItem fld.Value

If List1.List(looper) = fld.Value Then
GoTo hh
End If
Next

rs1.MoveNext
looper = looper + 1
Loop


hh:



End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then
Text5.Text = List1.List(0)
End If


If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
End Sub

Private Sub Text5_LostFocus()
List1.Visible = False
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then
save_me
End If
End Sub

Private Sub Text9_GotFocus()

If Command1.Enabled = False Then

Text9.Text = ""
End If
'MsgBox item
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
add_me
End If
If KeyCode = 114 Then
update_me
End If
If KeyCode = 13 Then


Dim fld As ADODB.Field
Dim rs1 As ADODB.Recordset
Dim conn As ADODB.Connection

Set rs1 = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT CATEGORY FROM stock_info where ITEM_CODE = " & "'" & UCase(Text1.Text) & "'", conn

If rs1.EOF Then
MsgBox "Item does not exist."
Exit Sub
End If

save_me
End If
End Sub
