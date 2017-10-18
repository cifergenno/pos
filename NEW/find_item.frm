VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form find_item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mater List"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "find_item.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   14505
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "find_item.frx":747C
      Left            =   7560
      List            =   "find_item.frx":7489
      TabIndex        =   1
      Text            =   "Stock on Hand"
      Top             =   7320
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   14055
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   1815
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   9
         Cols            =   9
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "MARGIN"
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
         Left            =   10800
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         Caption         =   "CP"
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
         Left            =   8520
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "RP"
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
         Left            =   9720
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Supplier Name"
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
         Left            =   11880
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Item Code"
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
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Model/Size"
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
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Description"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Category"
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
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   195
      Left            =   960
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter"
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
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel/ ESC"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   1455
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
      Left            =   2520
      TabIndex        =   0
      Top             =   7320
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "find_item.frx":74B2
      Left            =   240
      List            =   "find_item.frx":74C5
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Description"
      Top             =   7320
      Width           =   2055
   End
End
Attribute VB_Name = "find_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public conn As ADODB.Connection
Public fld1 As ADODB.Field
Public fld2 As ADODB.Field
Public fld3 As ADODB.Field
Public fld4 As ADODB.Field
Public fld5 As ADODB.Field
Public fld6 As ADODB.Field
Public rs1 As ADODB.Recordset
Public rs2 As ADODB.Recordset
Public rs3 As ADODB.Recordset
Public rs4 As ADODB.Recordset
Public rs5 As ADODB.Recordset
Public rs6 As ADODB.Recordset
Public rs7 As ADODB.Recordset
Public rs8 As ADODB.Recordset
Public rs9 As ADODB.Recordset
Public rs10 As ADODB.Recordset
Dim axx As Boolean

Dim textb As Integer
Dim b As Long



Private Sub Combo1_Click()
Combo1.Locked = True
End Sub

Private Sub Combo1_DropDown()
Combo1.Locked = False
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Text1.Text = ""
If KeyCode = 114 Then
Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
Unload Me
main.Enabled = True
main.Show
main.item_code.SetFocus
End Sub

Private Sub Command2_Click()
trabaho
End Sub

Private Sub Form_Load()
'axx = True
'Text1.SetFocus







Me.Show

grid.ColWidth(0) = 300
grid.ColWidth(1) = 1500
grid.ColWidth(2) = 3100
grid.ColWidth(3) = 1830
grid.ColWidth(4) = 1450
grid.ColWidth(5) = 1250
grid.ColWidth(6) = 1000
grid.ColWidth(7) = 1050
grid.ColWidth(8) = 2050
grid.ColAlignment(7) = 1
textb = 1
b = 0


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
Set rs7 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
Set rs9 = New ADODB.Recordset
Set rs10 = New ADODB.Recordset

rs1.Open "SELECT CATEGORY FROM stock_info", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info", conn
rs3.Open "SELECT MODEL FROM stock_info", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info", conn
rs6.Open "SELECT RP FROM stock_info", conn
rs7.Open "SELECT CP FROM stock_info", conn
rs8.Open "SELECT MARGIN FROM stock_info", conn
rs9.Open "SELECT MARGIN_PESO FROM stock_info", conn
rs10.Open "SELECT STOCK_ON_HAND FROM stock_info", conn

For X = 0 To 8
grid.ColAlignment(X) = 1
Next

Dim ab As Integer
Dim ba As Integer
b = 0
d = 0
Do Until rs1.EOF
d = d + 1


b = b + 1
grid.Rows = b

For Each fld1 In rs1.Fields
grid.TextMatrix(b - 1, 1) = fld1.Value
Next

For Each fld1 In rs8.Fields
grid.TextMatrix(b - 1, 11) = fld1.Value
Next
rs8.MoveNext

For Each fld1 In rs2.Fields
grid.TextMatrix(b - 1, 2) = fld1.Value
Next


For Each fld1 In rs3.Fields
grid.TextMatrix(b - 1, 3) = fld1.Value
Next

For Each fld1 In rs4.Fields
grid.TextMatrix(b - 1, 4) = fld1.Value
Next

For Each fld1 In rs5.Fields
grid.TextMatrix(b - 1, 8) = fld1.Value
Next

For Each fld1 In rs6.Fields
grid.TextMatrix(b - 1, 6) = "" & fld1.Value
Next

For Each fld1 In rs7.Fields
grid.TextMatrix(b - 1, 5) = "" & fld1.Value
Next
 
 For Each fld1 In rs9.Fields
 grid.TextMatrix(b - 1, 7) = fld1.Value
 Next

For Each fld1 In rs10.Fields
grid.TextMatrix(b - 1, 10) = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
rs6.MoveNext
rs7.MoveNext
rs9.MoveNext
rs10.MoveNext
'rs8.MoveNext
Loop

grid.Sort = 1
End Sub

Private Sub List1_Click()
item_desc.Text1 = List2.List(List1.ListIndex)
item_desc.Text2 = List4.List(List1.ListIndex)


Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT RP FROM stock_info where ITEM_CODE = " & item_desc.Text2, conn
For Each fld1 In rs1.Fields
item_desc.Text5 = fld1.Value
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = True
main.Show
main.item_code.SetFocus
End Sub

Private Sub grid_Click()


Text4.Text = grid.Row
Text2.Text = grid.TextMatrix(grid.Row, 4)
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
rs1.Open "SELECT RP FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn
rs8.Open "SELECT STOCK_ON_HAND FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn

Do Until rs1.EOF
For Each fld1 In rs1.Fields
item_desc.Text5 = fld1.Value
Next

For Each fld1 In rs8.Fields
item_desc.Text3 = fld1.Value
Next

If Combo2.Text = "Stock on Hand" Then

Text3.Text = grid.TextMatrix(grid.Row, 10)
End If


If Combo2.Text = "Retail Price" Then

Text3.Text = "Php." & grid.TextMatrix(grid.Row, 6)
End If

If Combo2.Text = "Margin" Then
5
ff = Split(grid.TextMatrix(grid.Row, 11), ".")
grid.TextMatrix(grid.Row, 11) = Left(grid.TextMatrix(grid.Row, 11), Len(ff(0)) + 3)
Text3.Text = grid.TextMatrix(grid.Row, 11) & "%"
End If




rs8.MoveNext
rs1.MoveNext
Loop
End Sub

Private Sub grid_DblClick()
trabaho
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)







If KeyCode = 13 Then
trabaho
End If



If KeyCode = 27 Then
Unload Me
main.Enabled = True
main.Show
main.item_code.Text = ""

End If

If KeyCode = 114 Then
Text1.SetFocus
Text1.Text = ""
End If
End Sub

Private Sub grid_RowColChange()
If Combo2.Text = "Stock on Hand" Then
Text3.Text = grid.TextMatrix(grid.Row, 10)
End If


If Combo2.Text = "Retail Price" Then
Text3.Text = "Php." & grid.TextMatrix(grid.Row, 6)
End If

If Combo2.Text = "Margin" Then
Dim ff() As String

ff = Split(grid.TextMatrix(grid.Row, 11), ".")
grid.TextMatrix(grid.Row, 11) = Left(grid.TextMatrix(grid.Row, 11), Len(ff(0)) + 3)
Text3.Text = grid.TextMatrix(grid.Row, 11) & "%"
End If


End Sub

Private Sub Text1_Change()
Text1.SetFocus

grid.Sort = 1

On Error GoTo outme


Dim looper As Integer

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
Set rs6 = New ADODB.Recordset
Set rs7 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
Set rs9 = New ADODB.Recordset
Set rs10 = New ADODB.Recordset
If Combo1.Text = "Description" Then

rs1.Open "SELECT CATEGORY FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs3.Open "SELECT MODEL FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs6.Open "SELECT RP FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs7.Open "SELECT CP FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs8.Open "SELECT MARGIN FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn
rs9.Open "SELECT MARGIN_PESO FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn

rs10.Open "SELECT STOCK_ON_HAND FROM stock_info where DESCRIPTION like " & "'" & UCase(Text1.Text) & "%'", conn

rs10.MoveFirst
rs1.MoveFirst
rs2.MoveFirst
rs3.MoveFirst
rs4.MoveFirst
rs5.MoveFirst
rs6.MoveFirst
rs7.MoveFirst
rs9.MoveFirst

End If

If Combo1.Text = "Category" Then
rs1.Open "SELECT CATEGORY FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs3.Open "SELECT MODEL FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs6.Open "SELECT RP FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs7.Open "SELECT CP FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs8.Open "SELECT MARGIN FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs9.Open "SELECT MARGIN_PESO FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.Open "SELECT STOCK_ON_HAND FROM stock_info where CATEGORY like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.MoveFirst



rs8.MoveFirst
rs1.MoveFirst
rs2.MoveFirst
rs3.MoveFirst
rs4.MoveFirst
rs5.MoveFirst
rs6.MoveFirst
rs7.MoveFirst
rs9.MoveFirst

End If

If Combo1.Text = "Item_Code" Then
rs1.Open "SELECT CATEGORY FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs3.Open "SELECT MODEL FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs6.Open "SELECT RP FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs7.Open "SELECT CP FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs8.Open "SELECT MARGIN FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs9.Open "SELECT MARGIN_PESO FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.Open "SELECT STOCK_ON_HAND FROM stock_info where ITEM_CODE like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.MoveFirst

rs8.MoveFirst


rs1.MoveFirst
rs2.MoveFirst
rs3.MoveFirst
rs4.MoveFirst
rs5.MoveFirst
rs6.MoveFirst
rs7.MoveFirst
rs9.MoveFirst
End If

If Combo1.Text = "Model/Size" Then
rs1.Open "SELECT CATEGORY FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs3.Open "SELECT MODEL FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs6.Open "SELECT RP FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs7.Open "SELECT CP FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs8.Open "SELECT MARGIN FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs9.Open "SELECT MARGIN_PESO FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.Open "SELECT STOCK_ON_HAND FROM stock_info where MODEL like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.MoveFirst

rs8.MoveFirst
rs1.MoveFirst
rs2.MoveFirst
rs3.MoveFirst
rs4.MoveFirst
rs5.MoveFirst
rs6.MoveFirst
rs7.MoveFirst
rs9.MoveFirst
End If

If Combo1.Text = "Supplier Name " Then
rs1.Open "SELECT CATEGORY FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs3.Open "SELECT MODEL FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs6.Open "SELECT RP FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs7.Open "SELECT CP FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs8.Open "SELECT MARGIN FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs9.Open "SELECT MARGIN_PESO FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.Open "SELECT STOCK_ON_HAND FROM stock_info where SUPPLIER_NAME like " & "'" & UCase(Text1.Text) & "%'", conn
rs10.MoveFirst

rs8.MoveFirst
rs1.MoveFirst
rs2.MoveFirst
rs3.MoveFirst
rs4.MoveFirst
rs5.MoveFirst
rs6.MoveFirst
rs9.MoveFirst
rs7.MoveFirst

End If

Dim ab As Integer
Dim ba As Integer


b = 0
Do Until rs4.EOF
b = b + 1
grid.Rows = b

For Each fld1 In rs1.Fields
grid.TextMatrix(b - 1, 1) = fld1.Value
Next

For Each fld1 In rs8.Fields
grid.TextMatrix(b - 1, 11) = fld1.Value
Next

For Each fld1 In rs10.Fields
grid.TextMatrix(b - 1, 10) = fld1.Value
Next


For Each fld1 In rs2.Fields
grid.TextMatrix(b - 1, 2) = fld1.Value
Next

For Each fld1 In rs3.Fields
grid.TextMatrix(b - 1, 3) = fld1.Value
Next

For Each fld1 In rs3.Fields
grid.TextMatrix(b - 1, 3) = fld1.Value
Next

For Each fld1 In rs4.Fields
grid.TextMatrix(b - 1, 4) = fld1.Value
Next

For Each fld1 In rs5.Fields
grid.TextMatrix(b - 1, 8) = fld1.Value
Next



For Each fld1 In rs6.Fields
grid.TextMatrix(b - 1, 6) = fld1.Value
Next

For Each fld1 In rs7.Fields
grid.TextMatrix(b - 1, 5) = fld1.Value
Next

For Each fld1 In rs9.Fields
grid.TextMatrix(b - 1, 7) = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
rs6.MoveNext
rs7.MoveNext
rs8.MoveNext
rs9.MoveNext
rs10.MoveNext
Loop
grid.Sort = 1
outme:
grid.Sort = 1
'MsgBox (Err.Description)
Exit Sub

End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
main.Enabled = True
main.Show
Exit Sub
End If

If KeyCode = 38 Or 40 Then
grid.SetFocus
End If

If KeyCode = 13 Then
Text4.Text = grid.Row
Text2.Text = grid.TextMatrix(grid.Row, 4)
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
rs1.Open "SELECT RP FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn
rs8.Open "SELECT STOCK_ON_HAND FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn

Do Until rs1.EOF
For Each fld1 In rs1.Fields
item_desc.Text5 = fld1.Value
Next

For Each fld1 In rs8.Fields
item_desc.Text3 = fld1.Value
Next
rs8.MoveNext
rs1.MoveNext
Loop
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Text1.SetFocus
End Sub

Private Sub trabaho()

item_desc.Show
'Me.Enabled = False
item_desc.Text1 = grid.TextMatrix(grid.Row, 2) & ",  " & grid.TextMatrix(grid.Row, 3)
item_desc.Text2 = grid.TextMatrix(grid.Row, 4)
item_desc.Text10 = grid.TextMatrix(grid.Row, 8)

Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
rs1.Open "SELECT RP FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn
rs8.Open "SELECT STOCK_ON_HAND FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn
rs2.Open "SELECT CP FROM stock_info where ITEM_CODE = '" & grid.TextMatrix(grid.Row, 4) & "'", conn

Dim aa() As String

Do Until rs1.EOF
For Each fld1 In rs1.Fields
item_desc.Text5 = fld1.Value
item_desc.Text5 = Val(item_desc.Text5) * 1.0000001
aa = Split(item_desc.Text5, ".")
item_desc.Text5.Text = Left(item_desc.Text5.Text, Len(aa(0)) + 3)
Next

For Each fld1 In rs2.Fields
item_desc.cp.Text = fld1.Value
Next

For Each fld1 In rs8.Fields
item_desc.Text3 = fld1.Value
If item_desc.Text3.Text = "" Or Val(item_desc.Text3.Text) = 0 Then
item_desc.Enabled = True
item_desc.Text3.Text = "Out of Stock"
item_desc.Text3.BackColor = &HC0&
item_desc.Text3.ForeColor = &HFFFF&
End If
Next
rs8.MoveNext
rs1.MoveNext
Loop


End Sub

