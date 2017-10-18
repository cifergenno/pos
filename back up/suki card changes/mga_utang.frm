VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mga_utang 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Credit Viewer"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List5 
      Height          =   450
      Left            =   2760
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Command3"
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
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   2520
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox dd 
      Height          =   285
      Left            =   9480
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox mm 
      Height          =   285
      Left            =   10440
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox yy 
      Height          =   375
      Left            =   11400
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Create Excel file"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "&Ok/Exit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   12255
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   7
         BackColorBkg    =   16777215
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Balance"
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
         Left            =   10200
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Control No."
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
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Address"
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
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Name"
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
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Contact No."
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
         Left            =   6960
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Credit List"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   12495
   End
End
Attribute VB_Name = "mga_utang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public axx As Boolean
Public acc As Integer

Private Sub Command1_Click()

Unload Me
history2.Show
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then

Unload Me
End If
End Sub

Private Sub Command2_Click()
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim AppXls As New Excel.Application
Dim ObjWb As Excel.Workbook
Dim ObjWs As Excel.Worksheet
Dim xx As Integer
Dim xx2 As Integer

xx = 1
xx2 = 3
Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.add
Set ObjWs = ObjWb.Worksheets.add
    
    ObjWs.Cells(1, 1) = "All Credit List as of: " & Now

    
    
    ObjWs.Cells(3, 1) = "NAME"
    ObjWs.Cells(3, 2) = "ADDRESS"
    ObjWs.Cells(3, 3) = "CONTACT NUMBER"
    ObjWs.Cells(3, 4) = "CONTROL NUMBER"
    ObjWs.Cells(3, 5) = "BALANCE"
   

    For aaa = 0 To grid2.Rows
    On Error GoTo patay_na
    
    ObjWs.Cells(4 + aaa, 1) = grid2.TextMatrix(aaa + 1, 1)
    ObjWs.Cells(4 + aaa, 2) = grid2.TextMatrix(aaa + 1, 2)
    ObjWs.Cells(4 + aaa, 3) = grid2.TextMatrix(aaa + 1, 3)
    ObjWs.Cells(4 + aaa, 4) = grid2.TextMatrix(aaa + 1, 4)
    ObjWs.Cells(4 + aaa, 5) = grid2.TextMatrix(aaa + 1, 5)
    
    Next
    
    
patay_na:

     
    ObjWb.SaveAs ("d:\Credit List.xls")
    ObjWb.Close (SaveChanges = False)
    MsgBox ("Saving finished." & vbNewLine & "Extracted list has been save to drive D:")

End Sub

Private Sub Command3_Click()
mga_utang_expire.text2.Text = List5.ListCount - 1
mga_utang_expire.Show

End Sub

Private Sub Form_Load()

Me.Show
text2.Text = Now
Dim bb() As String
Dim bbb() As String
bb = Split(text2.Text, " ")
bbb = Split(bb(0), "/")
mm.Text = bbb(0)
dd.Text = bbb(1)
yy.Text = bbb(2)

Dim adlaw As Integer
Dim buwan As Integer
Dim tuig As Integer


axx = True

Dim sulod_sa As String

Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs7 As ADODB.Recordset
Dim rs8 As ADODB.Recordset
Dim rs9 As ADODB.Recordset
Dim rs10 As ADODB.Recordset
Dim rs12 As ADODB.Recordset
Dim rs13 As ADODB.Recordset
Dim rs14 As ADODB.Recordset
Dim rs15 As ADODB.Recordset
Dim rs16 As ADODB.Recordset
Dim rs17 As ADODB.Recordset
Dim fld As ADODB.Field
Dim fld2 As ADODB.Field

grid2.ColAlignment(1) = 0
'grid2.ColAlignment(6) = 2
grid2.RowHeight(0) = 1
grid2.ColWidth(0) = 2
grid2.ColWidth(1) = 2900
grid2.ColWidth(2) = 3640
grid2.ColWidth(3) = 1390
grid2.ColWidth(4) = 1800
grid2.ColWidth(5) = 1700





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
Set rs11 = New ADODB.Recordset
Set rs12 = New ADODB.Recordset
Set rs13 = New ADODB.Recordset
Set rs14 = New ADODB.Recordset
Set rs15 = New ADODB.Recordset
Set rs16 = New ADODB.Recordset







rs7.Open "SELECT DUE_DATE FROM utang", conn
rs8.Open "SELECT CUSTOMER_ID FROM utang", conn

Do Until rs7.EOF

For Each fld In rs7.Fields

If Trim(fld.Value) = "" Then
GoTo sunod_n
Else
aaa = fld.Value
End If

aa = Split(aaa, "/")

If Val(dd.Text) > Val(aa(1)) Then
adlaw = Val(aa(1)) + 30
buwan = Val(aa(0)) - 1
adlaw = adlaw - Val(dd.Text)

Else
adlaw = Val(aa(1)) - Val(dd.Text)
buwan = Val(aa(0))
 End If
 



If Val(mm.Text) > buwan Then
buwan = buwan + 12
tuig = Val(aa(2)) - 1
buwan = buwan - Val(mm.Text)
Else
buwan = Val(aa(0)) - buwan
tuig = Val(aa(2))
End If


tuig = tuig - Val(yy.Text)


adlaw = adlaw + (buwan * 30) + (tuig * 12 * 30)

List3.AddItem adlaw & "   " & fld.Value

If adlaw <= 10 Then
    For Each fld1 In rs8.Fields
    List1.AddItem fld1.Value
    List4.AddItem adlaw
    Next

End If

Next
sunod_n:

rs7.MoveNext
rs8.MoveNext
Loop




Dim ihapp As Integer
ihapp = 0
List2.AddItem List1.List(0)
List5.AddItem List4.List(0)
  mga_utang_expire.List1.AddItem List1.List(0)
  mga_utang_expire.List2.AddItem List4.List(0)
For aa = 0 To List1.ListCount
    For bbc = 0 To List2.ListCount - 1
        If List1.List(aa) = List2.List(bbc) Then
        ihapp = 2
        End If
    Next
    If ihapp = 0 Then
    List2.AddItem List1.List(aa)
    List5.AddItem List4.List(aa)
    mga_utang_expire.List1.AddItem List1.List(aa)
    mga_utang_expire.List2.AddItem List4.List(aa)
    Else
    ihapp = 0
    End If
Next



Command3.Caption = List2.ListCount - 1 & " warnig(s)"

If (List2.ListCount - 1) = 0 Then
Command3.Enabled = False
Else
Command3.Enabled = True
End If






'display------------------

rs1.Open "SELECT CUSTOMER_ID FROM customer", conn
rs2.Open "SELECT NAME FROM customer", conn
rs3.Open "SELECT ADDRESS FROM customer", conn
rs4.Open "SELECT NUMBER FROM customer", conn
rs5.Open "SELECT CARD_NUMBER FROM customer", conn

Dim ihap As Integer
ihap = 1

Do Until rs1.EOF


grid2.Rows = ihap + 1

For Each fld In rs1.Fields



rs6.Open "SELECT BALANCE FROM utang WHERE CUSTOMER_ID = '" & fld.Value & "'", conn


If rs6.EOF = True Then
GoTo sunod_napud
End If
grid2.TextMatrix(ihap, 4) = fld.Value
Do Until rs6.EOF

For Each fld2 In rs6.Fields
grid2.TextMatrix(ihap, 5) = fld2.Value
Next

rs6.MoveNext
Loop


Next



For Each fld In rs2.Fields
grid2.TextMatrix(ihap, 1) = fld.Value
Next

For Each fld In rs3.Fields
grid2.TextMatrix(ihap, 2) = fld.Value
Next

For Each fld In rs4.Fields
grid2.TextMatrix(ihap, 3) = fld.Value
Next


ihap = ihap + 1

sunod_napud:

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs6.Close


Loop





End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
main.Enabled = True
main.item_code.SetFocus
Unload Me
End Sub

Private Sub grid2_DblClick()
Text1.Text = grid2.TextMatrix(grid2.Row, 4)
'MsgBox Text1.Text
customer_ko.Show
End Sub

Private Sub Label10_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = 2
grid2.Sort = acc
End Sub

Private Sub Label11_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = 4
grid2.Sort = acc
End Sub

Private Sub Label13_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = 3
grid2.Sort = acc
End Sub

Private Sub Label2_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = 5
grid2.Sort = acc
End Sub

Private Sub Label9_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = 1
grid2.Sort = acc

End Sub
