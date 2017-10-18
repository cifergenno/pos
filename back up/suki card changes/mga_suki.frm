VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mga_suki 
   Caption         =   "SUKI CARD FORM"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   8
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
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "Amount"
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
         Left            =   10080
         TabIndex        =   7
         Top             =   240
         Width           =   1095
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2415
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
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Control #"
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
         Left            =   6120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Suki Card"
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
         Left            =   6960
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Total Points"
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
         Left            =   8640
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Suki Card's Master List"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "mga_suki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public axx As Boolean
Public acc As Integer

Private Sub Command1_Click()
Unload Me
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
    
    ObjWs.Cells(1, 1) = "All SUKI Card Member as of: " & Now

    
    
    ObjWs.Cells(3, 1) = "NAME"
    ObjWs.Cells(3, 2) = "ADDRESS"
    ObjWs.Cells(3, 3) = "CONTACT NUMBER"
    ObjWs.Cells(3, 4) = "CONTROL NUMBER"
    ObjWs.Cells(3, 5) = "SUKI CARD NUMBER"
    ObjWs.Cells(3, 6) = "TOTAL POINTS"
    ObjWs.Cells(3, 7) = "AMOUNT"
     

    For aaa = 0 To grid2.Rows
    On Error GoTo patay_na
    
    ObjWs.Cells(4 + aaa, 1) = grid2.TextMatrix(aaa + 1, 1)
    ObjWs.Cells(4 + aaa, 2) = grid2.TextMatrix(aaa + 1, 2)
    ObjWs.Cells(4 + aaa, 3) = grid2.TextMatrix(aaa + 1, 7)
    ObjWs.Cells(4 + aaa, 4) = grid2.TextMatrix(aaa + 1, 3)
    ObjWs.Cells(4 + aaa, 5) = grid2.TextMatrix(aaa + 1, 4)
    ObjWs.Cells(4 + aaa, 6) = grid2.TextMatrix(aaa + 1, 5)
    ObjWs.Cells(4 + aaa, 7) = grid2.TextMatrix(aaa + 1, 6)
    Next
    
    
patay_na:

     
    ObjWb.SaveAs ("d:\List of All Suki Card Holder.xls")
    ObjWb.Close (SaveChanges = False)
    MsgBox ("Saving finished." & vbNewLine & "Extracted list has been save to drive D:")

End Sub

Private Sub Form_Load()

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

grid2.ColAlignment(1) = 0
'grid2.ColAlignment(6) = 2
grid2.RowHeight(0) = 1
grid2.ColWidth(0) = 2
grid2.ColWidth(1) = 2450
grid2.ColWidth(2) = 3350
grid2.ColWidth(3) = 800
grid2.ColWidth(4) = 1700
grid2.ColWidth(5) = 1500
grid2.ColWidth(6) = 1700

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


rs1.Open "SELECT CUSTOMER_ID FROM customer", conn
rs2.Open "SELECT NAME FROM customer", conn
rs3.Open "SELECT ADDRESS FROM customer", conn
rs4.Open "SELECT CARD_NUMBER FROM customer", conn
rs6.Open "SELECT NUMBER FROM customer", conn

Dim ihap As Integer
ihap = 1

rs4.MoveFirst
Do Until rs4.EOF


For Each fld In rs4.Fields


If fld.Value <> "unregistered" Then
grid2.TextMatrix(ihap, 4) = fld.Value

Else
GoTo sunod
End If




If fld.Value <> "" Then
grid2.TextMatrix(ihap, 4) = fld.Value

Else
GoTo sunod
End If

Next


For Each fld In rs6.Fields
grid2.TextMatrix(ihap, 7) = fld.Value
Next

For Each fld In rs1.Fields
grid2.TextMatrix(ihap, 3) = fld.Value

Next


For Each fld In rs2.Fields
grid2.TextMatrix(ihap, 1) = fld.Value
Next


For Each fld In rs3.Fields
grid2.TextMatrix(ihap, 2) = fld.Value
Next

ihap = ihap + 1
grid2.Rows = ihap + 1
sunod:

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext

Loop
'grid2.Rows = ihap - 1
Dim ihapa As Integer
Dim aaa As String

ihapa = 0

For ihapa = 0 To ihap

aaa = grid2.TextMatrix(ihapa, 3)

rs5.Open "SELECT POINTS FROM suki_card  where CARD_NUMBER = '" & grid2.TextMatrix(ihapa, 3) & "'", conn

'MsgBox grid2.TextMatrix(ihapa, 3)


Do Until rs5.EOF

For Each fld In rs5.Fields
grid2.TextMatrix(ihapa, 5) = Val(grid2.TextMatrix(ihapa, 5)) + Val(fld.Value)

grid2.TextMatrix(ihapa, 6) = Val(grid2.TextMatrix(ihapa, 5)) * 0.3

grid2.TextMatrix(ihapa, 5) = Left(grid2.TextMatrix(ihapa, 5), 8)

Next





rs5.MoveNext
Loop




rs5.Close

If grid2.TextMatrix(ihapa, 6) = "" Then
grid2.TextMatrix(ihapa, 6) = "0.00" & "        "
grid2.TextMatrix(ihapa, 5) = "0.00"
End If




'unsa = Split(Text7.Text, ".")
'Text7.Text = "Php. " & Left(Text7.Text, (Len(unsa(0)) + 3))



Dim aa() As String
Dim aaa1() As String

aaa1 = Split(grid2.TextMatrix(ihapa, 5), ".")
grid2.TextMatrix(ihapa, 5) = Left(grid2.TextMatrix(ihapa, 5), Len(aaa1(0)) + 3)

aa = Split(grid2.TextMatrix(ihapa, 6), ".")
grid2.TextMatrix(ihapa, 6) = Left(grid2.TextMatrix(ihapa, 6), Len(aa(0)) + 3) & "        "
Next


grid2.Rows = grid2.Rows - 1



End Sub

Private Sub grid2_DblClick()
Text1.Text = grid2.TextMatrix(grid2.Row, 3)
'MsgBox Text1.Text
suki_ko.Show
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

grid2.Col = 3
grid2.Sort = acc
End Sub

Private Sub Label15_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = 6
grid2.Sort = acc
End Sub

Private Sub Label16_Click()
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

Private Sub Label17_Click()
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
