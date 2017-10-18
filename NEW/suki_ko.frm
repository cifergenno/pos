VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form suki_ko 
   Caption         =   "SUKI CARD FORM"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   12180
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Create Excel File"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   " 0.00"
      Top             =   6240
      Width           =   2655
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1335
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4095
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4095
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
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text4 
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
      Height          =   405
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
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
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   2535
         Left            =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   7
         BackColorBkg    =   16777215
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   2
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
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Amount"
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
         Left            =   8880
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Price"
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
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Qty"
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
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
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
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Date Sold"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount       Value"
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
      Left            =   3960
      TabIndex        =   23
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Points"
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
      Left            =   360
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Customer's profile"
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
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   6855
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
      Left            =   240
      TabIndex        =   17
      Top             =   1320
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
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   6720
      TabIndex        =   15
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Control #"
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
      Left            =   7440
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "suki_ko"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Public fld As ADODB.Field
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
Public rs11 As ADODB.Recordset
Public rs12 As ADODB.Recordset
Public rs13 As ADODB.Recordset
Public rs14 As ADODB.Recordset
Public rs15 As ADODB.Recordset
Public rs16 As ADODB.Recordset
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
    
    ObjWs.Cells(1, 1) = "Name:"
    ObjWs.Cells(1, 2) = Text1.Text
    ObjWs.Cells(2, 1) = "Customer Control Number:"
    ObjWs.Cells(2, 2) = Text8.Text
    ObjWs.Cells(3, 1) = "SUKI Card Number:"
    ObjWs.Cells(3, 2) = Text3.Text
    ObjWs.Cells(4, 1) = "Address:"
    ObjWs.Cells(4, 2) = Text2.Text
    ObjWs.Cells(5, 1) = "Contact Number:"
    ObjWs.Cells(5, 2) = Text4.Text
    
    
    
    ObjWs.Cells(8, 1) = "Item Code"
    ObjWs.Cells(8, 2) = "Description"
    ObjWs.Cells(8, 3) = "Quantity"
    ObjWs.Cells(8, 4) = "Price"
    ObjWs.Cells(8, 5) = "Amount"
    ObjWs.Cells(8, 6) = "Date Sold"


    For aaa = 0 To grid2.Rows
    On Error GoTo patay_na
    
    ObjWs.Cells(9 + aaa, 1) = grid2.TextMatrix(aaa + 1, 1)
    ObjWs.Cells(9 + aaa, 2) = grid2.TextMatrix(aaa + 1, 2)
    ObjWs.Cells(9 + aaa, 3) = grid2.TextMatrix(aaa + 1, 3)
    ObjWs.Cells(9 + aaa, 4) = grid2.TextMatrix(aaa + 1, 4)
    ObjWs.Cells(9 + aaa, 5) = grid2.TextMatrix(aaa + 1, 5)
    ObjWs.Cells(9 + aaa, 6) = grid2.TextMatrix(aaa + 1, 6)
    
    Next
    
    
patay_na:

ObjWs.Cells(9 + aaa + 2, 1) = "Total Points:"
ObjWs.Cells(9 + aaa + 2, 2) = Text7.Text

ObjWs.Cells(9 + aaa + 3, 1) = "Total Points:"
ObjWs.Cells(9 + aaa + 3, 2) = Text7.Text

     
    ObjWb.SaveAs ("d:\" & Text1.Text & "'s SUKI Card Print Out.xls")
    ObjWb.Close (SaveChanges = False)
 MsgBox ("Saving finished." & vbNewLine & "Extracted list has been save to drive D:")

End Sub

Private Sub Form_Load()

Text8.Text = mga_suki.Text1.Text

axx = True

On Error GoTo bye
'Text8.Text = transac.Text12.Text
grid2.ColAlignment(1) = 0
grid2.ColAlignment(2) = 0
'grid2.ColAlignment(6) = 1
grid2.RowHeight(0) = 1
grid2.ColWidth(0) = 2
grid2.ColWidth(1) = 1800
grid2.ColWidth(2) = 4980
grid2.ColWidth(3) = 550
grid2.ColWidth(4) = 1080
grid2.ColWidth(5) = 1200
grid2.ColWidth(6) = 1500
'grid2.ColWidth(7) = 1200
'grid2.ColWidth(8) = 2000

grid2.Col = grid2.Col
grid2.Sort = 1


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



rs1.Open "SELECT CARD_NUMBER FROM customer WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs2.Open "SELECT NAME FROM customer WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs3.Open "SELECT ADDRESS FROM customer WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs4.Open "SELECT NUMBER FROM customer WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn


For Each fld In rs1.Fields
Text3.Text = fld.Value
Next

For Each fld In rs2.Fields
Text1.Text = fld.Value
Next

For Each fld In rs3.Fields
Text2.Text = fld.Value
Next

For Each fld In rs4.Fields
Text4.Text = fld.Value
Next



rs5.Open "SELECT ITEM_CODE FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn
rs6.Open "SELECT DESCRIPTION FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn
rs7.Open "SELECT QUANTITY FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn
rs8.Open "SELECT PRICE FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn
rs9.Open "SELECT AMOUNT FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn
rs10.Open "SELECT DATE_SOLD FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn
rs11.Open "SELECT POINTS FROM suki_card WHERE CARD_NUMBER = '" & Text8.Text & "'", conn


Dim cccx() As String



Dim ihap As Integer

ihap = 1

Do Until rs5.EOF

'MsgBox ihap

For Each fld In rs9.Fields
If fld.Value = "" Then
GoTo sunod_na
End If
grid2.TextMatrix(ihap, 5) = Trim(fld.Value)
Next




For Each fld In rs5.Fields
grid2.TextMatrix(ihap, 1) = Trim(fld.Value)
Next

For Each fld In rs6.Fields
grid2.TextMatrix(ihap, 2) = Trim(fld.Value)
Next

For Each fld In rs7.Fields
grid2.TextMatrix(ihap, 3) = Trim(fld.Value)
Next

For Each fld In rs8.Fields
grid2.TextMatrix(ihap, 4) = Trim(fld.Value)
Next


For Each fld In rs10.Fields
cccx = Split(Trim(fld.Value), " ")
grid2.TextMatrix(ihap, 6) = cccx(0)
Next

ihap = ihap + 1
grid2.Rows = ihap + 1

sunod_na:


For Each fld In rs11.Fields

Text6.Text = Val(fld.Value) + Val(Text6.Text)
Text7.Text = Val(Text6.Text) * 0.3

Next

rs5.MoveNext
rs6.MoveNext
rs7.MoveNext
rs8.MoveNext
rs9.MoveNext
rs10.MoveNext
rs11.MoveNext


Loop
Dim unsa2() As String
Text6.Text = Text6.Text
unsa = Split(Text7.Text, ".")
Text7.Text = "Php. " & Left(Text7.Text, (Len(unsa(0)) + 3))

unsa2 = Split(Text6.Text, ".")
Text6.Text = Left(Text6.Text, (Len(unsa2(0)) + 3))

grid2.Rows = ihap
'Text6.Text = Val(Text5.Text) - Val(Text7.Text)
bye:
'MsgBox Err.Description
Exit Sub

End Sub

Private Sub grid2_Click()
If axx = True Then
acc = 1
axx = False
Else
acc = 2
axx = True
End If

grid2.Col = grid2.Col
grid2.Sort = acc

End Sub


