VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form customer5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   25
      Text            =   "Text5"
      Top             =   6720
      Visible         =   0   'False
      Width           =   150
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   480
      Width           =   3015
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "100000"
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Frame Frame4 
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   14295
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   2535
         Left            =   360
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   840
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   10
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
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Price"
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
         Left            =   6000
         TabIndex        =   24
         Top             =   240
         Width           =   1095
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
         Left            =   11280
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Due Date"
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
         Left            =   12480
         TabIndex        =   22
         Top             =   240
         Width           =   1455
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
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Descrition"
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
         TabIndex        =   14
         Top             =   240
         Width           =   3615
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
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Debit"
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
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Credit"
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
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Balance"
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
         Left            =   10080
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   2040
      Width           =   3015
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   1440
      Width           =   3015
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   2040
      Width           =   4575
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
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1440
      Width           =   4575
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
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
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
      Left            =   9480
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
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
      Left            =   9120
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
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
      TabIndex        =   4
      Top             =   2160
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   855
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
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "customer5"
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

Private Sub Form_Load()

axx = True
Dim ccc As String
Text8.Text = history2.Text3.Text
On Error GoTo bye
'Text8.Text = transac.Text12.Text
grid2.ColAlignment(1) = 0
'grid2.ColAlignment(6) = 2
grid2.RowHeight(0) = 1
grid2.ColWidth(0) = 2
grid2.ColWidth(1) = 1800
grid2.ColWidth(2) = 3730
grid2.ColWidth(3) = 1210
grid2.ColWidth(4) = 580
grid2.ColWidth(5) = 1080
grid2.ColWidth(6) = 1200
grid2.ColWidth(7) = 1200
grid2.ColWidth(8) = 1200
grid2.ColWidth(9) = 1800

grid2.Col = 8
grid2.Sort = 0

'GoTo bye
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

'GoTo bye
rs5.Open "SELECT CUSTOMER_ID FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs6.Open "SELECT ITEM_CODE FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn 'rs7.Open "SELECT BAL FROM customer WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs8.Open "SELECT DESCRIPTION FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs9.Open "SELECT QUANTITY FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs10.Open "SELECT RP FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs11.Open "SELECT DEBIT FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs12.Open "SELECT CREDIT FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs13.Open "SELECT BALANCE FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs14.Open "SELECT DATE_SOLD FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn
rs15.Open "SELECT DUE_DATE FROM utang WHERE CUSTOMER_ID = '" & Text8.Text & "'", conn


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

Dim ihap As Integer
ihap = 1
Dim vall() As String


Do Until rs5.EOF
'MsgBox ihap


For Each fld In rs6.Fields
grid2.TextMatrix(ihap, 1) = fld.Value
Next

For Each fld In rs8.Fields
grid2.TextMatrix(ihap, 2) = fld.Value
Next

For Each fld In rs9.Fields
grid2.TextMatrix(ihap, 4) = fld.Value
Next

For Each fld In rs10.Fields
grid2.TextMatrix(ihap, 3) = fld.Value
Next

For Each fld In rs11.Fields
grid2.TextMatrix(ihap, 5) = fld.Value
Next

For Each fld In rs12.Fields
grid2.TextMatrix(ihap, 6) = Trim(fld.Value)
Next

For Each fld In rs13.Fields



Text5.Text = Val(fld.Value) * 1.00000001
aaa = Split(Text5.Text, ".")
Text5.Text = Left(Text5.Text, Len(aaa(0)) + 3)

grid2.TextMatrix(ihap, 7) = Text5.Text
ccc = fld.Value
Next

For Each fld In rs14.Fields
vall = Split(Trim(fld.Value), " ")
grid2.TextMatrix(ihap, 8) = vall(0)
Next


For Each fld In rs15.Fields
grid2.TextMatrix(ihap, 9) = fld.Value & "       "
Next
rs6.MoveNext
rs5.MoveNext
rs8.MoveNext
rs9.MoveNext
rs10.MoveNext
rs11.MoveNext
rs12.MoveNext
rs13.MoveNext
rs14.MoveNext
rs15.MoveNext
ihap = ihap + 1
grid2.Rows = ihap + 1
Loop

'MsgBox ihap

grid2.Rows = ihap
Text7.Text = "Php. " & Text5.Text
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

