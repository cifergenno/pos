VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   13
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   12
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   11
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   9
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   8
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Text            =   "COLS"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Text            =   "book2"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   480
      TabIndex        =   0
      Text            =   "ROW"
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook

Private Sub Command1_Click()

Dim vall() As String
Dim rowss As Integer
Dim coll As Integer


Dim fld As ADODB.Field
Dim conn As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim cat As String
Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open
conn.Execute "truncate stock_info"

rowsa = 2
For rowss = 2 To Val(Text1.Text)

Text2(0).Text = rowsa
    Text2(1).Text = xlsheet.Cells(rowss, 1) ' row 2 col 1
   Text2(2).Text = xlsheet.Cells(rowss, 2)  ' row 2 col 1
   Text2(3).Text = xlsheet.Cells(rowss, 3)   ' row 2 col 1
    Text2(4).Text = xlsheet.Cells(rowss, 4)   ' row 2 col 1
    Text2(5).Text = xlsheet.Cells(rowss, 5)   ' row 2 col 1
   Text2(6).Text = xlsheet.Cells(rowss, 6)  ' row 2 col 1
   Text2(7).Text = xlsheet.Cells(rowss, 7)  ' row 2 col 1
   Text2(8).Text = xlsheet.Cells(rowss, 8)  ' row 2 col 1
    Text2(9).Text = xlsheet.Cells(rowss, 9)  ' row 2 col 1
   Text2(10).Text = xlsheet.Cells(rowss, 10)  ' row 2 col 1
    Text2(11).Text = xlsheet.Cells(rowss, 11)  ' row 2 col 1
              
   


conn.Execute "INSERT INTO stock_info (CATEGORY, MODEL,DESCRIPTION,ITEM_CODE, DATE_RECEIVED, SUPPLIER_NAME, CP, RP,MARGIN_PESO,MARGIN,STOCK_ON_HAND)" _
& "values ('" & Text2(1).Text & "', '" & Text2(2).Text & "', '" & Text2(3).Text & "', '" & Text2(4).Text & "', '" & Text2(5).Text & "', '" & Text2(6).Text & "', '" & Text2(7).Text & "', '" & Text2(8).Text & "', '" & Text2(9).Text & "', '" & Text2(10).Text & "', '" & Text2(11).Text & "')"

rowsa = rowsa + 1
Next



    xl.ActiveWorkbook.Close False, "c:\book1.xls"
    xl.Quit
    
     Set xlwbook = Nothing
    Set xl = Nothing
End Sub



Private Sub Form_Load()
    Set xlwbook = xl.Workbooks.Open("c:\book1.xls")
    Set xlsheet = xlwbook.Sheets.Item(1)
End Sub

