VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form extract 
   Caption         =   "Report Extractor"
   ClientHeight    =   3900
   ClientLeft      =   7395
   ClientTop       =   4380
   ClientWidth     =   4635
   Icon            =   "extract me.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "extract me.frx":1C64C
   ScaleHeight     =   3900
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   7200
      TabIndex        =   16
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7920
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   4560
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox yy 
      Height          =   495
      Left            =   7800
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox dd 
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox mm 
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   360
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      Format          =   37683200
      CurrentDate     =   40591
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      Format          =   37683200
      CurrentDate     =   40591
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5520
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Extract Sales Report"
      Default         =   -1  'True
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Extrac excel file to drive D:"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4095
   End
End
Attribute VB_Name = "extract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()




Label2.Caption = "Please wait a moment!!"
If Val(Text4.Text) = 0 Then
Exit Sub
End If
extract_sales
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 End If
End Sub


Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 End If
End Sub

Private Sub dd_GotFocus()
dd.Text = ""
End Sub

Private Sub dd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 End If
 
 If Len(dd.Text) >= 2 Then
yy.SetFocus
End If
End Sub

Private Sub DTPicker2_Change()
Text2.Text = DTPicker2.Value
End Sub

Private Sub DTPicker3_Change()
Text3.Text = DTPicker3.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub extract_sales()

MsgBox "Extract report from " & Text2.Text & " to " & Text3.Text
Label2.Caption = "Please wait a moment!!"
Label4.Caption = "Gathering all items!"

Dim aa() As String
Dim texto As String
Dim adlaw As Integer
Dim adlaw2 As Integer
Dim buwan As Integer
Dim tuig As Integer
Dim texto1 As String
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs7 As ADODB.Recordset
Dim rs8 As ADODB.Recordset
Dim rs9 As ADODB.Recordset

Dim fld As ADODB.Field
Dim conn As ADODB.Connection
Dim aab(20) As String
Dim i1 As Integer
Dim i2 As Integer
Dim ihap As Integer
Dim AppXls As New Excel.Application
Dim ObjWb As Excel.Workbook
Dim ObjWs As Excel.Worksheet
Dim xx As Integer
Dim xx2 As Integer

ihap = 0
i1 = 0
i2 = 0
xx = 1
xx2 = 2

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset
Set rs7 = New ADODB.Recordset
Set rs8 = New ADODB.Recordset
Set rs9 = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

conn.Execute "truncate sales_extrac"
conn.Execute "truncate payment"

rs1.Open "SELECT * FROM stock_info", conn
 
Dim xx1() As String


texto = Text2.Text
texto1 = Text3.Text

 
Do Until rs1.EOF
i1 = 0
    For Each fld In rs1.Fields
        aab(i1) = fld.Value
        i1 = i1 + 1
    Next
 
    i2 = i2 + 1
 
    Label4.Caption = "Gathering all items! Item count: " & i2

'----------------AGEING-----------------

adlaw = 0
buwan = 0
tuig = 0
 
If aab(4) = "" Then
aaa = "0/0/0"
Else
aaa = aab(4)
End If

aa = Split(aaa, "/")


If Val(dd.Text) < Val(aa(1)) Then
adlaw = Val(dd.Text) + 30
buwan = Val(mm.Text) - 1
adlaw = (adlaw - Val(aa(1)))
Else
adlaw = Val(dd.Text) - Val(aa(1))
buwan = Val(mm.Text)
End If
 
If Val(aa(0)) > buwan Then
buwan = buwan + 12
tuig = Val(yy.Text) - 1
buwan = (buwan - Val(aa(0)))
Else
buwan = buwan - Val(aa(0))
tuig = Val(aa(2))
End If

tuig = tuig - Val(aa(2))

adlaw = adlaw + (buwan * 31) + (tuig * 12 * 30) + 1

'--------- AGEING end------------
 

conn.Execute "INSERT INTO sales_extrac (DATE_RECEIVED, CATEGORY, MODEL,DESCRIPTION,ITEM_CODE, SUPPLIER_NAME, CP, RP,MARGIN_PESO,MARGIN, STOCK_ON_HAND, AGEING)" _
& "values ('" & aab(4) & "', '" & aab(0) & "', '" & aab(2) & "', '" & aab(1) & "', '" & aab(3) & "', '" & aab(5) & "', '" & aab(6) & "', '" & aab(7) & "', '" & aab(8) & "', '" & aab(9) & "%" & "', '" & aab(10) & "', '" & adlaw & "')"
    rs1.MoveNext
Loop

'conn.Execute "INSERT INTO sales_extrac (DATE_RECEIVED) values('      ')"

'conn.Execute "INSERT INTO sales_extrac (DATE_RECEIVED, DATE_SOLD, CATEGORY, MODEL,DESCRIPTION,ITEM_CODE, SUPPLIER_NAME, CP, RP,MARGIN_PESO,MARGIN,  AGEING)" _
'& "values ('DATE_SOLD', 'ITEM_CODE', 'DESCRIPTION', 'DISCOUNT', 'GROSS', 'CASHER', 'INVOICE', ' ', ' ', ' ', ' ', ' ')"


'----- inserting value ends -----------


 Dim counterx As Boolean
 counterx = False
 
 Dim aac(20) As String
 Dim stock(20) As String
 Dim count_me As Integer
 count_me = 0
 
'-----------------other value start-------------------


xx1 = Split(Text2.Text, "/")
 
mm.Text = xx1(0)
dd.Text = xx1(1)
yy.Text = xx1(2)


For aaaa = 1 To Val(Text4.Text)
List1.AddItem aaaa
 
rs2.Open "SELECT * FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn

Label2.Caption = "Creating excel file..."

Do Until rs2.EOF
    i11 = 1
    
    For Each fld In rs2.Fields
        aac(i11) = fld.Value
        i11 = i11 + 1
    Next
    
    'MsgBox aac(2) & " " & aac(1)
   
    rs3.Open "SELECT * FROM sales_extrac where ITEM_CODE like '" & aac(5) & "'", conn
    count_me = 0
    
    
     If rs3.EOF Then
    
    If aac(13) = "--Suki Card--" Then
    Else
    conn.Execute "INSERT INTO payment (NET,DISCOUNT,DATE_SOLD, ITEM_CODE, DESCRIPTION, QTY, AMOUNT,CASHER, INVOICE)" _
    & "values ('" & aac(13) & "','" & aac(12) & "','" & aac(1) & "', '" & aac(5) & "', '" & aac(4) & "', '" & aac(6) & "', '" & aac(11) & "', '" & aac(14) & "', '" & aac(15) & "')"
    End If
    
    End If
    
    Do Until rs3.EOF
    i11 = 1
    
    For Each fld In rs3.Fields
        stock(i11) = fld.Value
        i11 = i11 + 1
    
    Next
    
    If stock(12) = "" Then
    stock(12) = 0
    End If
    
    If stock(13) = "" Then
    stock(12) = 0
    End If
    
    If stock(14) = "" Then
    stock(12) = 0
    End If
    
    If stock(15) = "" Then
    stock(12) = 0
    End If
    
    conn.Execute "UPDATE sales_extrac SET DATE_SOLD = '" & aac(1) & "' where ITEM_CODE = '" & aac(5) & "'"
    conn.Execute "UPDATE sales_extrac SET CASHER = '" & aac(14) & "' where ITEM_CODE = '" & aac(5) & "'"
    conn.Execute "UPDATE sales_extrac SET INVOICE = '" & aac(15) & "' where ITEM_CODE = '" & aac(5) & "'"
 
    conn.Execute "UPDATE sales_extrac SET QUANTITY = '" & Val(stock(12)) + Val(aac(6)) & "' where ITEM_CODE = '" & aac(5) & "'"
    conn.Execute "UPDATE sales_extrac SET GROSS_SALES = '" & Val(stock(13)) + Val(aac(11)) & "' where ITEM_CODE = '" & aac(5) & "'"
    conn.Execute "UPDATE sales_extrac SET DISCOUNT = '" & Val(stock(14)) + Val(aac(12)) & "' where ITEM_CODE = '" & aac(5) & "'"
    conn.Execute "UPDATE sales_extrac SET NET_SALES = '" & Val(stock(15)) + Val(aac(13)) & "' where ITEM_CODE = '" & aac(5) & "'"
    
    
    rs3.MoveNext
    count_me = count_me + 1
    Loop
    
    rs3.Close
    
    
    

rs2.MoveNext

Loop

 
mm = Val(mm.Text)
dd = Val(dd.Text)
yy = Val(yy.Text)
dd = dd + 1

If dd = 32 Then
mm = mm + 1
dd = 1
End If

If mm = 13 Then
yy = yy + 10
mm = 1
End If

mm.Text = mm
yy.Text = yy
dd.Text = dd

rs2.Close

Next




Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add
Set ObjWs = ObjWb.Worksheets.Add


    
    ObjWs.Cells(1, 1) = "DATE_RECEIVED"
    ObjWs.Cells(1, 2) = "DATE_SOLD"
    ObjWs.Cells(1, 3) = "CATEGORY"
    ObjWs.Cells(1, 4) = "MODEL"
    ObjWs.Cells(1, 5) = "DESCRIPTION"
    ObjWs.Cells(1, 6) = "ITEM CODE"
    ObjWs.Cells(1, 7) = "SUPPLIER_NAME"
    ObjWs.Cells(1, 8) = "CP"
    ObjWs.Cells(1, 9) = "RP"
    ObjWs.Cells(1, 10) = "MARGIN_PESO"
    ObjWs.Cells(1, 11) = "MARGIN"
    ObjWs.Cells(1, 12) = "QTY SOLD"
    ObjWs.Cells(1, 13) = "GROSS_SALES"
    ObjWs.Cells(1, 15) = "NET_SALES"
    ObjWs.Cells(1, 14) = "DISCOUNT"
    ObjWs.Cells(1, 16) = "STOCK ON HAND"
    ObjWs.Cells(1, 17) = "AGEING"
    ObjWs.Cells(1, 18) = "CASHER"
    ObjWs.Cells(1, 19) = "INVOICE"
    


    rs3.Open "SELECT * FROM sales_extrac", conn
    Dim count As Integer
    count = 2
    Do Until rs3.EOF
    i11 = 1
    
    For Each fld In rs3.Fields
        stock(i11) = fld.Value
         i11 = i11 + 1
    Next
    
    rs3.MoveNext
    
    ObjWs.Cells(count, 1) = stock(1)
    ObjWs.Cells(count, 2) = stock(2)
    ObjWs.Cells(count, 3) = stock(3)
    ObjWs.Cells(count, 4) = stock(4)
    ObjWs.Cells(count, 5) = stock(5)
    ObjWs.Cells(count, 6) = stock(6)
    ObjWs.Cells(count, 7) = stock(7)
    ObjWs.Cells(count, 8) = stock(8)
    ObjWs.Cells(count, 9) = stock(9)
    ObjWs.Cells(count, 10) = stock(10)
    ObjWs.Cells(count, 11) = stock(11)
    ObjWs.Cells(count, 12) = stock(12)
    ObjWs.Cells(count, 13) = stock(15)
    ObjWs.Cells(count, 15) = stock(13)
    ObjWs.Cells(count, 14) = stock(14)
    ObjWs.Cells(count, 16) = stock(16)
    ObjWs.Cells(count, 17) = stock(17)
    ObjWs.Cells(count, 18) = stock(18)
    ObjWs.Cells(count, 19) = stock(19)
    count = count + 1
    
    Loop


DTPicker2.Format = dtpLongDate
DTPicker3.Format = dtpLongDate

    ObjWb.SaveAs ("d:\Sales Report on " & Text1.Text & " to " & Text5.Text & ".xls")

    ObjWb.Close (SaveChanges = False)
    
    
    
    Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add
Set ObjWs = ObjWb.Worksheets.Add
    
        count = 1

    ObjWs.Cells(count, 1) = "Payment List and Unknown Item(s)"
    ObjWs.Cells(count + 1, 1) = "DATE_SOLD"
    ObjWs.Cells(count + 1, 2) = "ITEM_CODE"
    ObjWs.Cells(count + 1, 3) = "DESCRIPTION"
    ObjWs.Cells(count + 1, 4) = "QTY"
    ObjWs.Cells(count + 1, 5) = "GROSS_SALES"
    ObjWs.Cells(count + 1, 6) = "DISCOUNT/SUKI CARD AMOUNT"
    ObjWs.Cells(count + 1, 7) = "NET_SALES"
    ObjWs.Cells(count + 1, 8) = "CASHER"
    ObjWs.Cells(count + 1, 9) = "INVOICE"
    
    
    
      rs4.Open "SELECT * FROM payment", conn
    count = 3
    Do Until rs4.EOF
    
     i11 = 1
    
    
    For Each fld In rs4.Fields
        stock(i11) = fld.Value
         i11 = i11 + 1
    Next
    
    rs4.MoveNext
    
    ObjWs.Cells(count, 1) = stock(1)
    ObjWs.Cells(count, 2) = stock(2)
    ObjWs.Cells(count, 3) = stock(3)
    ObjWs.Cells(count, 4) = stock(4)
    ObjWs.Cells(count, 5) = stock(7)
    ObjWs.Cells(count, 6) = stock(6)
    ObjWs.Cells(count, 7) = stock(5)
    ObjWs.Cells(count, 8) = stock(8)
    ObjWs.Cells(count, 9) = stock(9)
    count = count + 1
    
    Loop
    
    
    ObjWb.SaveAs ("d:\Payment and Unknown Sales on " & Text1.Text & " to " & Text5.Text & ".xls")

    ObjWb.Close (SaveChanges = False)
    
    
    Me.Enabled = True

sibat:
If Err.Description = "" Then
MsgBox ("Saving finished!!")
Else
MsgBox Err.Description
End If
Unload Me
Exit Sub
    
End Sub


Private Sub Form_Load()
DTPicker2.Value = Format(Now, "d  mm, yyyy")
DTPicker3.Value = Format(Now, "d  mm, yyyy")
Text2.Text = DTPicker2.Value
Text3.Text = DTPicker3.Value
End Sub

Private Sub mm_Click()
mm.Text = ""
End Sub

Private Sub mm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
main.Enabled = True
main.Show
End If
End Sub

Private Sub Text2_Change()




On Error GoTo ss
Dim aa11() As String
Dim bb() As String
aa = Split(Text2.Text, "/")


If aa(0) = "1" Then
Text1.Text = "January " & aa(1) & ", " & aa(2)
End If

If aa(0) = "2" Then
Text1.Text = "February " & aa(1) & ", " & aa(2)
End If
If aa(0) = "3" Then
Text1.Text = "March " & aa(1) & ", " & aa(2)
End If
If aa(0) = "4" Then
Text1.Text = "April " & aa(1) & ", " & aa(2)
End If
If aa(0) = "5" Then
Text1.Text = "May " & aa(1) & ", " & aa(2)
End If
If aa(0) = "6" Then
Text1.Text = "June " & aa(1) & ", " & aa(2)
End If
If aa(0) = "7" Then
Text1.Text = "July " & aa(1) & ", " & aa(2)
End If
If aa(0) = "8" Then
Text1.Text = "August " & aa(1) & ", " & aa(2)
End If
If aa(0) = "9" Then
Text1.Text = "September " & aa(1) & ", " & aa(2)
End If
If aa(0) = "10" Then
Text1.Text = "October " & aa(1) & ", " & aa(2)
End If
If aa(0) = "11" Then
Text1.Text = "November " & aa(1) & ", " & aa(2)
End If
If aa(0) = "12" Then
Text1.Text = "December " & aa(1) & ", " & aa(2)
End If






bb = Split(Text3.Text, "/")

If Val(bb(1)) < Val(aa(1)) Then
adlaw = Val(bb(1)) + 31
buwan = Val(bb(0)) - 1
adlaw = (adlaw - Val(aa(1)))
Else
adlaw = Val(bb(1)) - Val(aa(1))
buwan = Val(bb(0))
End If
 
If Val(aa(0)) > buwan Then
buwan = buwan + 12
tuig = Val(bb(2)) - 1
buwan = (buwan - Val(aa(0)))
Else
buwan = buwan - Val(aa(0))
tuig = Val(aa(2))
End If

tuig = tuig - Val(aa(2))
adlaw = adlaw + (buwan * 31) + (tuig * 12 * 31)
Text4.Text = adlaw + 1


ss:

End Sub

Private Sub Text3_Change()




On Error GoTo ss
Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")
dd.Text = bb(1)
mm.Text = bb(0)
yy.Text = bb(2)




If bb(0) = "1" Then
Text5.Text = "January " & bb(1) & ", " & bb(2)
End If

If bb(0) = "2" Then
Text5.Text = "February " & bb(1) & ", " & bb(2)
End If
If bb(0) = "3" Then
Text5.Text = "March " & bb(1) & ", " & bb(2)
End If
If bb(0) = "4" Then
Text5.Text = "April " & bb(1) & ", " & bb(2)
End If
If bb(0) = "5" Then
Text5.Text = "May " & bb(1) & ", " & bb(2)
End If
If bb(0) = "6" Then
Text5.Text = "June " & bb(1) & ", " & bb(2)
End If
If bb(0) = "7" Then
Text5.Text = "July " & bb(1) & ", " & bb(2)
End If
If bb(0) = "8" Then
Text5.Text = "August " & bb(1) & ", " & bb(2)
End If
If bb(0) = "9" Then
Text5.Text = "September " & bb(1) & ", " & bb(2)
End If
If bb(0) = "10" Then
Text5.Text = "October " & bb(1) & ", " & bb(2)
End If
If bb(0) = "11" Then
Text5.Text = "November " & bb(1) & ", " & bb(2)
End If
If bb(0) = "12" Then
Text5.Text = "December " & bb(1) & ", " & bb(2)
End If



If Val(bb(1)) < Val(aa(1)) Then
adlaw = Val(bb(1)) + 30
buwan = Val(bb(0)) - 1
adlaw = (adlaw - Val(aa(1)))
Else
adlaw = Val(bb(1)) - Val(aa(1))
buwan = Val(bb(0))
End If
 
If Val(aa(0)) > buwan Then
buwan = buwan + 12
tuig = Val(bb(2)) - 1
buwan = (buwan - Val(aa(0)))
Else
buwan = buwan - Val(aa(0))
tuig = Val(aa(2))
End If

tuig = tuig - Val(aa(2))
adlaw = adlaw + (buwan * 30) + (tuig * 12 * 30)
Text4.Text = adlaw + 1
ss:
End Sub

Private Sub yy_GotFocus()
yy.Text = ""
End Sub

Private Sub yy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 main.Enabled = True
 main.Show
 End If
End Sub
