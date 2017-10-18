VERSION 5.00
Begin VB.Form extract 
   Caption         =   "Report Extractor"
   ClientHeight    =   3930
   ClientLeft      =   7395
   ClientTop       =   4380
   ClientWidth     =   4650
   Icon            =   "extract me.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "extract me.frx":1C64C
   ScaleHeight     =   3930
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Height          =   855
      Left            =   11880
      TabIndex        =   29
      Text            =   "Text11"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   855
      Left            =   12000
      TabIndex        =   28
      Text            =   "0"
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox List5 
      Height          =   3765
      Left            =   4440
      TabIndex        =   27
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   9600
      TabIndex        =   26
      Text            =   "0"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   735
      Left            =   9600
      TabIndex        =   25
      Text            =   "0"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   9600
      TabIndex        =   24
      Text            =   "0"
      Top             =   720
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   3375
      Left            =   6360
      TabIndex        =   23
      Top             =   360
      Width           =   5055
   End
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   3855
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   8400
      TabIndex        =   21
      Top             =   4080
      Width           =   6255
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7560
      TabIndex        =   19
      Text            =   "Text5"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox yy 
      Height          =   405
      Left            =   7320
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox dd 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox mm 
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5520
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5280
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Year"
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Day"
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "    Month"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Year"
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Day"
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "    Month"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   7560
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Stock Inventory"
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
      Left            =   3360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   495
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
      TabIndex        =   30
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Extrac exel file to drive D:"
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
      TabIndex        =   2
      Top             =   3120
      Width           =   4095
   End
End
Attribute VB_Name = "extract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

Text2.Text = ""
Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)





Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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


Text4.Text = adlaw





End Sub

Private Sub Combo1_Click()
Combo1.Locked = True
Text2.Text = ""
Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)
End Sub

Private Sub Combo1_DropDown()
Combo1.Locked = False




End Sub

Private Sub Combo2_Change()
Text2.Text = ""
Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)



Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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


Text4.Text = adlaw




End Sub

Private Sub Combo2_Click()
Combo2.Locked = True
Text2.Text = ""
Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)
End Sub

Private Sub Combo2_DropDown()
Combo2.Locked = False
End Sub

Private Sub Combo3_Change()
Text2.Text = ""
Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)



Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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


Text4.Text = adlaw



End Sub

Private Sub Combo3_Click()
Combo3.Locked = True
Text2.Text = ""
Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)
End Sub

Private Sub Combo3_DropDown()
Combo3.Locked = False
End Sub

Private Sub Combo4_Change()
Text3.Text = ""
Text3.Text = Combo4.ListIndex + 1 & "/" & Combo5.List(Combo5.ListIndex) & "/" & Combo6.List(Combo6.ListIndex)




Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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


Text4.Text = adlaw




End Sub

Private Sub Combo4_Click()
Combo4.Locked = True
Text3.Text = ""
Text3.Text = Combo4.ListIndex + 1 & "/" & Combo5.List(Combo5.ListIndex) & "/" & Combo6.List(Combo6.ListIndex)
End Sub

Private Sub Combo4_DropDown()
Combo4.Locked = False
End Sub

Private Sub Combo5_Change()
Text3.Text = ""
Text3.Text = Combo4.ListIndex + 1 & "/" & Combo5.List(Combo5.ListIndex) & "/" & Combo6.List(Combo6.ListIndex)



Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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


Text4.Text = adlaw




End Sub

Private Sub Combo5_Click()
Combo5.Locked = True
Text3.Text = ""
Text3.Text = Combo4.ListIndex + 1 & "/" & Combo5.List(Combo5.ListIndex) & "/" & Combo6.List(Combo6.ListIndex)
End Sub

Private Sub Combo5_DropDown()
Combo5.Locked = False
End Sub

Private Sub Combo6_Change()
Text3.Text = ""
Text3.Text = Combo4.ListIndex + 1 & "/" & Combo5.List(Combo5.ListIndex) & "/" & Combo6.List(Combo6.ListIndex)




Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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


Text4.Text = adlaw




End Sub

Private Sub Combo6_Click()
Combo6.Locked = True
Text3.Text = ""
Text3.Text = Combo4.ListIndex + 1 & "/" & Combo5.List(Combo5.ListIndex) & "/" & Combo6.List(Combo6.ListIndex)
End Sub

Private Sub Combo6_DropDown()
Combo6.Locked = False
End Sub

Private Sub Command1_Click()


Label2.Caption = "Please wait a moment!!"


If Val(Text4.Text) = 0 Then
Exit Sub
End If
'Me.Visible = False
extract_sales


End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 End If
End Sub

Private Sub Command2_Click()
'Shell ("d:/sales report on 11-2-2010.xls")
'Shell ("sales report.xls")
'extract_stock
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me

 End If
 

 
 
End Sub



Private Sub extract_sales()

Text6.Text = Combo1.Text & " " & Combo2.Text & ", " & Combo3.Text & " to " & Combo4.Text & " " & Combo5.Text & ", " & Combo6.Text
MsgBox "Extract report from " & Combo1.Text & " " & Combo2.Text & ", " & Combo3.Text & " to " & Combo4.Text & " " & Combo5.Text & ", " & Combo6.Text


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


Set AppXls = CreateObject("Excel.Application")
Set ObjWb = AppXls.Workbooks.Add
Set ObjWs = ObjWb.Worksheets.Add

    ObjWs.Cells(1, 1) = "Item List"
    ObjWs.Cells(1, 4) = "Total No. of Items:"
    ObjWs.Cells(1, 8) = "Total No. of Payments:"
    ObjWs.Cells(1, 9) = "0"
    
    ObjWs.Cells(2, 1) = "DATE_RECEIVED"
    ObjWs.Cells(2, 2) = "DATE_SOLD"
    ObjWs.Cells(2, 3) = "CATEGORY"
    ObjWs.Cells(2, 4) = "MODEL"
    ObjWs.Cells(2, 5) = "DESCRIPTION"
    ObjWs.Cells(2, 6) = "ITEM CODE"
    ObjWs.Cells(2, 7) = "SUPPLIER_NAME"
    ObjWs.Cells(2, 8) = "CP"
    ObjWs.Cells(2, 9) = "RP"
    ObjWs.Cells(2, 10) = "MARGIN_PESO"
    ObjWs.Cells(2, 11) = "MARGIN"
    ObjWs.Cells(2, 12) = "QTY SOLD"
    ObjWs.Cells(2, 13) = "GROSS_SALES"
    ObjWs.Cells(2, 15) = "NET_SALES"
    ObjWs.Cells(2, 14) = "DISCOUNT"
    ObjWs.Cells(2, 16) = "STOCK ON HAND"
    ObjWs.Cells(2, 17) = "AGEING"
    ObjWs.Cells(2, 18) = "CASHER"
    ObjWs.Cells(2, 19) = "INVOICE"
     


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
rs1.Open "SELECT * FROM stock_info", conn
 
 
 
mm.Text = (Combo1.ListIndex + 1)
dd.Text = Combo2.List(Combo2.ListIndex)
yy.Text = Combo3.List(Combo3.ListIndex)

texto = Combo1.List(Combo1.ListIndex) & " " & dd & ", " & yy
texto1 = Combo4.List(Combo4.ListIndex) & " " & Combo5.List(Combo5.ListIndex) & ", " & Combo6.List(Combo6.ListIndex)

 
Do Until rs1.EOF
i1 = 0
    For Each fld In rs1.Fields
        aab(i1) = fld.Value
        i1 = i1 + 1
    Next
 
    i2 = i2 + 1
    
    
    Text7.Text = i2
    Label4.Caption = "Gathering all items! Item count: " & i2
    List4.AddItem aab(0)
 
 
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
adlaw = Val(dd.Text) + 31
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


adlaw = adlaw + (buwan * 31) + (tuig * 12 * 30)

'--------- AGEING end------------
 


    ObjWs.Cells(i2 + 2, 1) = aab(4)
    ObjWs.Cells(i2 + 2, 2) = ""
    ObjWs.Cells(i2 + 2, 3) = aab(0)
    ObjWs.Cells(i2 + 2, 4) = aab(2)
    ObjWs.Cells(i2 + 2, 5) = aab(1)
    ObjWs.Cells(i2 + 2, 6) = aab(3)
    ObjWs.Cells(i2 + 2, 7) = aab(5)
    ObjWs.Cells(i2 + 2, 8) = aab(6)
    ObjWs.Cells(i2 + 2, 9) = aab(7)
    ObjWs.Cells(i2 + 2, 10) = aab(8)
    ObjWs.Cells(i2 + 2, 11) = aab(9) & "%"
    ObjWs.Cells(i2 + 2, 12) = ""
    ObjWs.Cells(i2 + 2, 13) = ""
    ObjWs.Cells(i2 + 2, 15) = ""
    ObjWs.Cells(i2 + 2, 14) = ""
    ObjWs.Cells(i2 + 2, 16) = aab(10)
    ObjWs.Cells(i2 + 2, 17) = adlaw
    ObjWs.Cells(i2 + 2, 18) = ""
    ObjWs.Cells(i2 + 2, 19) = ""
 
 
 
 
    rs1.MoveNext
Loop



 ObjWs.Cells(1, 5) = i2
 
 ObjWs.Cells(i2 + 5, 1) = "List of Payments"
 
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 1) = "DATE SOLD"
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 2) = "ITEM CODE"
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 3) = "DESCRIPTION"
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 4) = "DISCOUNT"
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 5) = "GROSS"
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 6) = "CASHER"
    ObjWs.Cells(ObjWs.Cells(1, 5) + 6, 7) = "INVOICE"
  
 
 
 
 Dim counterx As Boolean
 counterx = False
 
 Dim aac(20) As String
'-----------------other value start-------------------
  For aaaa = 1 To Val(Text4.Text) + 20
  List1.AddItem aaaa
 
 rs2.Open "SELECT * FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
 
 
Label2.Caption = "Creating exel file!!"
 Do Until rs2.EOF
    i11 = 0
    
    For Each fld In rs2.Fields
        aac(i11) = fld.Value
        i11 = i11 + 1
        
    Next
    
    i11 = 0
    List3.AddItem Trim(aac(4))

    For nextfor = 3 To Val(ObjWs.Cells(1, 5))
    
   
    Label4.Caption = "Date: " & mm.Text & "/" & dd.Text & "/" & yy.Text & ". Searching " & Trim(aac(4)) & " on Item No. " & Val(Text10.Text)
    
    
    
    If Trim(aac(4)) = Trim(ObjWs.Cells(nextfor, 6)) Then
  
    Text11.Text = mm.Text & "/" & dd.Text & "/" & yy.Text
    List2.AddItem Trim(aac(4))
    Text9.Text = Trim(aac(4))
    
    On Error GoTo patay_h
    
    Text8.Text = Val(Text8.Text) + 1
    ObjWs.Cells(nextfor, 2) = aac(0)
    ObjWs.Cells(nextfor, 12) = Val(ObjWs.Cells(nextfor, 12)) + Val(aac(5))
     On Error GoTo patay_h
    ObjWs.Cells(nextfor, 13) = Val(ObjWs.Cells(nextfor, 13)) + Val(aac(12))
     On Error GoTo patay_h
    ObjWs.Cells(nextfor, 15) = Val(ObjWs.Cells(nextfor, 15)) + Val(aac(10))
     On Error GoTo patay_h
    ObjWs.Cells(nextfor, 14) = Val(ObjWs.Cells(nextfor, 14)) + Val(aac(11))
    
    ObjWs.Cells(nextfor, 18) = aac(13)
    ObjWs.Cells(nextfor, 19) = aac(14)
    
    counterx = True
    Text10.Text = 0
    
    Else
patay_h:
    Text10.Text = Val(Text10.Text) + 1
    If (Val(Text10.Text) + 3) = Val(ObjWs.Cells(1, 5)) Then
    counterx = False
    Text10.Text = 0
    End If
    
    End If
    
       
    
    Next
    
    
    
    If counterx = False Then
    
    Text10.Text = 0
    counterx = False
    End If
    
    

    If counterx = False Then
    
    List5.AddItem Trim(aac(0)) & " " & aac(4)
    counterx = False
    Text10.Text = 0
    
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 1) = aac(0)
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 2) = aac(4)
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 3) = aac(3)
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 4) = aac(11)
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 5) = aac(12)
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 6) = aac(13)
    ObjWs.Cells(Val(ObjWs.Cells(1, 5)) + 7 + Val(ObjWs.Cells(1, 9)), 7) = aac(14)
   
    ObjWs.Cells(1, 9) = Val(ObjWs.Cells(1, 9)) + 1

    
    
    
    End If
    


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



'-----------------other value end-----------------------

 

 


aba:
yyy:
unsa_ni:
'MsgBox Combo1.Text & " " & Combo2.Text & ", " & Combo3.Text & " to " & Combo4.Text & " " & Combo5.Text & ", " & Combo6.Text

     
     
     

     
     
     'Me.Visible = True
    ObjWb.SaveAs ("d:\sales report on " & texto & " to " & texto1 & ".xls")
    
    'ObjWb.SaveAs ("d:\sales gghereport on " & Text3.Text & " to " & Text4.Text & ".xls")
    
    
    ObjWb.Close (SaveChanges = False)
    Me.Enabled = True
  
'Shell ("d:\sales report on " & mm.Text & "-" & dd.Text & "-" & yy.Text & ".xls")

'main.Show
'main.Enabled = True
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
Combo1.AddItem "January"
Combo1.AddItem "Febuary"
Combo1.AddItem "March"
Combo1.AddItem "April"
Combo1.AddItem "May"
Combo1.AddItem "June"
Combo1.AddItem "July"
Combo1.AddItem "August"
Combo1.AddItem "September"
Combo1.AddItem "October"
Combo1.AddItem "November"
Combo1.AddItem "December"

For aa = 1 To 31
Combo2.AddItem aa
Next

Combo3.AddItem "2008"
Combo3.AddItem "2009"
Combo3.AddItem "2010"
Combo3.AddItem "2011"
Combo3.AddItem "2012"
Combo3.AddItem "2013"
Combo3.AddItem "2014"
Combo3.AddItem "2015"
Combo3.AddItem "2016"
Combo3.AddItem "2017"
Combo3.AddItem "2018"
Combo3.AddItem "2019"
Combo3.AddItem "2020"
Combo3.AddItem "2021"
Combo3.AddItem "2022"
Combo3.AddItem "2023"
Combo3.AddItem "2024"
Combo3.AddItem "2025"
Combo3.AddItem "2026"
Combo3.AddItem "2027"
Combo3.AddItem "2028"
Combo3.AddItem "2029"
Combo3.AddItem "2030"


Combo4.AddItem "January"
Combo4.AddItem "Febuary"
Combo4.AddItem "March"
Combo4.AddItem "April"
Combo4.AddItem "May"
Combo4.AddItem "June"
Combo4.AddItem "July"
Combo4.AddItem "August"
Combo4.AddItem "September"
Combo4.AddItem "October"
Combo4.AddItem "November"
Combo4.AddItem "December"

For aa = 1 To 31
Combo5.AddItem aa
Next

Combo6.AddItem "2008"
Combo6.AddItem "2009"
Combo6.AddItem "2010"
Combo6.AddItem "2011"
Combo6.AddItem "2012"
Combo6.AddItem "2013"
Combo6.AddItem "2014"
Combo6.AddItem "2015"
Combo6.AddItem "2016"
Combo6.AddItem "2017"
Combo6.AddItem "2018"
Combo6.AddItem "2019"
Combo6.AddItem "2020"
Combo6.AddItem "2021"
Combo6.AddItem "2022"
Combo6.AddItem "2023"
Combo6.AddItem "2024"
Combo6.AddItem "2025"
Combo6.AddItem "2026"
Combo6.AddItem "2027"
Combo6.AddItem "2028"
Combo6.AddItem "2029"
Combo6.AddItem "2030"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'main.Enabled = True
'main.Show
'Unload Me
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
Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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

Private Sub Text3_Change()
On Error GoTo ss
Dim aa() As String
Dim bb() As String
aa = Split(Text2.Text, "/")
bb = Split(Text3.Text, "/")

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
