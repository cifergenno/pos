VERSION 5.00
Begin VB.Form extract 
   Caption         =   "Report Extractor"
   ClientHeight    =   3615
   ClientLeft      =   7395
   ClientTop       =   4380
   ClientWidth     =   4635
   Icon            =   "extract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "extract.frx":1C64C
   ScaleHeight     =   3615
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   2880
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
      Left            =   4920
      TabIndex        =   3
      Top             =   3720
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
      Top             =   3240
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
Dim texto As String

'GoTo yyy
Label2.Caption = "Please wait a moment!!"
'Me.Enabled = False
Dim ihap As Integer
ihap = 0

Dim aa() As String
Dim adlaw As Integer
Dim adlaw2 As Integer
Dim buwan As Integer
Dim tuig As Integer
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
Dim rs11 As ADODB.Recordset
Dim rs12 As ADODB.Recordset
Dim rs13 As ADODB.Recordset
Dim rs14 As ADODB.Recordset
Dim rs15 As ADODB.Recordset
Dim rs16 As ADODB.Recordset
Dim rs17 As ADODB.Recordset
Dim rs18 As ADODB.Recordset
Dim rs19 As ADODB.Recordset
Dim rs20 As ADODB.Recordset
Dim rs21 As ADODB.Recordset
Dim rs22 As ADODB.Recordset
Dim rs23 As ADODB.Recordset
Dim conn As ADODB.Connection
Dim fld As ADODB.Field
Dim fld1 As ADODB.Field
Dim fld2 As ADODB.Field

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
Set rs12 = New ADODB.Recordset
Set rs13 = New ADODB.Recordset
Set rs14 = New ADODB.Recordset
Set rs15 = New ADODB.Recordset
Set rs11 = New ADODB.Recordset
Set rs12 = New ADODB.Recordset
Set rs13 = New ADODB.Recordset
Set rs14 = New ADODB.Recordset
Set rs15 = New ADODB.Recordset
Set rs16 = New ADODB.Recordset
Set rs17 = New ADODB.Recordset
Set rs18 = New ADODB.Recordset
Set rs19 = New ADODB.Recordset
Set rs20 = New ADODB.Recordset
Set rs21 = New ADODB.Recordset
Set rs22 = New ADODB.Recordset
Set rs23 = New ADODB.Recordset

Dim texto1 As String






Set conn = New ADODB.Connection
Dim item_ccode As String
 conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
 conn.Open
 
 'GoTo yyy
 GoTo ahem
 
conn.Execute "TRUNCATE sales_extrac"


 
 '----------------------------
 
 
 mm.Text = (Combo1.ListIndex + 1)
 dd.Text = Combo2.List(Combo2.ListIndex)
 yy.Text = Combo3.List(Combo3.ListIndex)
 'MsgBox mm & "/" & dd & "/" & yy
  'GoTo ahem
 'GoTo patay
texto = Combo1.List(Combo1.ListIndex) & " " & dd & ", " & yy
texto1 = Combo4.List(Combo4.ListIndex) & " " & Combo5.List(Combo5.ListIndex) & ", " & Combo6.List(Combo6.ListIndex)
 

rs1.Open "SELECT CATEGORY FROM stock_info", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info", conn
rs3.Open "SELECT MODEL FROM stock_info", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info", conn
rs11.Open "SELECT DATE_RECEIVED FROM stock_info", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info", conn
rs7.Open "SELECT CP FROM stock_info", conn
rs6.Open "SELECT RP FROM stock_info", conn
rs10.Open "SELECT MARGIN_PESO FROM stock_info", conn
rs9.Open "SELECT MARGIN FROM stock_info", conn
rs8.Open "SELECT STOCK_ON_HAND FROM stock_info", conn






'GoTo unsa_ni


Do Until rs4.EOF




For Each fld In rs4.Fields
conn.Execute "INSERT INTO sales_extrac (ITEM_CODE)" _
& "values ('" & fld.Value & "')"
item_ccode = fld.Value
Next

For Each fld In rs1.Fields
conn.Execute "UPDATE po_extrac SET CATEGORY = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next

For Each fld In rs2.Fields
conn.Execute "UPDATE po_extrac SET DESCRIPTION = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next

For Each fld In rs3.Fields
conn.Execute "UPDATE po_extrac SET MODEL = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next

For Each fld In rs5.Fields
conn.Execute "UPDATE po_extrac SET SUPPLIER_NAME = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next

For Each fld In rs6.Fields
conn.Execute "UPDATE po_extrac SET RP = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next

For Each fld In rs7.Fields
conn.Execute "UPDATE po_extrac SET CP = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next

For Each fld In rs8.Fields
conn.Execute "UPDATE po_extrac SET STOCK_ON_HAND = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next




adlaw = 0
buwan = 0
tuig = 0
For Each fld In rs11.Fields
conn.Execute "UPDATE sales_extrac SET DATE_RECEIVED = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
If Trim(fld.Value) = "" Then
aaa = "0/0/0"
Else
aaa = fld.Value
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


adlaw = adlaw + (buwan * 30) + (tuig * 12 * 30)


conn.Execute "UPDATE salse_extrac SET AGEING = " & "'" & adlaw & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
Next


rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
rs6.MoveNext
rs7.MoveNext
rs8.MoveNext
rs11.MoveNext

'rs17.MoveNext



ihpa = ihap + 1

Loop

ahem:





    ihap = 0
     
 'Do Until (Combo4.ListIndex + 1) = mm.Text And Combo5.List(Combo5.ListIndex) = dd.Text And Combo6.List(Combo6.ListIndex) = yy.Text
  For aaaa = 0 To Val(Text4.Text)
  

rs12.Open "SELECT PCS FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
'rs13.Open "SELECT TOTAL FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
'rs14.Open "SELECT CASHER FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
'rs15.Open "SELECT INVOICE FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
'rs16.Open "SELECT DATE_SOLD FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
rs17.Open "SELECT ITEM_CODE FROM sales where DATE_SOLD like  " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn

'rs20.Open "SELECT GROSS FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
'rs21.Open "SELECT DISCOUNT FROM sales where DATE_SOLD like  " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn




Do Until rs17.EOF



For Each fld In rs17.Fields
item_ccode = fld.Value
Next

For Each fld In rs12.Fields

    rs18.Open "SELECT QUANTITY FROM po_extrac WHERE ITEM_CODE = '" & item_ccode & "'", conn
    rs21.Open "SELECT QTY_SOLD FROM po_extract where ITEM_CODE = '" & item_ccode & "'", conn


    For Each fld1 In rs18.Fields

        If fld1.Value = "" Then
        conn.Execute "UPDATE po_extrac SET QUANTITY = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
        Else
        conn.Execute "UPDATE po_extrac SET QUANTITY = " & "'" & Val(fld.Value) + Val(fld1.Value) & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
        End If
    Next
    rs18.Close
    
    For Each fld2 In rs21.Fields
        If fld2.Value = "" Then
        conn.Execute "UPDATE po_extrac SET QUANTITY = " & "'" & fld.Value & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
        Else
        conn.Execute "UPDATE po_extrac SET QUANTITY = " & "'" & Val(fld.Value) + Val(fld2.Value) & "'" & " WHERE ITEM_CODE = '" & item_ccode & "'"
        End If
    Next
    rs21.Close
Next
        
        





rs12.MoveNext
'rs13.MoveNext
'rs14.MoveNext
'rs15.MoveNext
'rs16.MoveNext
rs17.MoveNext
'rs21.MoveNext
'rs20.MoveNext

ihpa = ihap + 1




Loop





rs12.Close
'rs13.Close
'rs14.Close
'rs15.Close
rs17.Close
'rs16.Close
'rs21.Close
'rs20.Close



mm = Val(mm.Text)
dd = Val(dd.Text)
yy = Val(yy.Text)
dd = dd + 1
If dd = 31 Then
mm = mm + 1
End If
If mm = 12 Then
yy = yy + 10
End If

mm.Text = mm
yy.Text = yy
dd.Text = dd
'MsgBox mm & "/" & dd & "/" & yy

Next
'Unload Me
'main.Show
'main.Enabled = True


'rs16.Close

'GoTo sibat
'-----------------------------------
'-----------------------------------
'-----------------------------------

Label2.Caption = "Creating exel file!!"
aba:
yyy:
unsa_ni:
'MsgBox Combo1.Text & " " & Combo2.Text & ", " & Combo3.Text & " to " & Combo4.Text & " " & Combo5.Text & ", " & Combo6.Text

Dim AppXls As New Excel.Application
Dim ObjWb As Excel.Workbook
Dim ObjWs As Excel.Worksheet
Dim xx As Integer
Dim xx2 As Integer
     
     
xx = 1
xx2 = 2
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
     
     
     rs20.Open "SELECT * FROM sales_extrac", conn
     
     
     Do Until rs20.EOF

For Each fld In rs20.Fields
ObjWs.Cells(xx2, xx) = fld.Value
xx = xx + 1
Next
xx = 1
xx2 = xx2 + 1
rs20.MoveNext
Loop
     
     
     

     
     
     'Me.Visible = True
    ObjWb.SaveAs ("d:\sales report on " & texto & " to " & texto1 & ".xls")
    
    'ObjWb.SaveAs ("d:\sales gghereport on " & Text3.Text & " to " & Text4.Text & ".xls")
    
    
    ObjWb.Close (SaveChanges = False)
    Me.Enabled = True
   MsgBox ("Saving finished!!")
'Shell ("d:\sales report on " & mm.Text & "-" & dd.Text & "-" & yy.Text & ".xls")
Unload Me
'main.Show
'main.Enabled = True
sibat:
Exit Sub
    
End Sub


Private Sub extract_stock()

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
Set ObjWb = AppXls.Workbooks.Add
Set ObjWs = ObjWb.Worksheets.Add
    
    ObjWs.Cells(1, 1) = "CATEGORY"
    ObjWs.Cells(1, 2) = "DESCRIPTION"
    ObjWs.Cells(1, 3) = "MODEL"
    ObjWs.Cells(1, 4) = "ITEM CODE"
    ObjWs.Cells(1, 5) = "UNIT"
    ObjWs.Cells(1, 6) = "DATE RECEIVED"
    ObjWs.Cells(1, 7) = "SPPLIER NAME"
    ObjWs.Cells(1, 8) = "CP"
    ObjWs.Cells(1, 9) = "RP"
    ObjWs.Cells(1, 10) = "STOCK ON HAND"
     
     
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
    conn.Open
    Set rs = New ADODB.Recordset
 
 rs.Open "SELECT * FROM stock_info", conn
Do Until rs.EOF

For Each fld In rs.Fields
ObjWs.Cells(xx2, xx) = fld.Value
xx = xx + 1
Next
xx = 1
xx2 = xx2 + 1
rs.MoveNext
Loop
     
     
     
     
     
     
     
    ObjWb.SaveAs ("d:\stock report.xls")
    ObjWb.Close (SaveChanges = False)
    MsgBox ("Saving finished!!")
    
Unload Me
main.Show
main.Enabled = True
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
