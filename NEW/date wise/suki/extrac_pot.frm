VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DATE__WISE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Wise Extractor"
   ClientHeight    =   3870
   ClientLeft      =   7320
   ClientTop       =   4305
   ClientWidth     =   4440
   Icon            =   "extrac_pot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "extrac_pot.frx":1C64C
   ScaleHeight     =   3870
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List4 
      Height          =   3180
      Left            =   8040
      TabIndex        =   18
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   9120
      TabIndex        =   17
      Text            =   "0"
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   3375
      Left            =   5520
      TabIndex        =   16
      Top             =   3960
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   3000
      TabIndex        =   15
      Top             =   3960
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   720
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   8880
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   10878976
      CurrentDate     =   40594
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   10878976
      CurrentDate     =   40594
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8520
      Top             =   4200
   End
   Begin VB.TextBox yy 
      Height          =   405
      Left            =   7320
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox dd 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox mm 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1680
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
      Caption         =   "Genarate Report"
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   1
      Top             =   3240
      Width           =   4095
   End
End
Attribute VB_Name = "DATE__WISE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adlaw As Integer
Public buwan As Integer
Public tuig As Integer
Public iha As Integer
Public ihpa As Integer


Public yy2 As Integer

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

Private Sub DTPicker1_Change()
Text2.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
Text3.Text = DTPicker2.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me

 End If
 

 
 
End Sub



Private Sub extract_sales()


Dim texto As String

'GoTo yyy
Label2.Caption = "Please wait a moment!!"
'Me.Enabled = False
Dim ihap As Integer
ihap = 0

'On Error GoTo sibat

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
Dim rs9 As ADODB.Recordset

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
Set rs9 = New ADODB.Recordset




Dim texto1 As String
Dim report(20) As String





Set conn = New ADODB.Connection
Dim item_ccode As String
 conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
 conn.Open



 '----------------------------
 
 Dim mm1 As String
 Dim dd1 As String
 Dim yy1 As String



conn.Execute "TRUNCATE date_wise"
conn.Execute "truncate bad_items"




For aaaa = 1 To Val(Text4.Text)


    rs1.Open "SELECT * FROM sales where DATE_SOLD like " & "'" & mm.Text & "/" & dd.Text & "/" & yy.Text & "%'", conn
    
    
    Do Until rs1.EOF
        i11 = 1
    
 Text6.Text = Val(Text6.Text) + 1
    For Each fld In rs1.Fields
        report(i11) = fld.Value
        i11 = i11 + 1
        
    Next
    
    Dim trimer() As String
    Dim query(20) As String
    
    report(1) = Trim(report(1))
    trimer = Split(report(1), " ")
    report(1) = trimer(0)


      
        Dim ok As Boolean
        
        rs2.Open "SELECT * FROM  date_wise where DATE_SOLD like '" & report(1) & "%'", conn
        
        If rs2.EOF Then
        ok = True
        Else
        ok = False
        End If
        
     rs2.Close
       
        
' if end of file (wa pa na sulod???)
        
        If ok = True Then
'        report(10) = Val(report(10))
'        report(9) = Val(report(9))
'        report(6) = Val(report(6))
'                        On Error GoTo mail
'
'mail:
'
'
 '              If report(9) <> 0 Then
 '               conn.Execute "INSERT INTO date_wise ( DATE_SOLD,CP,DISCOUNT,GROSS_SALE, GROSS_MARGIN,SUKI,QTY_SOLD)" _
 '              & "values ('" & report(1) & "','" & report(9) & "','" & report(11) * Val(report(6)) & "','" & (Val(report(10)) - Val(report(9))) * Val(report(6)) & "','" & Val(report(12)) + Val(report(13)) & "','" & report(6) & "')"
 '               End If
 
 
        rs3.Open "SELECT * FROM  stock_info where ITEM_CODE = '" & report(5) & "'", conn
                

' eval if good item

         If rs3.EOF Then
         
'if bad item
            conn.Execute "INSERT INTO bad_items ( KIND,CP,GROSS_SALE,GROSS_MARGIN,QTY_SOLD,DISCOUNT,DESCRIPTION, ITEM_CODE)" _
            & "values ('" & report(1) & "','" & report(9) & "','" & report(13) & "','" & report(16) & "','" & report(6) & "','" & report(12) & "','" & report(4) & "','" & report(5) & "')"
         
         Else
'if good item

            conn.Execute "INSERT INTO date_wise ( DATE_SOLD,CP,GROSS_SALE,GROSS_MARGIN,QTY_SOLD,DISCOUNT)" _
            & "values ('" & report(1) & "','" & report(9) & "','" & report(13) & "','" & report(16) & "','" & report(6) & "','" & report(12) & "')"


            List1.AddItem report(9)
            'List2.AddItem report(6)
            List3.AddItem Val(report(9)) * Val(report(6))

         End If
         rs3.Close
'eval end of items
        End If
        
' wala pa na panso end




'if item already exist
        
If ok = False Then
         
        
        
        
'        Do Until rs2.EOF
        
        


'for suki card -----
'        If report(5) = "-SUKI_CARD-" Then
'        For Each fld2 In rs5.Fields
'        conn.Execute "UPDATE date_wise SET SUKI_CARD = '" & Val(report(12)) + Val(fld2.Value) & " ' WHERE DATE_SOLD = '" & report(1) & "'"
'        Next
        
'        Else
'
' if not suki card

'        For Each fld2 In rs2.Fields
''        conn.Execute "UPDATE date_wise SET GROSS_SALE = '" & ((Val(report(11)) * Val(report(6))) + Val(fld2.Value)) - Val(report(12)) & " ' WHERE DATE_SOLD = '" & report(1) & "'"
'        Next
        
'        For Each fld2 In rs3.Fields
'        conn.Execute "UPDATE date_wise SET GROSS_MARGIN = '" & ((Val(report(10)) - Val(report(9))) * Val(report(6))) + Val(fld2.Value) & "' WHERE DATE_SOLD = '" & report(1) & "'"
'        Next
        
'        For Each fld2 In rs4.Fields
'        conn.Execute "UPDATE date_wise SET QTY_SOLD = '" & Val(report(6)) + Val(fld2.Value) & "' WHERE DATE_SOLD = '" & report(1) & "'"
'        Next
        
'        For Each fld2 In rs6.Fields
'        conn.Execute "UPDATE date_wise SET CP = '" & Val(report(9)) + Val(fld2.Value) & "' WHERE DATE_SOLD = '" & report(1) & "'"
'        Next
        
        
'        End If
'
        
'for discount ------

'        If Val(report(12)) <> 0 And report(5) <> "" And report(5) <> "-SUKI_CARD-" Then
'        On Error GoTo ops
'ops:
'        For Each fld2 In rs7.Fields
'        conn.Execute "UPDATE date_wise SET DISCOUNT = '" & Val(report(12)) + Val(fld2.Value) & "' WHERE DATE_SOLD = '" & report(1) & "'"
'        Next
        '
'        End If
        
        

        
         
        
        
       
      
'        rs2.MoveNext
'        rs3.MoveNext
'        rs4.MoveNext
'        rs.MoveNext
'        rs6.MoveNext
        
        
        
'        Loop
            
                
                
                
                
                
' if item exist in date_wise start ---------------



rs3.Open "SELECT * FROM  stock_info where ITEM_CODE = '" & report(5) & "'", conn
                
                
                
  ' eval if good item / if exist on stock_info

         If rs3.EOF Then
         
'if bad item
            conn.Execute "INSERT INTO bad_items ( KIND,CP,GROSS_SALE,GROSS_MARGIN,QTY_SOLD,DISCOUNT,DESCRIPTION, ITEM_CODE)" _
            & "values ('" & report(1) & "','" & report(9) & "','" & report(13) & "','" & report(16) & "','" & report(6) & "','" & report(12) & "','" & report(4) & "','" & report(5) & "')"
         
         Else
'if good item

'            conn.Execute "INSERT INTO date_wise ( DATE_SOLD,CP,GROSS_SALE,GROSS_MARGIN,QTY_SOLD,SUKI,DISCOUNT)" _
'            & "values ('" & report(1) & "','" & report(9) & "','" & report(13) & "','" & report(16) & "','" & report(6) & "','" & report(12) & "')"



        rs2.Open "SELECT * FROM  date_wise where DATE_SOLD like '" & report(1) & "%'", conn
    
        Do Until rs2.EOF
        i11 = 1
        For Each fld In rs2.Fields
        query(i11) = fld.Value
        i11 = i11 + 1
        Next
    
        rs2.MoveNext
        Loop
      
rs2.Close

        conn.Execute "UPDATE date_wise SET CP = '" & (Val(report(9)) * Val(report(6))) + Val(query(2)) & "' WHERE DATE_SOLD = '" & report(1) & "'"

        conn.Execute "UPDATE date_wise SET GROSS_SALE = '" & Val(report(13)) + Val(query(3)) & " ' WHERE DATE_SOLD = '" & report(1) & "'"
        
        conn.Execute "UPDATE date_wise SET GROSS_MARGIN = '" & Val(report(16)) + Val(query(4)) & "' WHERE DATE_SOLD = '" & report(1) & "'"

        conn.Execute "UPDATE date_wise SET QTY_SOLD = '" & Val(report(6)) + Val(query(5)) & "' WHERE DATE_SOLD = '" & report(1) & "'"

        conn.Execute "UPDATE date_wise SET DISCOUNT = '" & Val(report(12)) + Val(query(7)) & "' WHERE DATE_SOLD = '" & report(1) & "'"

            List1.AddItem report(6)
            
            List3.AddItem Val(report(9)) * Val(report(6))
           List4.AddItem (Val(report(9)) * Val(report(6))) + Val(query(2))
           List2.AddItem query(2)

         End If
         rs3.Close
'eval end of items
                
                
                
                
                
End If

                
           '     report(2) = Val(report(2))
                
                


            
        
        'On Error GoTo out_me

        
        
        
       
    rs1.MoveNext
    Loop
        
    rs1.Close
    
    
    
    
    
    
    





mm = Val(mm.Text)
dd = Val(dd.Text)
yy = Val(yy.Text)
dd = dd + 1
If dd = 32 Then
mm = mm + 1
dd = 1
End If
If mm = 13 Then
yy = yy + 1
mm = 1
End If

mm.Text = mm
yy.Text = yy
dd.Text = dd

texto2 = mm.Text & "/" & Val(dd.Text) & "/" & yy.Text
Next
out_me:



'-----------------------------------
'-----------------------------------
'-----------------------------------


'for stock on hand
    
    



Label2.Caption = "Creating exel file!!"
aba:
yyy:



rs7.Open "SELECT * FROM bad_items", conn

     Dim ii2 As Integer
     ii2 = 2
     Do Until rs7.EOF
        ii11 = 1
        
    For Each fld In rs7.Fields
        report(ii11) = fld.Value
        ii11 = ii11 + 1
    Next
    
    Dim split_me() As String
    
    split_me = Split(Trim(report(8)), " ")
    
    If split_me(0) = "Suki" And report(7) <> "--Suki Card--" Then
 
      rs9.Open "SELECT SUKI_CARD FROM  date_wise where DATE_SOLD = '" & report(1) & "'", conn
      
      Do Until rs9.EOF
      
      For Each fld2 In rs9.Fields
      
      conn.Execute "UPDATE date_wise SET SUKI_CARD = '" & Val(report(7)) + Val(fld2.Value) & " ' WHERE DATE_SOLD = '" & report(1) & "'"
      
       Next
       rs9.MoveNext
    Loop
    rs9.Close
    End If
        
     rs7.MoveNext
    Loop
     


    
    'for stock on hand
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
    
    
    ObjWs.Cells(1, 1) = "DATE SOLD"
    ObjWs.Cells(1, 2) = "COST OF GOODS"
    ObjWs.Cells(1, 3) = "GROSE SALES"
    ObjWs.Cells(1, 4) = "GROSS MARGIN"
    ObjWs.Cells(1, 5) = "QTY SOLD"
    ObjWs.Cells(1, 6) = "SUKI_CARD"
    ObjWs.Cells(1, 7) = "DISCOUNT"

     
     rs7.Close
     
     rs7.Open "SELECT * FROM date_WISE", conn
     Dim i2 As Integer
     i2 = 2
     Do Until rs7.EOF
        i11 = 1
        
    For Each fld In rs7.Fields
        report(i11) = fld.Value
        i11 = i11 + 1
    Next
    
    ObjWs.Cells(i2, 1) = report(1)
    ObjWs.Cells(i2, 2) = report(2)
    ObjWs.Cells(i2, 3) = report(3)
    ObjWs.Cells(i2, 4) = Val(report(3)) - Val(report(2))
    ObjWs.Cells(i2, 5) = report(5)
    If Val(report(6)) = 0 Then
    report(6) = 0
    End If
    ObjWs.Cells(i2, 6) = report(6)
    ObjWs.Cells(i2, 7) = report(7)
    rs7.MoveNext
    i2 = i2 + 1
    
    Loop
    
    
   
    
    
     
     

     
     
     'Me.Visible = True
    ObjWb.SaveAs ("d:\DATE WISE report starting from " & Text1.Text & " to " & Text5.Text & ".xls")
    
    'ObjWb.SaveAs ("d:\sales gghereport on " & Text3.Text & " to " & Text4.Text & ".xls")
    
    
    ObjWb.Close (SaveChanges = False)
    Me.Enabled = True
  
part_:


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
DTPicker1.Value = Format(Now, "d  mm, yyyy")
DTPicker2.Value = Format(Now, "d  mm, yyyy")
Text2.Text = DTPicker1.Value
Text3.Text = DTPicker2.Value
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
Dim aa11() As String
Dim bb() As String
aa = Split(Text2.Text, "/")

dd.Text = aa(1)
mm.Text = aa(0)
yy.Text = aa(2)

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
