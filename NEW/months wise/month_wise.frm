VERSION 5.00
Begin VB.Form M__WISE 
   Caption         =   "Date Wise Extractor"
   ClientHeight    =   9405
   ClientLeft      =   7395
   ClientTop       =   4380
   ClientWidth     =   15300
   Icon            =   "month_wise.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "month_wise.frx":1C64C
   ScaleHeight     =   9405
   ScaleWidth      =   15300
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4920
      TabIndex        =   26
      Text            =   "Text7"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ListBox List6 
      Height          =   2790
      Left            =   600
      TabIndex        =   25
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ListBox List5 
      Height          =   2595
      Left            =   10200
      TabIndex        =   24
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox List4 
      Height          =   2595
      Left            =   10680
      TabIndex        =   23
      Top             =   4440
      Width           =   2295
   End
   Begin VB.ListBox List3 
      Height          =   2985
      Left            =   8160
      TabIndex        =   22
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   4320
      TabIndex        =   21
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   975
      Left            =   840
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8520
      Top             =   4200
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
      Left            =   6360
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   7560
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   2520
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
Attribute VB_Name = "M__WISE"
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

Private Sub Combo3_LostFocus()

'Text2.Text = Combo1.ListIndex + 1 & "/" & Combo2.List(Combo2.ListIndex) & "/" & Combo3.List(Combo3.ListIndex)
mm.Text = Combo1.ListIndex + 1
dd.Text = Combo2.List(Combo2.ListIndex)
yy.Text = Combo3.List(Combo3.ListIndex)

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
MsgBox "Genarate Date Wise Report starting from " & Combo1.Text & " " & Combo2.Text & ", " & Combo3.Text & " to " & Combo4.Text & " " & Combo5.Text & ", " & Combo6.Text
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
Dim item_ccode As String
Dim texto1 As String
Dim report(20) As String


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

Set conn = New ADODB.Connection
  Dim count_er As Integer
    count_er = 0


 conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
 conn.Open


 '----------------------------
 
 Dim mm1 As String
 Dim dd1 As String
 Dim yy1 As String
 mm.Text = (Combo1.ListIndex + 1)
 dd.Text = Combo2.List(Combo2.ListIndex)
 yy.Text = Combo3.List(Combo3.ListIndex)
 
  mm1 = (Combo4.ListIndex + 1)
 dd1 = Combo5.List(Combo5.ListIndex)
 yy2 = Combo6.List(Combo6.ListIndex)

texto = Combo1.List(Combo1.ListIndex) & " " & dd & ", " & yy
texto1 = Combo4.List(Combo4.ListIndex) & " " & Combo5.List(Combo5.ListIndex) & ", " & Combo6.List(Combo6.ListIndex)



texto3 = mm.Text & "/" & Val(dd.Text) + 1 & "/" & yy.Text
conn.Execute "TRUNCATE month_wise"
For aaaa = 0 To (Val(Text4.Text) / 30) + 1


    rs1.Open "SELECT * FROM sales where DATE_SOLD like " & "'" & mm.Text & "/%/" & yy.Text & "%'", conn
    
    
    Do Until rs1.EOF
        i11 = 1
    
    For Each fld In rs1.Fields
        report(i11) = fld.Value
        i11 = i11 + 1
        
    Next
    
    Dim trimer() As String
    
    report(1) = Trim(report(1))
    trimer = Split(report(1), " ")
    report(1) = trimer(0)
    List3.AddItem report(1)
    trimer = Split(report(1), "/")
    report(1) = trimer(0) & "/" & trimer(2)

    List2.AddItem report(1)

        rs2.Open "SELECT GROSS_SALE FROM  month_wise where DATE_SOLD = '" & report(1) & "'", conn
        rs3.Open "SELECT GROSS_MARGIN FROM month_wise where DATE_SOLD = '" & report(1) & "'", conn
        rs4.Open "SELECT QTY_SOLD FROM month_wise where DATE_SOLD = '" & report(1) & "'", conn
        rs5.Open "SELECT S_D FROM month_wise where DATE_SOLD = '" & report(1) & "'", conn
        
        
      
      
        Dim ok As Boolean
        
        If rs2.EOF Then
        ok = True
        Else
        ok = False
        End If
       
        
        
        If ok = True Then
        report(10) = Val(report(10))
        report(9) = Val(report(9))
        report(6) = Val(report(6))
                        On Error GoTo mail

mail:

                'If report(9) = 0 Then
                'conn.Execute "INSERT INTO cat_wise (CATEGORY,GROSS_SALE, GROSS_MARGIN,QTY_SOLD) values ('(Suki Card/Payment)" & report(2) & "','" & report(11) & "',' - - - -','" & report(6) & "')"
                'MsgBox "dd"
                ' End If
                 
            
               
                If report(12) = "--Suki Card--" Then
                report(12) = "0"
                End If
                If report(13) = "--Suki Card--" Then
                report(13) = report(11)
                End If
                
               If report(9) <> 0 Then
                conn.Execute "INSERT INTO month_wise (DATE_SOLD,GROSS_SALE, GROSS_MARGIN,S_D,QTY_SOLD) values ('" & report(1) & "','" & report(11) * Val(report(6)) & "','" & (Val(report(10)) - Val(report(9))) * Val(report(6)) & "','" & Val(report(12)) + Val(report(13)) & "','" & report(6) & "')"
                List4.AddItem report(11) * Val(report(6))
                List5.AddItem report(1)
                End If
                
        End If

        
        
        
If ok = False Then
         
        Do Until rs2.EOF
        
        For Each fld2 In rs2.Fields
        conn.Execute "UPDATE month_wise SET GROSS_SALE = '" & ((Val(report(11)) * Val(report(6))) + Val(fld2.Value)) - Val(report(12)) & " ' WHERE DATE_SOLD = '" & report(1) & "'"
        List4.AddItem fld2.Value
        Next
        
        For Each fld2 In rs3.Fields
        conn.Execute "UPDATE month_wise SET GROSS_MARGIN = '" & ((Val(report(10)) - Val(report(9))) * Val(report(6))) + Val(fld2.Value) & "' WHERE DATE_SOLD = '" & report(1) & "'"
        Next
        
        For Each fld2 In rs4.Fields
        conn.Execute "UPDATE month_wise SET QTY_SOLD = '" & Val(report(6)) + Val(fld2.Value) & "' WHERE DATE_SOLD = '" & report(1) & "'"
        
        Next



        If report(12) = "--Suki Card--" Then
        report(12) = "0"
        End If
        If report(13) = "--Suki Card--" Then
        report(13) = report(11)
        Else: report(13) = "0"
        End If
        
        
         
        For Each fld2 In rs5.Fields
        conn.Execute "UPDATE month_wise SET S_D = '" & Val(report(12)) + Val(report(13)) + Val(fld2.Value) & " ' WHERE DATE_SOLD = '" & report(1) & "'"
        Next
      
        rs2.MoveNext
        rs3.MoveNext
        rs4.MoveNext
        rs5.MoveNext
        
        
        Loop
            
                
End If

                
                report(2) = Val(report(2))
                
                


            
        
        'On Error GoTo out_me
        rs2.Close
        rs3.Close
        rs4.Close
        rs5.Close

        
        
        
        
        
    rs1.MoveNext
  
    
    Text7.Text = report(1) & " " & count_er
    count_er = count_er + 1
    Loop
        
    rs1.Close
    
    
    
    
    
    
    





mm = Val(mm.Text)
dd = Val(dd.Text)
yy = Val(yy.Text)
mm = mm + 1

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
    ObjWs.Cells(1, 2) = "GROSS SALES"
    ObjWs.Cells(1, 3) = "GROSE MARGIN"
    ObjWs.Cells(1, 4) = "QTY SOLD"
    ObjWs.Cells(1, 5) = "SUKI CARD/DISCOUNT"

     
     
     rs7.Open "SELECT * FROM month_WISE", conn
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
    ObjWs.Cells(i2, 4) = report(4)
    ObjWs.Cells(i2, 5) = report(5)
    rs7.MoveNext
    i2 = i2 + 1
    Loop
    
 
     
     
     

     
     
     'Me.Visible = True
    ObjWb.SaveAs ("d:\monhth WISE report starting from " & texto & " to " & texto1 & ".xls")
    
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
''Unload Me
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

Dim aa As Integer
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
