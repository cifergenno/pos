VERSION 5.00
Begin VB.Form transac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   5760
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List4 
      Height          =   1425
      ItemData        =   "transac.frx":0000
      Left            =   3360
      List            =   "transac.frx":0007
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   1620
      ItemData        =   "transac.frx":001A
      Left            =   1440
      List            =   "transac.frx":0021
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   450
      ItemData        =   "transac.frx":0030
      Left            =   480
      List            =   "transac.frx":0037
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   360
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1800
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "Text14"
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text13 
      Height          =   195
      Left            =   1080
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   4320
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Left            =   7680
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Suki Card"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7680
      Width           =   5130
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "&End Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   45
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text10"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text9"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   " ----------------------"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Text6"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Text            =   " ----------------------"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "--------------"
      Top             =   840
      Width           =   7575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Credit"
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
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cash Sales"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8160
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount  Value"
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
      Left            =   3600
      TabIndex        =   25
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Suki Points"
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
      Left            =   240
      TabIndex        =   24
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tendered"
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
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
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
      Left            =   4080
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "  Sub Total"
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
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "transac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public reg As Boolean
Public old_val As Double
Public conn As ADODB.Connection
Public fld1 As ADODB.Field
Public fld2 As ADODB.Field
Public fld3 As ADODB.Field
Public fld4 As ADODB.Field
Public fld5 As ADODB.Field
Public fld6 As ADODB.Field
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
Public i_save As Double
Public sev As Integer
Public sudli As Double
Public dakpa As Double


Private Sub Check1_Click()
Text9.Text = "  "
Text9.Text = ""
Text7.Text = Val(Text6.Text) - Val(Text11.Text)
old_val = Val(Text11.Text)
Check1.Enabled = True
Text11.Enabled = True
Text11.Locked = False

If Val(Text11.Text) <= Val(Text6.Text) Then
Text7.Text = Val(Text6.Text) - Val(Text11.Text)
Else
'Text7.Text = ""
End If

'Text5.SetFocus
End Sub

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Check1.Value = True Then
'dakpa = Val(Text11.Text)
'Text11.Text = 0
'End If
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Check1.Value = True Then
'dakpa = Val(Text11.Text)
'Text11.Text = dakpa
'End If
End Sub

Private Sub Command1_Click()
'customer.Text8.Text = Text12.Text

If Option2.Value = True Then
customer2.Show
customer2.Text8.Text = Text12.Text
End If

If Option3.Value = True Then
suki2.Show
suki2.Text8.Text = Text12.Text
End If

End Sub

Private Sub Command3_Click()
transac_me
End Sub


Private Sub Command2_Click()
Unload Me
main.Show
main.Enabled = True
End Sub


Private Sub Form_Load()










Command1.Enabled = False
Check1.Enabled = False

Text3.Enabled = True
Text3.Locked = True
Text3.Text = "               --------------------------------------"


'If Option3.Value = True Then
'If Option1 = True Then
'Check1.Enabled = True
'Else: ' Check1.Enabled = False
'End If
'End If



'Text5.Text = ""
Text6.Text = ""
Text9.Text = ""
Text10.Text = ""
'Text7.Text = ""
Text4.Text = main.Text6.Text


If main.plus_card.Text = "" Then
Option1.Value = True
Else

Option3.Value = True
'Text12.Text = main.card_text.Text
'
End If

Text8.Text = "0"

If Option3.Value = True Then






Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs3 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs7 = New ADODB.Recordset
rs2.Open "SELECT CUSTOMER_ID FROM customer where CARD_NUMBER = '" & main.card_text.Text & "'", conn


Do Until rs2.EOF

For Each fld1 In rs2.Fields
Text12.Text = fld1.Value
Next

 rs2.MoveNext
 Loop
 
 
 Text1.Text = main.plus_card.Text & " with Control No.: " & Text12.Text & " and Card No.: " & main.card_text.Text
 
 rs7.Open "SELECT POINTS FROM suki_card where CARD_NUMBER = '" & Text12.Text & "'", conn
 
 
 
 Do Until rs7.EOF
 For Each fld1 In rs7.Fields
Text8.Text = Val(fld1.Value) + Val(Text8.Text)
Next
rs7.MoveNext
Loop


'rs3.Close

'rs1.Close
'rs2.Close

End If




End Sub

Private Sub Form_Unload(Cancel As Integer)

main.Enabled = True
main.Show
main.item_code.SetFocus
End Sub

Private Sub List1_DblClick()
Text12.Text = ""
Dim holder As String
holder = List4.List(List1.ListIndex)
Text1.Text = List1.List(List1.ListIndex)
Text12.Text = holder
Text3.Text = Text1.Text
List1.Visible = False
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text12.Text = ""
Dim holder As String
holder = List4.List(List1.ListIndex)
Text1.Text = List1.List(List1.ListIndex)
Text12.Text = holder
Text3.Text = Text1.Text
List1.Visible = False
End If
End Sub

Private Sub Option1_Click()
Command1.Enabled = False
reg = False
Text1.Enabled = False
Text1.Locked = True
Text1.Text = "               --------------"
Text3.Enabled = False
Text3.Locked = True
Text3.Text = "               --------------------------------------"
Check1.Enabled = False
Check1.Value = 0
Text11.Locked = True
Text11.Text = ""
Text8.Text = ""
Text5.Enabled = False
Text7.Enabled = False
Text5.Text = " ----------------------"
Text7.Text = " ----------------------"
End Sub

Private Sub Option2_Click()

Command1.Enabled = True
reg = True
Text1.Locked = False
Text1.Enabled = True
Text1.Text = ""
Check1.Enabled = False
Check1.Value = 0
Text1.Text = "Enter Customer Name"
Text3.Text = "Customer's Name"
Text11.Locked = True
Text11.Text = ""
Text8.Text = ""
Text5.Text = ""
'Text4.Text = ""
Text12.Text = ""
Text5.Enabled = True

'Text5.SetFocus

End Sub

Private Sub Option3_Click()


Command1.Enabled = True
Text11.Text = ""
Text8.Text = ""

Check1.Enabled = True
Text3.Text = "Customer Name & Crtl #"
Text1.Text = "Enter Customer Name"
Text3.Enabled = False
Text1.Enabled = True
Text7.Text = ""
Text5.Text = ""
Text5.Enabled = False
Text7.Enabled = True
Text1.Locked = False
Text5.Text = " ----------------------"
Text7.Text = " ----------------------"
End Sub

Private Sub Option3_Validate(Cancel As Boolean)
If Option1 = True Then
Check1.Enabled = True
Else: Check1.Enabled = False
End If
End Sub

Private Sub Text1_Change()
Dim val1 As String
Dim val2 As String
Dim val3 As String
List1.Clear
List2.Clear
List3.Clear
List4.Clear

If Text1.Text = "" Then
List1.Visible = False
End If
Text3.Enabled = True
Text8.Text = ""
'On Error GoTo sibat
'Text12.Text = Text1.Text
On Error GoTo sibat
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open


Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset

If Option2.Value = True Then
rs1.Open "SELECT NAME FROM customer where NAME like '" & Text1.Text & "%'", conn
rs2.Open "SELECT CUSTOMER_ID FROM customer where NAME like '" & Text1.Text & "%'", conn


Do Until rs1.EOF
For Each fld1 In rs1.Fields
List2.AddItem fld1.Value
val1 = fld1.Value
Next
For Each fld1 In rs2.Fields
List4.AddItem fld1.Value
val2 = fld1.Value
Next
rs1.MoveNext
rs2.MoveNext
List1.AddItem val1 & ", with Control No.: " & val2
Loop

'rs2.Open "SELECT CUSTOMER_ID FROM customer where CARD_NUMBER = '" & Text1.Text & "'", conn
End If


If Option3.Value = True Then

rs1.Open "SELECT NAME FROM customer where NAME like '" & Text1.Text & "%'", conn
rs2.Open "SELECT CUSTOMER_ID FROM customer where NAME like '" & Text1.Text & "%'", conn
rs3.Open "SELECT CARD_NUMBER FROM customer where NAME like '" & Text1.Text & "%'", conn

Do Until rs1.EOF

For Each fld1 In rs3.Fields

If fld1.Value = "unregister" Then
GoTo sunod_na
End If
List3.AddItem fld1.Value
val3 = fld1.Value
Next

For Each fld1 In rs1.Fields
List2.AddItem fld1.Value
val1 = fld1.Value
Next
For Each fld1 In rs2.Fields
List4.AddItem fld1.Value
val2 = fld1.Value
Next

List1.AddItem val1 & " with Control No.: " & val2 & " and Card No.: " & val3
sunod_na:


rs3.MoveNext
 rs1.MoveNext
 rs2.MoveNext

 Loop
 
 
 
 
Dim co As Integer
Dim adders As Double
co = 0
rs3.Open "SELECT POINTS FROM suki_card where CARD_NUMBER = '" & Text12.Text & "'", conn
'rs3.MoveFirst

Do Until rs3.EOF
co = co + 1
For Each fld1 In rs3.Fields
adders = adders + Val(fld1.Value)
'MsgBox co & " " & Text8.Text & " " & fld1.Value
Next

rs3.MoveNext
Loop
Text8.Text = adders
End If



If Option1.Value = True Then
Exit Sub
End If




sibat:

'Text3.Text = ""
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
List1.Visible = True

If KeyCode = 40 Then
List1.SetFocus
End If


If KeyCode = 8 Then
If Text1.Text = "" Then
reg = False
Text1.Enabled = False
Text1.Locked = True
Text1.Text = "Enter Card Number"
Text3.Enabled = False
Text3.Locked = True
Option1.Value = True
Option2.Value = False
End If
End If
If KeyCode = 27 Then
Unload Me
main.Enabled = True
main.Show
End If



End Sub

Private Sub Text11_Change()
'MsgBox Text6.Text
On Error GoTo h

Check1.Enabled = True

If old_val < Val(Text11.Text) And Check1.Value = 1 Then

Text11.Text = old_val



End If



If Val(Text11.Text) <= Val(Text6.Text) Then
Text7.Text = Val(Text6.Text) - Val(Text11.Text)
Else
'Text7.Text = ""
End If

h:
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)

If Check1.Value = 1 Then
Text11.Enabled = True
Else
Text11.Enabled = False
End If

If KeyCode = 13 Then
transac_me
End If
End Sub

Private Sub Text12_Change()


If Option3.Value = False Then
Exit Sub
End If


If Val(Text12.Text) <> 0 Then
Command1.Enabled = True
Text5.Enabled = False
'Text5.Locked = False
Else
Command1.Enabled = False
End If





Set conn = New ADODB.Connection
Set rs3 = New ADODB.Recordset
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open


Dim co As Integer
Dim adders As Double
co = 0
rs3.Open "SELECT POINTS FROM suki_card where CARD_NUMBER = '" & Text12.Text & "'", conn
'rs3.MoveFirst

Do Until rs3.EOF
co = co + 1
For Each fld1 In rs3.Fields
adders = adders + Val(fld1.Value)
'MsgBox co & " " & Text8.Text & " " & fld1.Value
Next

rs3.MoveNext
Loop
Text8.Text = adders



End Sub

Private Sub Text4_Change()
Dim aa() As String
Text6.Text = Val(text2.Text) - Val(Text4.Text)
Text6.Text = Val(Text6.Text) * 1.0000001
aa = Split(Text6.Text, ".")
Text6.Text = Left(Text6.Text, (Len(aa(0)) + 3))


Text10.Text = Val(Text9.Text) - Val(Text7.Text)
Text10.Text = Val(Text10.Text) * 1.0000001
aa = Split(Text10.Text, ".")
Text10.Text = Left(Text10.Text, (Len(aa(0)) + 3))
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
main.Enabled = True
main.Show
End If

End Sub

Private Sub Text5_Change()
Dim bayad As Integer
Text6.Text = Val(text2.Text) - Val(Text4.Text)
Text6.Text = Val(Text6.Text) * 1.0000001
aa = Split(Text6.Text, ".")
Text6.Text = Left(Text6.Text, (Len(aa(0)) + 3))


Text7.Text = Val(Text6.Text) - Val(Text5.Text)
Text7.Text = Val(Text7.Text) * 1.0000001
aa = Split(Text7.Text, ".")
Text7.Text = Left(Text7.Text, (Len(aa(0)) + 3))

Text10.Text = Val(Text9.Text) - Val(Text7.Text)
Text10.Text = Val(Text10.Text) * 1.0000001
aa = Split(Text10.Text, ".")
Text10.Text = Left(Text10.Text, (Len(aa(0)) + 3))
End Sub

Private Sub Text5_Click()
If Text1.Text <> "Enter Customer Name" Or Text1.Text <> "--------------" Then


bayad.Show


Else

Exit Sub
End If
End Sub

Private Sub Text5_GotFocus()
If Text13.Text = "" Then
transac.Enabled = False
bayad.Show
Else
Exit Sub
End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
main.Show
Exit Sub
End If
If KeyCode = 13 Then
transac_me
End If
End Sub

Private Sub Text6_Change()
'If Val(Text11.Text) >= Val(Text6.Text) Then
'Text7.Text = Val(Text6.Text) - Val(Text11.Text)
'Else
'Text7.Text = ""
'End If
End Sub

Private Sub Text7_Change()
sudli = Val(Text7.Text)
End Sub

Private Sub Text8_Change()
On Error GoTo h
Dim aaa() As String

aaa = Split(Text8.Text, ".")
Text8.Text = Left(Text8.Text, (Len(aaa(0)) + 3))
Dim aa() As String
Text11.Text = Val(Text8.Text) * 0.3

Text11.Text = Val(Text11.Text) * 1.0000001
aa = Split(Text11.Text, ".")
Text11.Text = Left(Text11.Text, (Len(aa(0)) + 3))

h:
End Sub

Private Sub Text9_Change()


If Check1.Value = 1 Then
'Exit Sub
End If

On Error GoTo pass
Dim bayad As Double
Dim bayara As Integer
Dim aa() As String

Text6.Text = Val(text2.Text) - Val(Text4.Text)
Text6.Text = Val(Text6.Text) * 1.0000001
aa = Split(Text6.Text, ".")
Text6.Text = Left(Text6.Text, (Len(aa(0)) + 3))

bayad = Val(Text9.Text)
Text10.Text = ""
'bayaran = Val(Text6.Text)

If Option2.Value = True And Text5.Text <> "" Then
Text10.Text = bayad - Val(Text7.Text)

Text10.Text = Val(Text10.Text) * 1.0000001
aa = Split(Text10.Text, ".")
Text10.Text = Left(Text10.Text, (Len(aa(0)) + 3))
Exit Sub
End If

If Check1.Value = 1 Then
bayad = bayad + Val(Text11.Text)
Text10.Text = bayad - Val(Text6.Text)
'Exit Sub
End If


Text10.Text = bayad - Val(Text6.Text)


Text10.Text = Val(Text10.Text) * 1.0000001
aa = Split(Text10.Text, ".")
Text10.Text = Left(Text10.Text, (Len(aa(0)) + 3))

pass:
Exit Sub
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 27 Then
Unload Me
main.Show
Exit Sub
End If
If KeyCode = 13 Then
transac_me
End If
End Sub

Private Sub transac_me()

Dim pambayad As Boolean
pambayad = False
Dim bayad1 As Double
Dim bayad2 As Double

bayad1 = Val(Text6.Text)
bayad2 = Val(Text9.Text)

i_save = Val(Text11.Text)
If Check1.Value = 1 And Option3.Value = True Then
GoTo sunod
End If

If Check1.Value = 1 Then
GoTo sunod
End If



If Val(Text9.Text) < Val(Text6.Text) And Text5.Text = "" Then
Exit Sub
End If

sunod:

If MsgBox("Press OK to end transaction", vbYesNo) = vbYes Then
Else
Exit Sub
End If

main.text2.Text = Text10.Text

If KeyCode = 27 Then
Unload Me
End If


main.Text5.Text = "0.00"
main.Text6.Text = "0.00"
main.Text7.Text = "0.00"
main.plus_card.Text = ""

Dim cat As String
Dim mods As String
 Dim cip As String
 Dim pitsa As String
 Dim sup As String
 Dim conn As ADODB.Connection
 
On Error GoTo mee
'main.grid2.Refresh
main.Enabled = True


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
Dim marg As String
Dim descrp As String
Dim qq As Integer
qq = 0

Dim urasan As String
urasan = Now
Dim ihap As Integer
ihap = 0
Dim utang_u As Double
Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs7.Open "SELECT BALANCE FROM utang WHERE CUSTOMER_ID = '" & Text12.Text & "'", conn


Do Until rs7.EOF
ihap = ihap + 1
rs7.MoveNext
Loop
If ihap = 0 Then
Else
rs7.MoveFirst
'MsgBox ihap
For zz = 1 To ihap - 1
rs7.MoveNext
Next
For Each fld1 In rs7.Fields
'MsgBox fld1.Value
utang_u = Val(fld1.Value)
Next
End If
For qq = 0 To Val(main.Text1.Text) + 3
'MsgBox qq
'MsgBox main.grid2.TextMatrix(qq, 1)
rs1.Open "SELECT CATEGORY FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn
rs2.Open "SELECT MODEL FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn
rs3.Open "SELECT CP FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn
rs4.Open "SELECT DATE_RECEIVED FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn
rs5.Open "SELECT SUPPLIER_NAME FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn
rs6.Open "SELECT MARGIN FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn
rs10.Open "SELECT DESCRIPTION FROM stock_info WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'", conn



Do Until rs1.EOF

For Each fld1 In rs1.Fields
cat = fld1.Value
Next

For Each fld1 In rs6.Fields
marg = fld1.Value
Next

For Each fld1 In rs2.Fields
mods = fld1.Value
Next

For Each fld1 In rs3.Fields
cip = fld1.Value
Next

For Each fld1 In rs4.Fields
pitsa = fld1.Value
Next

For Each fld1 In rs5.Fields
sup = fld1.Value
Next

For Each fld1 In rs10.Fields
descrp = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
rs6.MoveNext
rs10.MoveNext

Loop
Dim dis As String
Dim aaaa As Double
    If main.grid2.TextMatrix(1, qq) = "" Then
    Else
        dis = main.grid2.TextMatrix(qq, 5)
        aaaa = Val(main.grid2.TextMatrix(qq, 6)) + Val(main.grid2.TextMatrix(qq, 5))
 
    If Option1.Value = True Then
     conn.Execute "INSERT INTO sales (DATE_SOLD,CATEGORY, MODEL,DESCRIPTION,ITEM_CODE,PCS, RECEIVED, SUPPLIER, CP, RP, TOTAL, CASHER,INVOICE,MARGIN_PESO,DISCOUNT, GROSS)" _
        & "values ('" & Now & "', '" & cat & "', '" & mods & "','" & LTrim(main.grid2.TextMatrix(qq, 2)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 1)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 3)) & "', '" & pitsa & "', '" _
        & sup & "', '" & cip & "', '" & main.grid2.TextMatrix(qq, 4) & "', '" & main.grid2.TextMatrix(qq, 6) & "', '" & main.casher.Text & "', '" & main.invoice.Text & "', '" & main.grid2.TextMatrix(qq, 7) & "', '" & dis & "', '" & aaaa & "')"
conn.Execute "UPDATE stock_info SET STOCK_ON_HAND = " & "'" & main.List1.List(qq - 1) - main.grid2.TextMatrix(qq, 3) & "'" & " WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'"

End If
    
  

  

If Option2.Value = True Then
    
    
conn.Execute "UPDATE stock_info SET STOCK_ON_HAND = " & "'" & main.List1.List(qq - 1) - main.grid2.TextMatrix(qq, 3) & "'" & " WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'"
    
        If Val(Text9.Text) = 0 Then
       Else
        conn.Execute "INSERT INTO sales (DATE_SOLD,CATEGORY, MODEL,DESCRIPTION,ITEM_CODE,PCS, RECEIVED, SUPPLIER, CP, RP, TOTAL, CASHER,INVOICE,MARGIN_PESO,DISCOUNT, GROSS)" _
        & "values ('" & Now & "', '" & "-------" & "', '" & "--------" & "','" & "Payment of " & Text1.Text & "', '" & "--------" & "', '" & "--------" & "', '" & "--------" & "', '" _
        & "--------" & "', '" & "--------" & "', '" & "--------" & "', '" & "--------" & "', '" & main.casher.Text & "', '" & main.invoice.Text & "', '" & "--------" & "', '" & "--------" & "', '" & Text7.Text & "')"
End If
    


utang_u = utang_u + Val(main.grid2.TextMatrix(qq, 6))
sev = Val(Text7.Text)

conn.Execute "INSERT INTO utang (DESCRIPTION, CUSTOMER_ID, ITEM_CODE, QUANTITY, RP, CREDIT, DATE_SOLD,DUE_DATE, INVOICE, BALANCE)" _
    & "values ('" & descrp & "', '" & Text12.Text & "', '" & main.grid2.TextMatrix(qq, 1) & "', '" & main.grid2.TextMatrix(qq, 3) & "','" & main.grid2.TextMatrix(qq, 4) & "','" & main.grid2.TextMatrix(qq, 6) & "', '" & Now & "', '" & Text13.Text & "', '" & main.invoice.Text & "', '" & utang_u & "')"


    
End If

'  rs7.Close








If Option3.Value = True Then
If Check1.Value = 1 Then

'If Val(Text6.Text) > Val(Text11.Text) And Val(Text9.Text) = 0 Then
'Exit Sub
'End If



If Val(Text6.Text) > (Val(Text11.Text) + Val(Text9.Text)) Then
MsgBox "Insufficient points to close the transaction."
Exit Sub
End If

'MsgBox "points"
  conn.Execute "INSERT INTO sales (DATE_SOLD,CATEGORY, MODEL,DESCRIPTION,ITEM_CODE,PCS, RECEIVED, SUPPLIER, CP, RP, TOTAL, CASHER,INVOICE,MARGIN_PESO,DISCOUNT, GROSS)" _
        & "values ('" & Now & "', '" & cat & "', '" & mods & "','" & LTrim(main.grid2.TextMatrix(qq, 2)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 1)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 3)) & "', '" & pitsa & "', '" _
        & sup & "', '" & cip & "', '" & main.grid2.TextMatrix(qq, 4) & "', '" & main.grid2.TextMatrix(qq, 6) & "', '" & main.casher.Text & "', '" & main.invoice.Text & "', '" & main.grid2.TextMatrix(qq, 7) & "', '" & dis & "', '" & "--Suki Card--" & "')"


conn.Execute "INSERT INTO suki_card (DATE_SOLD,DESCRIPTION,ITEM_CODE,QUANTITY, PRICE, AMOUNT,POINTS, CARD_NUMBER)" _
& "values ('" & Now & "','" & LTrim(main.grid2.TextMatrix(qq, 2)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 1)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 3)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 4)) & "', '" _
& main.grid2.TextMatrix(qq, 6) & "', '" & Val(main.grid2.TextMatrix(qq, 7)) * 0.15 & "', '" & Text12.Text & "')"
  
End If '' -------  end for check




If Check1.Value = 0 Then

'MsgBox "cash"
conn.Execute "INSERT INTO suki_card (DATE_SOLD,DESCRIPTION,ITEM_CODE,QUANTITY, PRICE, AMOUNT,POINTS, CARD_NUMBER)" _
& "values ('" & Now & "','" & LTrim(main.grid2.TextMatrix(qq, 2)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 1)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 3)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 4)) & "', '" _
& main.grid2.TextMatrix(qq, 6) & "', '" & Val(main.grid2.TextMatrix(qq, 7)) * 0.15 & "', '" & Text12.Text & "')"
  

  
  conn.Execute "INSERT INTO sales (DATE_SOLD,CATEGORY, MODEL,DESCRIPTION,ITEM_CODE,PCS, RECEIVED, SUPPLIER, CP, RP, TOTAL, CASHER,INVOICE,MARGIN_PESO,DISCOUNT, GROSS)" _
        & "values ('" & Now & "', '" & cat & "', '" & mods & "','" & LTrim(main.grid2.TextMatrix(qq, 2)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 1)) & "', '" & LTrim(main.grid2.TextMatrix(qq, 3)) & "', '" & pitsa & "', '" _
        & sup & "', '" & cip & "', '" & main.grid2.TextMatrix(qq, 4) & "', '" & main.grid2.TextMatrix(qq, 6) & "', '" & main.casher.Text & "', '" & main.invoice.Text & "', '" & main.grid2.TextMatrix(qq, 7) & "', '" & dis & "', '" & aaaa & "')"

'sMsgBox "cash 2"
End If  '-------end for uncheck




conn.Execute "UPDATE stock_info SET STOCK_ON_HAND = " & "'" & main.List1.List(qq - 1) - main.grid2.TextMatrix(qq, 3) & "'" & " WHERE ITEM_CODE = '" & main.grid2.TextMatrix(qq, 1) & "'"
'MsgBox main.List1.List(qq - 1) - main.grid2.TextMatrix(qq, 3)
End If






  

End If




  rs1.Close
  rs2.Close
  rs3.Close
  rs4.Close
  rs5.Close
  rs6.Close
  rs10.Close

'MsgBox qq + qq
Next


'DELETED NOT TO--------


For a = 1 To Val(main.Text1.Text) - 1
main.grid2.RemoveItem (a)
main.Refresh
Unload Me
main.Enabled = True
Next

'main.Enabled = True
'Unload Me

For a = 1 To Val(main.Text1.Text) - 1
main.grid2.RemoveItem (a)

Next


mee:

main.List1.Clear
Dim kk() As String
Dim kk1() As String


Text6.Text = Val(Text6.Text) / 0.03
kk = Split(Text6.Text, ".")
Text6.Text = Left(Text6.Text, Len(kk(0)) + 3)

'MsgBox "panso"
If Option3.Value = True And Check1.Value = 1 Then

conn.Execute "INSERT INTO suki_card (DATE_SOLD,POINTS, CARD_NUMBER)" _
& "values ('" & Now & "', '-" & i_save / 0.3 & "', '" & Text12.Text & "')"

'conn.Execute "INSERT INTO sales (DATE_SOLD, CASHER,INVOICE,DISCOUNT, GROSS,DESCRIPTION)" _
'        & "values ('" & Now & "', '" & main.casher.Text & "', '" & main.invoice.Text & "', '" & "--Suki Card--" & "', '" & sudli _
 '       & "', '" & "Suki Card ----- " & Text1.Text & "')"


main.Text1.Text = "1"
main.invoice.Text = Val(main.invoice) + 1
main.enter.Enabled = False
Unload Me
main.Enabled = True
Exit Sub

End If



If Option2.Value = True Then
'MsgBox Text7.Text
If Val(Text9.Text) = 0 Then
Else
conn.Execute "INSERT INTO utang (DATE_SOLD, CUSTOMER_ID, BALANCE, DEBIT, INVOICE, DESCRIPTION )" _
& "values ('" & Now & "','" & Text12.Text & "', '" & utang_u - sev & "', '" & sev & "', '" & main.invoice.Text & "', '" & "Credit payment of " & Text1.Text & "')"
End If
    
    
    main.Text1.Text = "1"
main.invoice.Text = Val(main.invoice) + 1
main.enter.Enabled = False
main.Refresh
Unload Me
main.Enabled = True
Exit Sub


End If

'MsgBox Err.Description
main.Text1.Text = "1"
main.invoice.Text = Val(main.invoice) + 1
main.enter.Enabled = False
main.Refresh
Unload Me
main.Enabled = True
Exit Sub

End Sub
