VERSION 5.00
Begin VB.Form pay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Form"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "pay.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "&Save/Enter"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&Cancel/Esc"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text8 
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
      Left            =   1560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text5 
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
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text4 
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
      Left            =   7200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2295
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
      Height          =   405
      Left            =   7200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2295
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
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3135
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
      Left            =   1560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   840
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Left            =   7080
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "pay"
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



Private Sub Command1_Click()
main.Show
main.Enabled = True
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Command2_Click()
enter_me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Form_Unload(Cancel As Integer)

main.Show
main.Enabled = True
Unload Me
main.item_code.SetFocus
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub Text5_Change()

On Error GoTo agoy


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

rs7.Open "SELECT BALANCE FROM utang WHERE CUSTOMER_ID = '" & Text5.Text & "'", conn

ihap = 0
Do Until rs7.EOF
ihap = ihap + 1
rs7.MoveNext
Loop
'If ihap <> 1 Then
'Exit Sub
'End If

rs7.MoveFirst
'MsgBox ihap
For zz = 1 To ihap - 1
rs7.MoveNext
Next
For Each fld1 In rs7.Fields
'MsgBox fld1.Value
Text8.Text = Val(fld1.Value)
Next


rs1.Open "SELECT CARD_NUMBER FROM customer WHERE CUSTOMER_ID = '" & Text5.Text & "'", conn
rs2.Open "SELECT NAME FROM customer WHERE CUSTOMER_ID = '" & Text5.Text & "'", conn
rs3.Open "SELECT ADDRESS FROM customer WHERE CUSTOMER_ID = '" & Text5.Text & "'", conn
rs4.Open "SELECT NUMBER FROM customer WHERE CUSTOMER_ID = '" & Text5.Text & "'", conn

'rs6.Open "SELECT CREDIT FROM customer WHERE CUSTOMER_ID = '" & Text4.Text & "'", conn
'rs7.Open "SELECT BAL FROM customer WHERE CUSTOMER_ID = '" & Text4.Text & "'", conn

For Each fld In rs1.Fields
Text4.Text = fld.Value
Next

For Each fld In rs2.Fields
Text1.Text = fld.Value
Next

For Each fld In rs3.Fields
text2.Text = fld.Value
Next

For Each fld In rs4.Fields
Text3.Text = fld.Value
Next


'For Each fld In rs6.Fields
'Text6.Text = fld.Value
'Next

'For Each fld In rs7.Fields
'Text8.Text = fld.Value
'Next

'Text7.Text = Val(Text6.Text) - Val(Text8.Text)

agoy:
Exit Sub
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
paminaw (KeyCode)
End Sub

Private Sub paminaw(keyhit As Integer)

If keyhit = 27 Then
main.Show
main.Enabled = True
Unload Me
End If

If keyhit = 13 Then

enter_me
End If

End Sub


Private Sub enter_me()
'On Error GoTo exit_me
If Val(Text8.Text) = 0 Then
MsgBox Text1.Text & " has fully paid his/her acount."
Exit Sub
End If
If Val(Text10.Text) = 0 Then
Exit Sub
End If


Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

conn.Execute "INSERT INTO utang (DATE_SOLD, CUSTOMER_ID, BALANCE, DEBIT, INVOICE )" _
& "values ('" & Now & "','" & Text5.Text & "', '" & Val(Text8.Text) - Val(Text10.Text) & "', '" & Text10.Text & "', '" & main.invoice.Text & "')"

    conn.Execute "INSERT INTO sales (DATE_SOLD,CATEGORY, MODEL,DESCRIPTION,ITEM_CODE,PCS, RECEIVED, SUPPLIER, CP, RP, TOTAL, CASHER,INVOICE,MARGIN_PESO,DISCOUNT, GROSS)" _
        & "values ('" & Now & "', '" & "-------" & "', '" & "--------" & "','" & "Payment of " & Text1.Text & "', '" & "--------" & "', '" & "--------" & "', '" & "--------" & "', '" _
        & "--------" & "', '" & "--------" & "', '" & "--------" & "', '" & "--------" & "', '" & main.casher.Text & "', '" & main.invoice.Text & "', '" & "--------" & "', '" & "--------" & "', '" & Text10.Text & "')"
       
main.Show
main.Enabled = True
Unload Me

exit_me:
Exit Sub

End Sub
