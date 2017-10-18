VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RIC'z Cycle Parts and Accessories Center POS System"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15225
   ForeColor       =   &H000000FF&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "main.frx":1C64C
   ScaleHeight     =   8115
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Pay &Credit F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox card_text 
      Height          =   405
      Left            =   2520
      TabIndex        =   41
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Left            =   1680
      TabIndex        =   36
      Top             =   6120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3120
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Log Out/Q"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FFFF&
      Caption         =   "&User Option/F2"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox search 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Text            =   "er"
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   31
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox plus_card 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Text            =   "Text8"
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   2280
   End
   Begin VB.CommandButton enter 
      BackColor       =   &H008080FF&
      Caption         =   "&Pay/ Enter"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton f6 
      BackColor       =   &H008080FF&
      Caption         =   " &Return  F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton f5 
      BackColor       =   &H008080FF&
      Caption         =   "&Add Client F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton f4 
      BackColor       =   &H008080FF&
      Caption         =   "   &Delete   F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton f3 
      BackColor       =   &H008080FF&
      Caption         =   "&Find Item F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1245
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   6480
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text6"
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Frame z 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   14535
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Left            =   12600
         TabIndex        =   27
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Discount"
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
         Left            =   11400
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Left            =   10080
         TabIndex        =   25
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Quantity"
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
         Left            =   9000
         TabIndex        =   14
         Top             =   0
         Width           =   975
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
         TabIndex        =   13
         Top             =   0
         Width           =   6615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      Begin VB.Label Label4 
         Caption         =   "Logo"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.TextBox item_code 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox invoice 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "00987"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox casher 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "dfdsf"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Image Image8 
      Height          =   915
      Left            =   13320
      Picture         =   "main.frx":27E81
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1170
   End
   Begin VB.Image Image7 
      Height          =   915
      Left            =   11880
      Picture         =   "main.frx":2A473
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1170
   End
   Begin VB.Image Image6 
      Height          =   1515
      Left            =   8160
      Picture         =   "main.frx":2C855
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2010
   End
   Begin VB.Image Image5 
      Height          =   915
      Left            =   6600
      Picture         =   "main.frx":2F1E4
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Image Image4 
      Height          =   915
      Left            =   5040
      Picture         =   "main.frx":31682
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   915
      Left            =   3480
      Picture         =   "main.frx":338C3
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   1920
      Picture         =   "main.frx":35DAB
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   360
      Picture         =   "main.frx":37F0A
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1440
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Change"
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
      TabIndex        =   37
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Suki Card:"
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
      Left            =   6840
      TabIndex        =   30
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time and time And time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   12600
      TabIndex        =   28
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   19
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Dicount:  "
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code:"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cashier:"
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
      Left            =   6960
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sulod As Integer
Public urasan As String
Public to_del As Integer

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


Private Sub casher_KeyDown(KeyCode As Integer, Shift As Integer)

listener (KeyCode)
End Sub



Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()


End Sub

Private Sub Command1_Click()
main.Enabled = False
'Unload item_desc
If MsgBox("Do you want to exit?", vbYesNo, "Ric's CyclePart and accessories Center") = vbYes Then
Unload Me
Else
'Unload item_desc
main.item_code.Text = ""
main.Show
main.Enabled = True
item_code.SetFocus
'Unload item_desc
Exit Sub
End If
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub Command2_Click()
main.Enabled = False
pay.Show
End Sub

Private Sub Command6_Click()
user.Show
Me.Enabled = False
End Sub


Private Sub Command6_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

'Public tig_ihap As Integer

Private Sub enter_Click()

transac.Show
transac.Text2.Text = Text5.Text
transac.Text4.Text = Text6.Text
transac.Text6.Text = Text7.Text

main.Enabled = False



End Sub

Private Sub enter_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub f3_Click()
find_item.Show
Me.Enabled = False

End Sub



Private Sub f3_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub f4_Click()
If Val(Text5.Text) > 0 Then
pad_on
End If

End Sub

Private Sub f4_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub f5_Click()
Me.Enabled = False
addclient.Show
End Sub

Private Sub f5_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub f6_Click()
Me.Enabled = False
returnme.Show
End Sub

Private Sub f6_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub Form_Load()


On Error GoTo hi
Text7.Text = "0.00"
Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Dim fld As ADODB.Field

conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT start FROM invoice", conn
rs2.Open "SELECT invoice FROM sales", conn
rs3.Open "SELECT end FROM invoice", conn



Do Until rs2.EOF
For Each fld In rs2.Fields
bx = Val(fld.Value)
Next
rs2.MoveNext
Loop


Do Until rs1.EOF
For Each fld In rs1.Fields
ax = Val(fld.Value)
Next



For Each fld In rs3.Fields
cx = Val(fld.Value)
Next

rs1.MoveNext
rs3.MoveNext
Loop

If bx <= cx And bx >= ax Then
invoice.Text = bx + 1

Else

invoice.Text = ax

End If






'For x = 0 To 6
grid2.ColAlignment(1) = 1
'Next


aaa = aaa * 0.00000001

Text5.Text = "0.00"
Text6.Text = "0.00"
Text7.Text = "0.00"
grid2.RowHeight(0) = 1
grid2.ColWidth(0) = 1
grid2.ColWidth(1) = 2200
grid2.ColWidth(2) = 6700
grid2.ColWidth(3) = 1090
grid2.ColWidth(4) = 1300
grid2.ColWidth(5) = 1200
grid2.ColWidth(6) = 2000

hi:
'MsgBox Err.Description
Exit Sub


End Sub

Private Sub Form_Resize()
'On Error GoTo hi
'z.Width = 12735
'z.Height = 3615
'Label9.Width = 2055
'Label9.Height = 375
'Label10.Width = 4815
'Label10.Height = 375
'Label11.Width = 975
'Label11.Height = 375
'Dim diff1 As Integer
'Dim diff2 As Integer
'diff1 = main.Width - 13500
'diff2 = main.Height - 9030

'z.Height = z.Height + diff2
'z.Width = z.Width + diff1
'Label9.Height = Label9.Height + diff2
'Label9.Width = Label9.Width + diff1
'Label10.Height = Label10.Height + diff2
'Label10.Width = Label10.Width + diff1
'Label11.Height = Label11.Height + diff2
'Label11.Width = Label11.Width + diff1


'hi:
'Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Enabled = False
'Unload item_desc
If MsgBox("Do you want to exit?", vbYesNo, "Ric's CyclePart and accessories Center") = vbYes Then
Unload Me
Else
'Unload item_desc
main.item_code.Text = ""
main.Show
main.Enabled = True
item_code.SetFocus
'Unload item_desc
Exit Sub
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
f3.Value = True
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
f4.Value = True
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
f5.Value = True
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
f6.Value = True
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
Command2.Value = True
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = False
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If enter.Enabled = False Then
Image6.Enabled = False
Else
Image6.Enabled = True
End If
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
enter.Value = True
End Sub

Private Sub Image7_Click()
Command1.Value = True
End Sub

Private Sub Image8_Click()
Command6.Value = True
End Sub

Private Sub invoice_DblClick()
main.Enabled = False
invoice_f.Show
End Sub

Private Sub List1_Click()
sulod = List1.ListIndex
End Sub


Private Sub grid2_Click()
On Error GoTo dakop
to_del = (grid2.Row)
dakop:
'MsgBox Err.Description
Exit Sub

End Sub

Private Sub grid2_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox "sss"



listener (KeyCode)
If KeyCode = 115 Then
If Val(Text5.Text) > 0 Then
If Val(Text5.Text) > 0 Then
pad_on
End If
End If
End If

If KeyCode = 27 Then
item_code.SetFocus
End If






End Sub

Private Sub invoice_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub



Private Sub item_code_KeyDown(KeyCode As Integer, Shift As Integer)

listener (KeyCode)
Exit Sub
End Sub


Private Sub List2_DblClick()
card_text.Text = List4.List(List2.ListIndex)
plus_card.Text = List3.List(List2.ListIndex)
 
List2.Visible = False
item_code.SetFocus
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 13 Then
card_text.Text = List4.List(List2.ListIndex)
plus_card.Text = List3.List(List2.ListIndex)
item_code.SetFocus
List2.Visible = False
End If
End Sub

Private Sub Picture1_Click()
Text6.SetFocus
End Sub

Private Sub puls_card_Change()



End Sub

Private Sub puls_card_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub



Private Sub plus_card_Change()

List2.Width = 4935
List2.Height = 1410
List2.Top = 2040
List2.Left = 8040


List2.Visible = True
List2.Clear
List3.Clear
List4.Clear
If plus_card.Text = "" Then

List2.Visible = False

End If



Dim conn As ADODB.Connection
Dim fld As ADODB.Field
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim na As String
Dim add As String
Dim card As String
Dim numb As String


Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset

conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

rs1.Open "SELECT NAME FROM customer where CARD_NUMBER like '" & plus_card.Text & "%'", conn
rs2.Open "SELECT ADDRESS FROM customer where CARD_NUMBER like '" & plus_card.Text & "%'", conn
rs3.Open "SELECT CARD_NUMBER FROM customer where CARD_NUMBER like '" & plus_card.Text & "%'", conn
rs4.Open "SELECT NUMBER FROM customer where CARD_NUMBER like '" & plus_card.Text & "%'", conn

Do Until rs1.EOF
On Error GoTo hi
For Each fld In rs1.Fields
na = fld.Value
Next

For Each fld In rs2.Fields
add = fld.Value
Next

For Each fld In rs3.Fields
card = fld.Value
Next

For Each fld In rs4.Fields
numb = fld.Value
Next


List2.AddItem card & ",     " & na & ",     " & add & ",     " & numb
List3.AddItem na
List4.AddItem card


rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext



Loop


'List2.AddItem card & ",     " & na & ",     " & add & ",     " & numb
'List3.AddItem na
'List4.AddItem card
hi:
End Sub




Private Sub plus_card_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
List2.SetFocus
End If
listener (KeyCode)
End Sub

Private Sub search_Change()
item_desc.Show
main.Enabled = False
End Sub

Private Sub Text1_Change()
grid2.Rows = Val(main.Text1.Text)
End Sub



Private Sub Text5_Change()

Text7.Text = Val(Text5.Text) - Val(Text6.Text)
Text7.Text = Val(Text5.Text) - Val(Text6.Text)
Text7.Text = Val(Text7.Text) * 1.000001
aaa = Split(Text7.Text, ".")
Text7.Text = Left(Text7.Text, Len(aaa(0)) + 3)
End Sub

Private Sub Text6_Change()
Dim aaa() As String
Dim bb As Double
Text7.Text = Val(Text5.Text) - Val(Text6.Text)
Text7.Text = Val(Text7.Text) * 1.000001
aaa = Split(Text7.Text, ".")
Text7.Text = Left(Text7.Text, Len(aaa(0)) + 3)

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
listener (KeyCode)
End Sub

Private Sub Text7_Change()
If Text7.Text = "0" Then
enter.Enabled = False
Else
enter.Enabled = True
End If



End Sub



Private Sub Text9_Click()
grid2.TextMatrix(Val(Text10.Text), 6) = Text9.Text
MsgBox Val(Text9.Text)
End Sub

Private Sub Timer1_Timer()
Dim today As Variant
today = Format(Now, " hh : mm : ss ampm" & vbNewLine & "d  mmmm, yyyy" & vbNewLine & "dddd")
lbltime.Caption = today
If enter.Enabled = True Then
Image6.Enabled = True
End If

If enter.Enabled = False Then
Image6.Enabled = False
End If
End Sub

Public Sub listener(keyhit As Integer)

'f1
If keyhit = 112 Then
help_.Show
End If


'f2
If keyhit = 113 Then
user.Show
Me.Enabled = False
End If

'f3
If keyhit = 114 Then
find_item.Show
Me.Enabled = False
End If

'f4

If KeyCode = 115 Then
If Val(Text5.Text) > 0 Then
pad_on
End If
End If


'f5
If keyhit = 116 Then
Me.Enabled = False
addclient.Show
End If

'f6
If keyhit = 117 Then
Me.Enabled = False
returnme.Show
End If

'f8 upload stock
If keyhit = 119 Then
Shell ("uploader.exe")
End If

'f11 po
If keyhit = 122 Then
Shell ("po-extracto.exe")
End If

'f12

If keyhit = 123 Then
upgrade.Show
End If

'f10
If keyhit = 121 Then
Shell ("extractor.exe")
End If

'f7

If keyhit = 118 Then
main.Enabled = False
pay.Show
End If

'space
If keyhit = 32 Then
tabb
End If

'enter5
If keyhit = 13 Then



If Val(Text5.Text) = 0 Then
Exit Sub
Else

enter.Value = True
Exit Sub


transac.Show
transac.Text2.Text = Text5.Text
transac.Text4.Text = Text6.Text
transac.Text6.Text = Text7.Text
main.Enabled = False
Exit Sub
End If

End If

'f2
If keyhit = 113 Then
user.Show
'Text6.SetFocus
End If

'F9
If keyhit = 120 Then
history2.Show
'main.Enabled = False
End If

'esc
If keyhit = 27 Then
search.Visible = False
item_code.SetFocus
End If

'q
If keyhit = 81 Then
main.Enabled = False
'Unload item_desc
If MsgBox("Do you want to exit?", vbYesNo, "Ric's CyclePart and accessories Center") = vbYes Then
Unload Me
Else
'Unload item_desc
main.item_code.Text = ""
main.Show
main.Enabled = True
item_code.SetFocus
'Unload item_desc
Exit Sub
End If
End If

End Sub


Private Sub pad_on()
'MsgBox grid2.Row
On Error GoTo errrr

Text6.Text = Val(Text6.Text) - Val(grid2.TextMatrix(Val(grid2.Row), 5))
Text4.Text = Val(Text4.Text) - Val(grid2.TextMatrix(Val(grid2.Row), 7))
'grid2.Redraw
Text5.Text = Val(Text5.Text) - (Val(grid2.TextMatrix(grid2.Row, 4))) * (Val(grid2.TextMatrix(grid2.Row, 3)))
grid2.RemoveItem grid2.Row
Text1.Text = Val(Text1.Text) - 1
grid2.Rows = Val(main.Text1.Text)
'MsgBox grid2.Rows
grid2.Refresh
errrr:
'MsgBox Err.Description
Exit Sub

End Sub

Private Sub tabb()

Dim fld1 As ADODB.Field

On Error GoTo outme


Dim looper As Integer

Dim a1 As String
Dim ax As String
Dim a2 As String
Dim a3 As String
Dim a4 As String
Dim a5 As String
Dim a6 As String

Set conn = New ADODB.Connection
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset
Set rs4 = New ADODB.Recordset
Set rs5 = New ADODB.Recordset
Set rs6 = New ADODB.Recordset

rs1.Open "SELECT RP FROM stock_info where ITEM_CODE = " & "'" & item_code.Text & "'", conn
rs2.Open "SELECT DESCRIPTION FROM stock_info where ITEM_CODE = " & "'" & item_code & "'", conn
rs3.Open "SELECT MODEL FROM stock_info where ITEM_CODE = " & "'" & item_code & "'", conn
rs4.Open "SELECT ITEM_CODE FROM stock_info where ITEM_CODE = " & "'" & item_code & "'", conn
rs5.Open "SELECT STOCK_ON_HAND FROM stock_info where ITEM_CODE = " & "'" & item_code & "'", conn



Do Until rs1.EOF
search.Text = ""

For Each fld1 In rs1.Fields
item_desc.Text5 = fld1.Value
search.Text = fld1.Value
Next


For Each fld1 In rs2.Fields
item_desc.Text1 = fld1.Value
Next

For Each fld1 In rs3.Fields
item_desc.Text1 = item_desc.Text1.Text & ",  " & fld1.Value
Next

For Each fld1 In rs4.Fields
item_desc.Text2 = fld1.Value
Next

For Each fld1 In rs5.Fields
item_desc.Text3 = fld1.Value
Next

rs1.MoveNext
rs2.MoveNext
rs3.MoveNext
rs4.MoveNext
rs5.MoveNext
Loop

outme:
Exit Sub

End Sub
