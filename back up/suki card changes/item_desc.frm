VERSION 5.00
Begin VB.Form item_desc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "item_desc.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1560
      TabIndex        =   20
      Text            =   "Text11"
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox cp 
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Text            =   "Text11"
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Text            =   "Text10"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   525
      Left            =   7080
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   600
      TabIndex        =   16
      Text            =   "Text8"
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Text7"
      Top             =   3195
      Width           =   2535
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   1
      Text            =   "Text6"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000008&
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   2400
      Width           =   2655
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
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   2400
      Width           =   2175
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
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   1680
      Width           =   2655
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
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "&Accept"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
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
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price:"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "In Stock"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Item  Code"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "item_desc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long
Private Sub Command1_Click()
Unload Me
find_item.Enabled = True
find_item.Show
End Sub

Private Sub Command2_Click()

transac

End Sub

Private Sub Form_Load()
Text4.Text = ""
Text6.Text = "0.00"

'Val(Text3.Text) = Val(Text3.Text)
'If Val(Text3.Text) < Val(Text4.Text) Or Val(Text3.Text) = 0 Then

'Text3.Text = "Out of Stock"
'Text3.BackColor = &HC0&
'Text3.ForeColor = &HFFFF&
'Else
'Text3.Text = Text9.Text
'Text3.BackColor = &H80000005
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
Unload Me
main.Enabled = True
main.item_code.SetFocus
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Text4_Change()
Dim aaa() As String
Dim r As Double
Text7.Text = (Val(Text5) * Val(Text4))
Text7.Text = Text7.Text * 1.00000001
aaa = Split(Text7.Text, ".")
Text7.Text = Left(Text7.Text, Len(aaa(0)) + 3)

If Val(Text3.Text) < Val(Text4.Text) Then

Text3.Text = "Out of Stock"
Text3.BackColor = &HC0&
Text3.ForeColor = &HFFFF&
Else
Text3.Text = Text9.Text
Text3.BackColor = &H80000005
End If

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
find_item.Enabled = True
find_item.Show
End If

If KeyCode = 13 Then
If Text4.Text = "" Then
Exit Sub
End If
transac
End If
If Text9.Text = "" Then
Text9.Text = Text3.Text
End If
Dim aaa() As String
Dim r As Double
Text7.Text = (Val(Text5) * Val(Text4))
Text7.Text = Text7.Text * 1.00000001
aaa = Split(Text7.Text, ".")
Text7.Text = Left(Text7.Text, Len(aaa(0)) + 3)

If Val(Text3.Text) < Val(Text4.Text) Then

Text3.Text = "Out of Stock"
Text3.BackColor = &HC0&
Text3.ForeColor = &HFFFF&
Else
Text3.Text = Text9.Text
Text3.BackColor = &H80000005
End If


End Sub

Private Sub Text6_Change()
On Error GoTo agoy

Text7.Text = ((Text4.Text * Text5.Text) - Val(Text6))
Text7.Text = Text7.Text * 1.00000001
aaa = Split(Text7.Text, ".")
Text7.Text = Left(Text7.Text, Len(aaa(0)) + 3)
agoy:
End Sub

Private Sub Text6_Click()
Text6.Text = ""
End Sub

Private Sub Text6_GotFocus()
Text6.Text = ""
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
find_item.Enabled = True
find_item.Show
End If
If KeyCode = 13 Then
If Text4.Text = "" Then
Exit Sub
End If
transac
End If
End Sub

Private Sub Text7_Change()
If Text7.Text = "0" Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
End Sub

Private Sub transac()

'On Error GoTo pugsit
Dim w As Double
If Text4.Text = "" Then
Exit Sub
End If
If Text3.Text = "Out of Stock" Then
Exit Sub
End If

Dim aas As Integer
Dim ssss As Double
Dim kk() As String
Dim sw As Integer


For aas = 0 To Val(main.Text1.Text) - 1
If main.grid2.TextMatrix(aas, 1) = text2.Text Then
main.grid2.TextMatrix(aas, 3) = Val(Text4.Text) + Val(main.grid2.TextMatrix(aas, 3))
main.grid2.TextMatrix(aas, 5) = Val(Text6.Text) + Val(main.grid2.TextMatrix(aas, 5))
Y = main.Text6.Text
main.Text6.Text = Val(Y) + Val(Text6.Text)



w = main.Text5.Text


sw = Val(w) + (Val(Text4.Text) * Val(Text5.Text))

sw = sw * 1.0000001
Text8.Text = sw

aa = Split(sw, ".")
main.Text5.Text = Left(Text8.Text, Len(aa(0)) + 3)


ssss = (Val(main.grid2.TextMatrix(aas, 3)) * Val(main.grid2.TextMatrix(aas, 4))) - Val(main.grid2.TextMatrix(aas, 5))
ssss = ssss * 1.000001
Text11.Text = ssss
kk = Split(Text11.Text, ".")
Text11.Text = Left(Text11.Text, Len(kk(0)) + 3)
main.grid2.TextMatrix(aas, 6) = Text11.Text & "      "
'main.Text5.Text = Val(main.grid2.TextMatrix(aas, 6)) + Val(main.grid2.TextMatrix(aas, 5))




GoTo pugsit
End If
Next









main.List1.AddItem Text3.Text
'im tig_ihap1 As Integer
tig_ihap1 = Val(main.Text1.Text)
'If Val(main.Text1.Text) = 9 Then
'main.grid2.Rows = Val(main.Text1.Text) + 1
'End If
main.grid2.Rows = tig_ihap1 + 1
'main.grid2.TextMatrix(tig_ihap1, 0) = tig_ihap1 + 1
main.grid2.TextMatrix(tig_ihap1, 1) = text2.Text
main.grid2.TextMatrix(tig_ihap1, 2) = Text1.Text
main.grid2.TextMatrix(tig_ihap1, 3) = Text4.Text
main.grid2.TextMatrix(tig_ihap1, 4) = Text5.Text
main.grid2.TextMatrix(tig_ihap1, 5) = Text6.Text
main.grid2.TextMatrix(tig_ihap1, 6) = Text7.Text & "      "
main.grid2.TextMatrix(tig_ihap1, 7) = ((Val(Text5.Text) - Val(cp.Text)) * Val(Text4.Text))

main.Text3.Text = (Val(Text10.Text) * Val(Text4.Text)) + Val(main.Text3.Text)
Dim qqq As Integer
'qqq = (Val(cp.Text) * Val(Text4.Text))
main.Text4.Text = Val(main.Text4.Text) + ((Val(Text5.Text) - Val(cp.Text)) * Val(Text4.Text))


main.Text1.Text = tig_ihap1
'main.grid2.TextMatrix(tig_ihap1, 5) = tig_ihap1
'Dim aa() As String
Y = main.Text6.Text
main.Text6.Text = Val(Y) + Val(Text6.Text)


w = main.Text5.Text
w = Val(w) + (Val(Text4.Text) * Val(Text5.Text))
w = w * 1.0000001
Text8.Text = w
aa = Split(w, ".")
main.Text5.Text = Left(Text8.Text, Len(aa(0)) + 3)

main.item_code.Text = ""


Unload Me
Unload find_item
main.Enabled = True
main.Show
tig_ihap1 = tig_ihap1 + 1
main.Text1.Text = tig_ihap1
pugsit:

main.item_code.Text = ""

main.Enabled = True
main.Show
Unload find_item
Unload Me
End Sub
