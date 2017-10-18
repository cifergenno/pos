VERSION 5.00
Begin VB.Form bayad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Text            =   "YYYY"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Text            =   "DD"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "bayad.frx":0000
      Left            =   960
      List            =   "bayad.frx":000D
      TabIndex        =   6
      Text            =   "MM"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Due Date:"
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
      Begin VB.TextBox mm 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Text            =   "mm"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox dd 
         Height          =   405
         Left            =   840
         TabIndex        =   1
         Text            =   "dd"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox yy 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Text            =   "yyyy"
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   600
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   1560
         X2              =   1320
         Y1              =   240
         Y2              =   600
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Credit Due Date"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "bayad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mm.Text = Combo1.Text
dd.Text = Combo2.Text
yy.Text = Combo3.Text
transac.Text13.Text = mm & "/" & dd & "/" & yy
'MsgBox mm & "/" & dd & "/" & yy
transac.Enabled = True
transac.Show
transac.Text5.SetFocus
transac.Text5.Enabled = True
transac.Text5.Text = ""
Unload Me

End Sub

Private Sub dd_GotFocus()
dd.Text = ""
End Sub

Private Sub dd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
transac.Show
End If
End Sub

Private Sub Form_Load()
 For aa = 1 To 12
Combo1.AddItem aa
Next
 For aa = 1 To 31
 Combo2.AddItem aa
 Next
 For aa = 2008 To 2030
 Combo3.AddItem aa
 Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
transac.Enabled = True
transac.Show
End Sub

Private Sub mm_Click()
mm.Text = ""
End Sub

Private Sub mm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
transac.Show
transac.Enabled = True
End If
End Sub

Private Sub yy_Change()
If Len(yy.Text) = 4 Then
Command1.SetFocus
End If
End Sub

Private Sub yy_GotFocus()
yy.Text = ""

End Sub

Private Sub yy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
transac.Show
transac.Enabled = True
End If
End Sub
