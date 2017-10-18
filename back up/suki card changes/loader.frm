VERSION 5.00
Begin VB.Form loader 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   5430
   ClientTop       =   4605
   ClientWidth     =   6645
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "loader.frx":0000
   ScaleHeight     =   2100
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox f11 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   840
      Top             =   2880
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Text            =   "100 %"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Opening BISU_Itext . . . "
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   2640
   End
   Begin VB.Label ping 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   5925
   End
   Begin VB.Label text2 
      BackColor       =   &H00FFFF80&
      Height          =   550
      Left            =   300
      TabIndex        =   2
      Top             =   650
      Width           =   6000
   End
End
Attribute VB_Name = "loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Dim x1 As Integer
Private pindot As Boolean
Dim x_1 As Integer
Dim y_1 As Integer


Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim fld As ADODB.Field
Dim fld2 As ADODB.Field
Dim conn As ADODB.Connection

Public counterx As Integer


Private Sub Form_Load()
'If Not f11.Text = "" Then
'Unload Me
'End If

counter = 0
Timer2.Enabled = False
x1 = 0
f11.Text = ""
counter1 = 0
text2.Width = 10
text2.Caption = ""
Text4.Text = "0%"
pindot = False





End Sub




Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pindot = True
x_1 = X
y_1 = Y
pindot = True

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'close1.Visible = True
'close2.Visible = False
Dim x2 As Integer
Dim y2 As Integer

If pindot = True Then
x2 = x_1 - X
y2 = y_1 - Y
loader.Left = loader.Left - x2
loader.Top = loader.Top - y2
End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pindot = False
End Sub

Private Sub Timer1_Timer()

text2.Width = text2.Width + 10
Text1 = text2.Width
'motor.Left = motor.Left + 10

If text2.Width > 2000 Then
ping = "Loading User Profile..."
End If

If text2.Width < 2000 Then
ping = "Downloading POS Data Base..."
End If

If text2.Width > 4000 Then
ping = "Creating Grids, Modules and GUI..."
End If

If text2.Width >= 6015 Then
If counterx = 0 Then
MsgBox "RIC'z Cycle Parts... POS System is now open for business."
f11.Text = ""
Timer1.Enabled = False
'Timer2.Enabled = True
Unload Me
main.Show
End If
End If

End Sub

Private Sub Timer2_Timer()
Unload Me
BISU_Itext.Show
End Sub
