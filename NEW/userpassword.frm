VERSION 5.00
Begin VB.Form userpassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ric's Cyle Part and Accessories Center"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "&Accept"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "User"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Return Item/s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "userpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
main.Enabled = True
Unload returnme
main.Show
End Sub
