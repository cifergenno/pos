VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'On Error GoTo bi
Dim fld As ADODB.Field
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set RS3 = New ADODB.Recordset
Set conn = New ADODB.Connection
Dim numb As String
Dim margin As Double



conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=pos;" & "UID=root;PWD=; OPTION=3"
conn.Open

Dim numb2 As String
rs.Open "SELECT CP FROM stock_info", conn
rs2.Open "SELECT RP FROM stock_info", conn
RS3.Open "SELECT ITEM_CODE FROM stock_info", conn
Dim aa() As String
Do Until rs.EOF

For Each fld In rs2.Fields
numb = fld.Value
Next


For Each fld In rs.Fields
 numb2 = fld.Value
Next

For Each fld In RS3.Fields
margin = ((Val(numb) - Val(numb2)) / Val(numb2)) * 100
aa = Split(margin, ".")
margin = Left(margin, (Len(aa(0)) + 3))


conn.Execute "UPDATE stock_info SET MARGIN = " & "'" & margin & "'" & " WHERE ITEM_CODE = '" & fld.Value & "'"
Next


rs.MoveNext
rs2.MoveNext
RS3.MoveNext
Loop
bi:
MsgBox "done"
Exit Sub

End Sub
