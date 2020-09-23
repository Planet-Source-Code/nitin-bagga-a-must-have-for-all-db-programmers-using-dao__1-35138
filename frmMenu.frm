VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Specify Database Path"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Report using MS Word"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect To Database"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If InitApplication(Me) Then
        MsgBox "Database Connected Successfully"
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
    Else
        SysForm.Show
        'MsgBox "Databrease Connection failed"
    End If
End Sub

Private Sub Command2_Click()

Dim myWord As Object
Set myWord = GetObject("", "Word.Basic")
myWord.AppMaximize ("Microsoft Word")
myWord.FileNew
myWord.FormatFont Points:="12", Font:="Arial", Bold:=1
    'myWord.formatparagraph Alignment:=1
myWord.Insert "List of Areas "
myWord.insertpara
myWord.insertpara
myWord.tableinserttable NumColumns:=2, NumRows:=1 ', ColumnWidth:=0.5

Dim rsTemp As Recordset
Dim str As String

str = "select * from list"
Set rsTemp = gdb.OpenRecordset(str)
    If rsTemp.EOF Then
        myWord.Insert "No Records found"
    Else
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            myWord.Insert CStr(rsTemp(0))
            myWord.nextcell
            myWord.Insert CStr(rsTemp(1))
            rsTemp.MoveNext
            myWord.nextcell
        Loop
    End If

End Sub

Private Sub Command3_Click()
SysForm.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End Sub
