VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SysForm 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Setup"
   ClientHeight    =   2625
   ClientLeft      =   2475
   ClientTop       =   1785
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CommandHelp 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Help"
      Height          =   425
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Properties"
      ForeColor       =   &H8000000E&
      Height          =   1845
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   5055
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Browse"
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox TextData 
         BackColor       =   &H00C0C000&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   750
         Width           =   3435
      End
      Begin VB.TextBox TextData 
         BackColor       =   &H00C0C000&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   0
         Text            =   "Access"
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label labelSect 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label labelSect 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base Path:"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   750
         Width           =   1305
      End
      Begin VB.Label labelSect 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base Type:"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.CommandButton ButtonCancel 
      BackColor       =   &H00FFFF00&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   425
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton ButtonOK 
      BackColor       =   &H00FFFF00&
      Caption         =   "&OK"
      Height          =   425
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "SysForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mDBType = 0
Private Const mDBPath = 1
Private mResult As Integer

Private Sub ButtonCancel_Click()
    DLGResult = vbCancel
End Sub
Private Sub ButtonOK_Click()
    SaveSetting gcAppName, "DataBase", "DBType", TextData(mDBType).Text
    SaveSetting gcAppName, "DataBase", "DBPath", TextData(mDBPath).Text
    DLGResult = vbOK
    
End Sub

Private Property Get DLGResult() As Integer
    DLGResult = mResult
End Property

Private Property Let DLGResult(iNewValue As Integer)
    mResult = iNewValue
    Unload Me
    'frmMenu.Show
End Property

Public Function Execute() As Boolean
    'Me.Show vbModal
    If DLGResult = vbOK Then
        Execute = True
    Else
        Execute = False
    End If
End Function

Private Sub Command1_Click()
    CommonDialog1.Filter = "Access Databases (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    TextData(mDBPath).Text = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
    'CenterForm Me
    TextData(mDBType).Text = GetSetting(gcAppName, "DataBase", "DBType", "Access")
    TextData(mDBPath).Text = GetSetting(gcAppName, "DataBase", "DBPath", "")
End Sub

Private Sub textData_GotFocus(Index As Integer)
    With TextData(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


