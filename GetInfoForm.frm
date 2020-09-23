VERSION 5.00
Begin VB.Form GetInfoForm 
   Caption         =   "SQL2CSV"
   ClientHeight    =   4305
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox TextFileF 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox UserPassF 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox UserNameF 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox TableF 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox DatabaseF 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox ServerNameF 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Text File"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "User  Password"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Table"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sql Database"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Sql Server Name"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Menu FileM 
      Caption         =   "File"
      Begin VB.Menu ExitM 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu HelpM 
      Caption         =   "Help"
      Begin VB.Menu NohelpM 
         Caption         =   "SorryNoHelp"
      End
      Begin VB.Menu AboutM 
         Caption         =   "About SQL2CSV"
      End
   End
End
Attribute VB_Name = "GetInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AboutM_Click()
   aboutform.Show
End Sub

Private Sub Command1_Click()

Dim cn As ADODB.Connection
Dim outtb As New ADODB.Recordset
Dim CONSTR As String


Set cn = New ADODB.Connection
   CONSTR = "driver={SQL Server};" & _
   "server=" & ServerNameF.Text & ";" & _
   "uid=" & UserNameF.Text & ";" & _
   "pwd=" & UserPassF.Text & ";" & _
   "database=" & DatabaseF.Text
   
   cn.ConnectionString = CONSTR
   cn.ConnectionTimeout = 30
   cn.Open
   outtb.Open TableF.Text, cn, adOpenStatic, adLockOptimistic

   EXPORTQUOCOM TextFileF.Text, outtb
   outtb.Close
   cn.Close
   
End Sub


Private Sub ExitM_Click()
  Unload Me
End Sub
