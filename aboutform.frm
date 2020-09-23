VERSION 5.00
Begin VB.Form aboutform 
   Caption         =   "About SQL2CSV"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton okabout 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Warning: This computer program is not protected!  So take it and use it and if you improve on it e-mail me at mygoohoo@yahoo.com"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"aboutform.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "aboutform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub okabout_Click()
    Unload Me
End Sub
