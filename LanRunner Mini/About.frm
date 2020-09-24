VERSION 5.00
Begin VB.Form About 
   Caption         =   "About"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   $"About.frx":0000
      Height          =   855
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Lan    Runner  Mini-V-1.0"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1500
      Index           =   0
      Left            =   120
      Picture         =   "About.frx":00C5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "About me,Mail  to : ravindran_ve@yahoo.com"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   $"About.frx":3AD9
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
End Sub
