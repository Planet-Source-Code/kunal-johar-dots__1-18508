VERSION 5.00
Begin VB.Form frmReSync 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmReSync.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Game Resync in Progress, please wait"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmReSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form blinks while game is resyncing
Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
DoEvents
    Image1.Visible = Not Image1.Visible
End Sub
