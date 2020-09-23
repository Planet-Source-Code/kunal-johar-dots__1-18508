VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help Window"
   ClientHeight    =   4695
   ClientLeft      =   435
   ClientTop       =   1455
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtfHelp 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8281
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":0000
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Used to show help files
Private Sub Form_Activate()
rtfHelp.SelStart = 0
frmHelp.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub rtfHelp_Change()

End Sub
