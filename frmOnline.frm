VERSION 5.00
Begin VB.Form frmOnline 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   1275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "I go First"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   390
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Height:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Width:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   390
      Width           =   495
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'start a new online game with this shortened version
    'of the Game Starter
    frmMain.Text1.Text = Text1.Text
    frmMain.Text2.Text = Text2.Text
    If Val(Text1.Text) > 14 Or Val(Text1.Text) < 2 Then
        MsgBox "The Height is not between 2 and 14!", vbInformation, "Error in Height"
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    If Val(Text2.Text) > 22 Or Val(Text2.Text) < 2 Then
        MsgBox "The width is not between 2 and 22!", vbInformation, "Error in Height"
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
        Exit Sub
    End If
    frmNet.SendData 0, "Height:" & Val(Text1.Text)
    frmNet.SendData 0, "lWidth:" & Val(Text2.Text)
    If chkStart.Value = 0 Then
        frmNet.SendData 0, "srTEXT:You go first"
        frmMain.turn = 1
    End If
    If chkStart.Value = 1 Then
        frmNet.SendData 0, "srTEXT:I go first"
        frmMain.turn = 0
    End If
    frmNet.SendData 0, "srTEXT:Level Height: " & Text1.Text
    frmNet.SendData 0, "srTEXT:Level Width: " & Text2.Text
    frmNet.SendData 0, "stGame:"
    Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()

End Sub
