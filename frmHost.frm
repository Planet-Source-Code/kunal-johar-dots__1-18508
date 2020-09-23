VERSION 5.00
Begin VB.Form frmHost 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Host a Game"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1110
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   1560
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "I go First"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Width:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1110
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Height:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Your Nick:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Your IP Address:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdListen_Click()
frmNet.PTime = 0.7 ' Ping time allowed
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
    'Now that we have a proper height and width...
If cmdListen.Caption = "Listen" Then 'Open the port to connections
    With frmNet.wskServer
        .LocalPort = 747
        .Listen
    End With
    frmMain.Text1.Visible = False
    frmMain.Text2.Visible = False
    frmMain.Label1.Visible = False
    frmMain.Label2.Visible = False
    frmMain.Command1.Visible = False
    frmMain.cmdOnline.Visible = True
    cmdListen.Caption = "Listening"
    frmMain.Text1.Text = frmHost.Text1.Text
    frmMain.Text2.Text = frmHost.Text2.Text
    Exit Sub
End If
If cmdListen.Caption = "Listening" Then
    With frmNet.wskServer
        If .State <> sckClosed Then .Close 'If it was listening for a connection, stop
    End With
    frmMain.Text1.Visible = True
    frmMain.Text2.Visible = True
    frmMain.Label1.Visible = True
    frmMain.Label2.Visible = True
    frmMain.cmdOnline.Visible = False
    frmMain.Command1.Visible = True
    cmdListen.Caption = "Listen"
End If
    
End Sub

Private Sub Form_Activate()
'Get name from registry
On Error Resume Next
frmConnect.txtNick.Text = GetSetting(App.Title, "Online", "Client", "Client")
frmHost.txtNick.Text = GetSetting(App.Title, "Online", "Server", "Server")
Load frmNet
txtHostIP.Text = frmNet.wskServer.LocalIP
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Stop serving incase exit
If cmdListen.Caption = "Listening" Then cmdListen_Click
End Sub

Private Sub Timer1_Timer()
'Error Checking...or when connected play game
If frmNet.wskServer.State = 7 Then
    Me.Hide: Timer1.Enabled = False
    frmNet.Visible = True
    frmNet.SetFocus
End If
End Sub

Private Sub txtNick_change()
'Save Server Name
SaveSetting App.Title, "Online", "Server", txtNick.Text
End Sub
