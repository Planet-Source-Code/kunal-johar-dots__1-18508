VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect to a Host"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1080
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtMyIp 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Your Nick:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Host's IP Address:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Your IP Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
frmNet.PTime = 0.7 'Slowest Ping Time allowed
    Dim isConnected As Boolean 'Are we there yet?
    Dim ipClient As String 'IP Address
    Dim tmrConnect As Integer ' Pause for connects
    Timer1.Enabled = True
    If frmNet.wskClient.State <> 0 Then frmNet.wskClient.Close
            'If its not closed, close it
    ipClient = Me.txtHostIP.Text 'Set the string to the box
    tmrConnect = 0 'Not a clue why I did this
22
    frmNet.wskClient.Connect ipClient, 747 'connect to port 747

    'Was a method of error checking but...screw
    'While (frmNet.wskClient.State <> sckConnected) And (tmrConnect < 300000)
        'DoEvents
        'tmrConnect = tmrConnect + 1
    'Wend
    'If tmrConnect >= 300000 Then
        'MsgBox "Remote Computer not Responding", vbOKOnly, "Error"
        'frmNet.wskClient.Close
        'Exit Sub
    'End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
'Get the names
frmConnect.txtNick.Text = GetSetting(App.Title, "Online", "Client", "Client")
frmHost.txtNick.Text = GetSetting(App.Title, "Online", "Server", "Server")
Load frmNet
frmNet.Hide
txtMyIp.Text = frmNet.wskClient.LocalIP

End Sub

Private Sub Timer1_Timer()
'New method of reconnecteing
Static connection As Double
If frmNet.wskClient.State = 7 Then
    Me.Hide: Timer1.Enabled = False
    frmNet.Visible = True
    frmNet.SetFocus
End If
connection = connection + 1
If connection >= 100 Then
    connection = 0
    MsgBox "Connection Failed", vbOKOnly, "Try Again"
End If
Timer1.Enabled = False
End Sub

Private Sub txtNick_change()
'Save the nickname into the Registry
SaveSetting App.Title, "Online", "Client", txtNick.Text
End Sub
