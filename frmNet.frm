VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TCP Communications Window"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   1200
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtInfo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   3375
   End
   Begin RichTextLib.RichTextBox rtfChat 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmNet.frx":0000
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   4200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   4200
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfColor 
      Height          =   30
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   53
      _Version        =   393217
      TextRTF         =   $"frmNet.frx":0082
   End
   Begin VB.Label lblConnect 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   $"frmNet.frx":010D
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PTime As Double 'Time to pause for
Public STtime As Double
'Explanation of Signals
    'buddyN:... <-- Remote user Nick
    'redBox:## <-- Red Box Count (for resync)
    'bluBox:## <-- Blue Box Count (for resync)
    'mouseX: <-- Postion of Mouse...see if mouse is moving
    'mouseY: <-- Position of Mouse...see if mouse is moving
    'ipCHAT:... <--Chatting between players
    'vClick:Index <-- Vertical Line Activation
    'hClick:Index <-- Horizontal Line Activation
Public NickName As String


Private Sub cmdSend_Click()
If txtInfo.Text <> "" Then
    If wskClient.State = 0 Then
        'server mode
        AddText frmHost.txtNick & ": " & txtInfo.Text
        SendData 0, "ipCHAT:" & txtInfo.Text
    End If
    If wskServer.State = 0 Then
        'client mode
        AddText frmConnect.txtNick & ": " & txtInfo.Text
        SendData 1, "ipCHAT:" & txtInfo.Text
    End If
End If

txtInfo.Text = ""
txtInfo.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtInfo.SetFocus
AlwaysOnTop Me, True
End Sub

Private Sub Form_Load()
PTime = 0.7 'ping
End Sub

Private Sub Form_Paint()
On Error Resume Next
AlwaysOnTop Me, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Hide 'do not unload the form...Will destroy connection
Cancel = 1
End Sub

Public Sub AddText(note As String)
'Type into textbox
rtfChat.Text = rtfChat.Text + vbCrLf + note
rtfChat.SelStart = Len(rtfChat.Text)
End Sub

Private Sub Label2_DblClick()
'show client IP
If wskServer.State = 0 Then Exit Sub
If InputBox("Developer Control Code Required", "DevPass:", "") = "yourpass" Then
    Label1.Visible = True: txtRemoteIP.Visible = True
    txtRemoteIP.Text = wskServer.RemoteHostIP
End If
End Sub



Private Sub Timer1_Timer()
lblConnect.Caption = "You are connected at: " & PTime & " s"
'Ping time speed, used to slow down transmissions
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
'Explanation of Signals
    'buddyN:... <-- Remote user Nick
    'redBox:## <-- Red Box Count (for resync)
    'bluBox:## <-- Blue Box Count (for resync)
    'upTurn:0 or 1 <-- Whose turn is it? (for resync)
    'mouseX: <-- Postion of Mouse...see if mouse is moving
    'mouseY: <-- Position of Mouse...see if mouse is moving
    'ipCHAT:... <--Chatting between players
    'vClick:Index <-- Vertical Line Activation
    'hClick:Index <-- Horizontal Line Activation
    
Dim strData As String, strData2 As String

    wskClient.GetData strData
    'split data
        On Error Resume Next
        strData2 = Mid(strData, 8, Len(strData))
        strData = Left(strData, 7)
    'StrData2 has the info we need
    'StrData is the type of transmittion
Select Case strData
    Case "xxPing:" 'Bases of Ping Function
        wskClient.SendData "xPing2:"
        STtime = Timer
    Case "xPing2:"
        PTime = Ping(STtime)
    Case "buddyN:"
        NickName = strData2 'Set nickname
        SendData 1, "buddyN:" & frmConnect.txtNick.Text 'send back my name
        
    Case "redBox:"
        If frmMain.RedBoxCount <> Val(strData2) Then
            'out of sync
        End If
        
    Case "blueBox:"
        If frmMain.BlueBoxCount <> Val(strData2) Then
            'out of sync
        End If
        
    Case "upTurn:"
        frmMain.turn = Val(strData2)
        
    Case "mouseX:"
        frmMain.NetMouseX = Val(strData2)
        
    Case "mouseY:"
        frmMain.NetMouseY = Val(strData2)
        
    Case "ipCHAT:"
        AddText NickName & ": " & strData2
        
    Case "srTEXT:"
        AddText strData2
        
    Case "vClick:"
        frmMain.makeVLine Val(strData2)
        
    Case "hClick:"
        frmMain.makeHLine Val(strData2)
        
    Case "discon:"
        On Error Resume Next
        MsgBox "Remote Computer Disconnected", vbCritical, "Closing your connection"
        wskClient.Close
        frmMain.mnuFileNew_Click
        Exit Sub
    
    Case "stGame:"
        frmMain.Command1.Caption = "Start!"
        frmMain.Command1_Click
        With frmMain
            .Text1.Visible = False
            .Text2.Visible = False
            .Label1.Visible = False
            .Label2.Visible = False
            .Command1.Visible = False
            .cmdOnline.Visible = True
        End With
        SendData 1, "gREADY:"
        frmMain.mnuMultiConnect.Enabled = False
        frmMain.mnuMultiHost.Enabled = False
        
    Case "Height:"
        frmMain.Text1.Text = Val(strData2)
        frmMain.Text1.Visible = False
        
    Case "lWidth:"
        frmMain.Text2.Text = Val(strData2)
        frmMain.Text2.Visible = False
    'resyncs
    Case "resyn1:"
        frmMain.turn = Val(strData2)
    Case "resyn2:"
        frmMain.Text1.Text = Val(strData2)
    Case "resyn3:"
        frmMain.Text2.Text = strData2
    Case "resyn4:"
        frmMain.RedBoxCount = 0
        frmMain.BlueBoxCount = 0
        frmMain.UnloadBoard
        frmMain.tturn.Enabled = False
        frmMain.CreateBoard
        frmMain.tturn.Enabled = True
        frmMain.ReSyncGame strData2
        frmMain.Show
        AlwaysOnTop frmReSync, False
        frmReSync.Hide
        SendData 1, "donSYN:"
    Case "resyn5:"
        frmMain.RedBoxCount = 0
        frmMain.BlueBoxCount = 0
    Case "resyn6:"
        frmMain.RedBoxCount = Val(strData2)
    Case "resyn7:"
        frmMain.BlueBoxCount = Val(strData2)
    Case "resyn8:"
        frmMain.turn = Val(strData2)
    Case "getRED:"
        If Val(strData2) <> frmMain.RedBoxCount Then frmMain.mnuOnlineSynch_Click
    Case "getBLU:"
        If Val(strData2) <> frmMain.BlueBoxCount Then frmMain.mnuOnlineSynch_Click
    Case "getTRN:"
        If Val(strData2) <> frmMain.turn Then frmMain.mnuOnlineSynch_Click
    Case Else
        Debug.Print strData & "   " & strData2
End Select
'AddText "debug___" & strData & strData2
End Sub

Private Sub wskServer_ConnectionRequest(ByVal requestID As Long)
    If wskServer.State <> sckClosed Then wskServer.Close
    wskServer.Accept requestID
    STtime = Timer
    SendData 0, "xxPing:"
    SendData 0, "buddyN:" & frmHost.txtNick.Text 'Send Nickname
    SendData 0, "Height:" & Val(frmHost.Text1.Text)
    SendData 0, "lWidth:" & Val(frmHost.Text2.Text)
    If frmHost.chkStart.Value = 0 Then
        SendData 0, "srTEXT:You go first"
        frmMain.turn = 1
    End If
    If frmHost.chkStart.Value = 1 Then
        SendData 0, "srTEXT:I go first"
        frmMain.turn = 0
    End If
    SendData 0, "srTEXT:Level Height: " & frmHost.Text1.Text
    SendData 0, "srTEXT:Level Width: " & frmHost.Text2.Text
    SendData 0, "stGame:" 'Now that data is all sent, start game
End Sub

Private Sub wskServer_DataArrival(ByVal bytesTotal As Long)

'Explanation of Signals
    'buddyN:... <-- Remote user Nick
    'redBox:## <-- Red Box Count (for resync)
    'bluBox:## <-- Blue Box Count (for resync)
    'upTurn:0 or 1 <-- Whose turn is it? (for resync)
    'mouseX: <-- Postion of Mouse...see if mouse is moving
    'mouseY: <-- Position of Mouse...see if mouse is moving
    'ipCHAT:... <--Chatting between players
    'vClick:Index <-- Vertical Line Activation
    'hClick:Index <-- Horizontal Line Activation
    
Dim strData As String, strData2 As String
    wskServer.GetData strData
    'split data
        strData2 = Mid(strData, 8, Len(strData))
        strData = Left(strData, 7)
    'StrData2 has the info we need
    'StrData is the type of transmittion
Select Case strData

    Case "xPing2:"
        PTime = Ping(STtime)
        wskServer.SendData "xPing2:"
        
    Case "buddyN:"
        NickName = strData2 'Set nickname
        
    Case "redBox:"
        If frmMain.RedBoxCount <> Val(strData2) Then
            'out of sync
        End If
        
    Case "blueBox:"
        If frmMain.BlueBoxCount <> Val(strData2) Then
            'out of sync
        End If
        
    Case "upTurn:"
        If frmMain.turn <> Val(strData2) Then
            'out of sync
        End If
        
    Case "mouseX:"
        frmMain.NetMouseX = Val(strData2)
        
    Case "mouseY:"
        frmMain.NetMouseY = Val(strData2)
        
    Case "ipCHAT:"
        AddText NickName & ": " & strData2
        
    Case "vClick:"
        frmMain.makeVLine Val(strData2)
        
    Case "hClick:"
        frmMain.makeHLine Val(strData2)
        
    Case "discon:"
        MsgBox "Remote Computer Disconnected", vbCritical, "Closing your connection"
        If wskServer.State <> 0 Then wskServer.Close
        frmMain.mnuFileNew_Click
        Exit Sub
        
    Case "gREADY:" 'start game
        On Error Resume Next
        frmHost.Visible = False
        frmMain.Command1.Caption = "Start!"
        frmMain.Command1_Click
        If frmHost.chkStart.Value = 0 Then
            SendData 0, "upTurn:1"
            frmMain.turn = 1
        End If
        If frmHost.chkStart.Value = 1 Then
            SendData 0, "upTurn:0"
            frmMain.turn = 0
        End If
        If frmOnline.Text1.Text <> "" Then
            If frmOnline.chkStart.Value = 0 Then
                SendData 0, "upTurn:1"
                frmMain.turn = 1
            End If
            If frmOnline.chkStart.Value = 1 Then
                SendData 0, "upTurn:0"
                frmMain.turn = 0
            End If
        End If
        frmMain.mnuMultiConnect.Enabled = False
        frmMain.mnuMultiHost.Enabled = False
        frmOnline.Text1.Text = ""
        frmOnline.Text2.Text = ""
        AddText "Game Started"
        'the other player's level is ready
    Case "donSYN:" 'Done syncing
        frmNet.SendData 0, "resyn5:" & frmMain.Box.Count
        frmNet.SendData 0, "resyn6:" & frmMain.RedBoxCount
        frmNet.SendData 0, "resyn7:" & frmMain.BlueBoxCount
        frmNet.SendData 0, "resyn8:" & frmMain.turn
        frmMain.Show
        AlwaysOnTop frmReSync, False
        frmReSync.Hide
    Case "askSYN:"
        frmMain.mnuOnlineSynch_Click
    Case "getRED:"
        If Val(strData2) <> frmMain.RedBoxCount Then frmMain.mnuOnlineSynch_Click
    Case "getBLU:"
        If Val(strData2) <> frmMain.BlueBoxCount Then frmMain.mnuOnlineSynch_Click
    Case "getTRN:"
        If Val(strData2) <> frmMain.turn Then frmMain.mnuOnlineSynch_Click
    Case Else
        Debug.Print strData & "   " & strData2
End Select
'AddText "Debug___" & strData
End Sub

Public Function SendData(ClientServer As Integer, Data As Variant) As Boolean
On Error Resume Next
    
    DoEvents
    
    Dim TimeOut As Long, Transferred As Boolean
    
    
3    Transferred = False
    Select Case ClientServer
        Case 0 'server
            wskServer.SendData Data
            SendData = True
        Case 1 'client
            wskClient.SendData Data
            SendData = True
    End Select
    If Data = "discon:" Then Pause 1 Else Pause PTime + 0.2
    
    Exit Function
    
123456
    SendData = False
    MsgBox Err.Description, vbSystemModal, "Error # " & Err.Number
End Function
Public Sub Pause(PauseTime As Double)
Dim Start, Finish, TotalTime
    
'Pause Function from Help put into this
   Start = Timer   ' Set start time.
   Do While Timer < Start + PauseTime
      DoEvents   ' Yield to other processes.
   Loop
   Finish = Timer   ' Set end time.
   TotalTime = Finish - Start   ' Calculate total time.

End Sub
Public Function Ping(ByVal Start As Double) As Double
Ping = Timer - Start 'Wonderful Pinging
End Function
