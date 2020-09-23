VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dots"
   ClientHeight    =   6135
   ClientLeft      =   -19695
   ClientTop       =   -165
   ClientWidth     =   8745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0442
   Moveable        =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrOn 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5445
      Top             =   990
   End
   Begin VB.Timer tmrURL 
      Enabled         =   0   'False
      Interval        =   12
      Left            =   8310
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   255
      Left            =   7440
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Multiplayer"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File"
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOnline 
      Caption         =   "Online Game"
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tmrSync 
      Interval        =   60000
      Left            =   4200
      Top             =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":074C
            Key             =   ""
            Object.Tag             =   "Connec&t"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A66
            Key             =   ""
            Object.Tag             =   "&Frequently Asked Questions"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EB8
            Key             =   ""
            Object.Tag             =   "&How to Play"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130A
            Key             =   ""
            Object.Tag             =   "&Rules of the Game"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":175C
            Key             =   ""
            Object.Tag             =   "New Game"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BAE
            Key             =   ""
            Object.Tag             =   "Synchronize"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2000
            Key             =   ""
            Object.Tag             =   "Disconnect"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2452
            Key             =   ""
            Object.Tag             =   "&Game Options"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28A4
            Key             =   ""
            Object.Tag             =   "&Host a Game"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BBE
            Key             =   ""
            Object.Tag             =   "&Show Multiplayer Chat Window"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3010
            Key             =   ""
            Object.Tag             =   "&Go to Chat Server"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3462
            Key             =   ""
            Object.Tag             =   "&General Help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38B4
            Key             =   ""
            Object.Tag             =   "&About Dots"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D06
            Key             =   ""
            Object.Tag             =   "E&xit"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4158
            Key             =   ""
            Object.Tag             =   "&New Local Game"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tturn 
      Interval        =   50
      Left            =   -75
      Top             =   3795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start!"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      MaxLength       =   2
      TabIndex        =   3
      Top             =   390
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   5295
      Left            =   240
      MouseIcon       =   "frmMain.frx":4472
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":477C
      ScaleHeight     =   5295
      ScaleWidth      =   8175
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   8175
      Begin VB.Shape Box 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   360
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape Dot 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   240
         Shape           =   3  'Circle
         Top             =   240
         Width           =   135
      End
      Begin VB.Label hHover 
         BackStyle       =   0  'Transparent
         Height          =   135
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label vHover 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   135
      End
      Begin VB.Line hLine 
         Index           =   0
         Visible         =   0   'False
         X1              =   360
         X2              =   600
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line vLine 
         Index           =   0
         Visible         =   0   'False
         X1              =   298
         X2              =   298
         Y1              =   600
         Y2              =   360
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Your Resolution is below 800x600 therefore Dots is looking ugly"
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label4 
      Height          =   2775
      Left            =   8400
      TabIndex        =   13
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image noMouse 
      Height          =   480
      Left            =   2880
      Picture         =   "frmMain.frx":4CFC
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image red 
      Height          =   480
      Left            =   2280
      Picture         =   "frmMain.frx":513E
      Top             =   240
      Width           =   480
   End
   Begin VB.Image blue 
      Height          =   480
      Left            =   2160
      Picture         =   "frmMain.frx":5448
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Boxes Left:"
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblRedBox 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblBlueBox 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Height:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Width:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   390
      Width           =   495
   End
   Begin VB.Label lblMessages 
      Caption         =   "Set the height and width of the gameboard and press Start!"
      Height          =   3975
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Local Game"
      End
      Begin VB.Menu mnuGameOptions 
         Caption         =   "&Game Options"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMulti 
      Caption         =   "&Multiplayer"
      Begin VB.Menu mnuMultiConnect 
         Caption         =   "Connec&t"
      End
      Begin VB.Menu mnuMultiHost 
         Caption         =   "&Host a Game"
      End
      Begin VB.Menu mnuMore 
         Caption         =   "-More Options"
      End
      Begin VB.Menu mnuMultiChatServ 
         Caption         =   "&Go to Chat Server"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "-Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMultiChat 
         Caption         =   "&Show Multiplayer Chat Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&General Help"
         Begin VB.Menu mnuHelpPlay 
            Caption         =   "&How to Play"
         End
         Begin VB.Menu mnuHelpRules 
            Caption         =   "&Rules of the Game"
         End
         Begin VB.Menu mnuHelpFAQ 
            Caption         =   "&Frequently Asked Questions"
         End
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Dots"
      End
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "Online Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOnlineNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuOnlineOther 
         Caption         =   "-Other"
      End
      Begin VB.Menu mnuOnlineSynch 
         Caption         =   "Synchronize Game"
      End
      Begin VB.Menu mnuOnlineDisconnect 
         Caption         =   "Disconnect"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dots 1.5 is by Kunal Johar
'Visit me on the web at http://www.bootlegzrus.com
'We are not just software!


'Anyway this game got its fare share of downloads around
'the internet including Download.com where it still is
'Anyway after taking ages of time to get Radiate to
'get this app approved (thus delaying the release)
'they kick me out a few weeks later.  So in return
'I see no point to holding on to the sourcecode of a
'freeware app.  If you find this useful all the better!

'And please don't laugh at some of the code
'you may want to clear the glitches and let me know
'we can do a joint rerelease
'Kunal@BootlegZRus.com

'Have fun with this code/game
'It works to a great extent now that the adbar is gone
'Everynow and then a resyncing error occurs, oh well

'The best thing in this program is the ping function
'Search PSCOde.com for Ping Function and i already
'took it out of here into an easier to understand
'compact version





'Cool Menu, you can search PS CODE for Cool Menu
'Placed Icons in Menus, to bad it required MSCOMCTLs
Private WithEvents HelpObj As HelpCallBack
Attribute HelpObj.VB_VarHelpID = -1
Const Space = 360, _
      MidL = 300, _
      StartTopLeft = 240, _
      SIZE = 480
Dim iHeight As Integer, _
    iWidth As Integer 'Height and width of the board
Public NetMouseX As Integer 'Was going to show the mouse of the
Public NetMouseY As Integer 'Other player...until I trashed that Idea
Public CliSend As Boolean 'Helps with internet sendng
Dim SyncCount As Integer 'Are we in sync
'Player Stuff
    Public turn As Integer ' 0 = red, 1 = blue
    'Red Player
        Public RedBoxCount As Single 'How many boxes
    'Blue Player
        Public BlueBoxCount As Single

        
Public Sub CreateBoard()
UnloadBoard 'Stop errors

'Structure of Board
    '2 Shapes - and |
    '2 transparent labels over them to get events
    'Then a dot between every 2 lines and a box bewteen every 4 dots
    Dim i As Integer, j As Integer 'For Loop Counters
    Dim CurNum As Integer
        'Do Initial Settings for Dots and Lines
    
    
    'Sets the defaults of all objects on the board
    Dot(0).Visible = True
    hLine(0).Visible = False
    hLine(0).Tag = ""
    vLine(0).Tag = ""
    vLine(0).Visible = False
    Box(0).Visible = False
    
'For every row/column draw the needed stuff and set it to defaults
For j = 0 To iHeight - 1
    If j = 0 Then
        For i = 1 To iWidth - 1
            Load Dot(i)
            With Dot(i)
                .Container = Picture1
                .Top = StartTopLeft + (j * Space)
                .Left = (StartTopLeft) + (i * Space)
                .Visible = True
            End With
        Next i
     End If
    If j <> 0 Then
        For i = 0 To iWidth - 1
            CurNum = i + ((j) * iWidth)
            Load Dot(CurNum)
            With Dot(CurNum)
                .Container = Picture1
                .Top = StartTopLeft + (j * Space)
                .Left = (StartTopLeft) + (i * Space)
                .Visible = True
            End With
        Next i
    End If
    If j = 0 Then
        For i = 1 To iWidth - 2
            Load hLine(i)
            With hLine(i)
                .Container = Picture1
                .X1 = StartTopLeft + (Space * (i))
                .X2 = .X1 + (SIZE)
                .Y1 = MidL
                .Y2 = MidL
                .Visible = False
            End With
            Load hHover(i)
            With hHover(i)
                .Container = Picture1
                .Visible = True
                .Left = (StartTopLeft) + Dot(0).Width + (i * Space)
                .Top = (StartTopLeft) + (j * Space)
            End With
        Next i
    End If
    If j <> 0 Then
        For i = 0 To iWidth - 2
            CurNum = hLine.Count
            Load hLine(CurNum)
            With hLine(CurNum)
                .Container = Picture1
                .X1 = StartTopLeft + (Space * (i))
                .X2 = .X1 + (SIZE)
                .Y1 = StartTopLeft + (j * Space) + (0.5 * Dot(0).Height)
                .Y2 = .Y1
                .Visible = False
            End With
            Load hHover(CurNum)
            With hHover(CurNum)
                .Container = Picture1
                .Visible = True
                .Left = (StartTopLeft) + Dot(0).Width + (i * Space)
                .Top = (StartTopLeft) + (j * Space)
            End With
        Next
    End If
    If j = 0 Then
        For i = 1 To iWidth - 2
            Load Box(i)
            With Box(i)
                .Container = Picture1
                .Top = (j + 1) * Space
                .Left = (i + 1) * Space
                .Visible = False
            End With
        Next i
    End If
    If j <> 0 And j <> iHeight - 1 Then
        For i = 0 To iWidth - 2
            CurNum = Box.Count
            Load Box(CurNum)
            With Box(CurNum)
                .Container = Picture1
                .Top = (j + 1) * Space
                .Left = (i + 1) * Space
                .Visible = False
            End With
        Next
    End If
    If j = 0 Then
        For i = 1 To iWidth - 1
            Load vLine(i)
            With vLine(i)
                .Container = Picture1
                .Y1 = StartTopLeft + (Space * (j))
                .Y2 = .Y1 + (SIZE)
                .X1 = MidL + (Space * i)
                .X2 = MidL + (Space * i)
                .Visible = False
            End With
            Load vHover(i)
            With vHover(i)
                .Container = Picture1
                .Visible = True
                .Left = StartTopLeft + (i * Space)
                .Top = StartTopLeft + Dot(0).Height + (j * Space)
            End With
        Next
    End If
    If j <> 0 And j <> iHeight - 1 Then
        For i = 0 To iWidth - 1
            CurNum = vLine.Count
            Load vLine(CurNum)
            With vLine(CurNum)
                .Container = Picture1
                .Y1 = StartTopLeft + (Space * (j))
                .Y2 = .Y1 + (SIZE)
                .X1 = MidL + (Space * i)
                .X2 = MidL + (Space * i)
                .Visible = False
            End With
            Load vHover(CurNum)
            With vHover(CurNum)
                .Container = Picture1
                .Visible = True
                .Left = StartTopLeft + (i * Space)
                .Top = StartTopLeft + Dot(0).Height + (j * Space)
            End With
        Next
    End If
Next j
'I hope you guys don't mind that I didnt comment that
'It was a pain to finally get it right and I dont even want to look at it again
End Sub
Public Sub UnloadBoard()
'Unloads Everything cause the game is over
Dim i As Integer
    For i = 1 To Dot.Count - 1
        Unload Dot(i)
    Next
    For i = 1 To hLine.Count - 1
       Unload hLine(i)
    Next
    For i = 1 To vLine.Count - 1
        Unload vLine(i)
    Next
    For i = 1 To hHover.Count - 1
        Unload hHover(i)
    Next
    For i = 1 To vHover.Count - 1
        Unload vHover(i)
    Next
    For i = 1 To Box.Count - 1
        Unload Box(i)
    Next
End Sub

Private Sub cmdOnline_Click()
'Used to Control Internet Games
If frmNet.wskServer.State <> 0 Then mnuOnlineNew.Visible = True Else mnuOnlineNew.Visible = False
PopupMenu mnuOnline
End Sub

Public Sub Command1_Click() 'Too lazy to change from Command1
If Command1.Caption = "Start!" Then
'New Game
RedBoxCount = 0
BlueBoxCount = 0
turn = 0
tturn.Enabled = True
    'Error checking
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
    iHeight = Val(Text1.Text) 'Set height/width
    iWidth = Val(Text2.Text)
    
    lblMessages.Caption = "Loading Level"
Picture1.Visible = False
DoEvents
CreateBoard
DoEvents
Picture1.Visible = True
Command1.Caption = "Restart!"
Exit Sub
End If
If Command1.Caption = "Restart!" Then
    Picture1.Visible = False
    Command1.Caption = "Start!"
    lblMessages.Caption = "Start a game!"
End If

'Call CreateBoard(Val(Text1.Text), Val(Text2.Text))
End Sub

'Next three subs are for crappy computers with <800x600 resolution
Private Sub Command2_Click()
PopupMenu mnuFile
End Sub

Private Sub Command3_Click()
PopupMenu mnuMulti
End Sub

Private Sub Command4_Click()
PopupMenu mnuHelp
End Sub

Private Sub Form_Activate()
'Decide to use Ugly mode or good mode
If Screen.Width < 12000 Then
    mnuMulti.Visible = False
    mnuFile.Visible = False
    mnuHelp.Visible = False
    Me.Top = -300
    Me.Left = 0
    Label5.Visible = True
    Command2.Visible = True
    Command3.Visible = True
    Command4.Visible = True
End If
Me.lblMessages.Caption = "Start a New Game"
With lblMessages
    .Font.SIZE = 24
End With
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
'Type in chat box
If frmNet.Visible Then
    frmNet.SetFocus
    frmNet.txtInfo.SetFocus
End If

End Sub

Private Sub Form_Load()
On Error GoTo 532
SyncCount = 0 'Synchronized
CliSend = False
GetReg 'Get registry settings

'MsgBox "Dots Private Beta, No Internet, only 2 player, 1 computer available, send bugs to me at xxdakmanxx@yahoo.com"


turn = 0 'Red turn
  Set HelpObj = New HelpCallBack 'Used for Cool menu

  Call mCoolMenu.Install(Me.hWnd, HelpObj, ImageList1, True, True)
  
'Any property function must be used AFTER
'installation

'If the FontName property is nothing,
'CoolMenu uses the form's text style and size
'If you set FontName to something, default size
'and color will be used.
'Setting size without FontName as no effect
'  Call mCoolMenu.FontName(Me.hWnd, "Tahoma")
'  Call mCoolMenu.FontSize(Me.hWnd, 8)
'  Call mCoolMenu.ForeColor(Me.hWnd, &H80)

'This is yet to be resolved: bright colors on
'selection bar should print text in dark color
'  Call mCoolMenu.SelectColor(Me.hWnd, vbWhite)

Exit Sub
532
MsgBox Err.Description, vbSystemModal, "Error #" & Err.Number
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'get out of program
Cancel = 1 'Dont get out like this!, click the exit button
mnuFileExit_Click
End Sub

Private Sub Form_Terminate()
mnuFileExit_Click
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do
DoEvents
Call Module1.over 'Help exit
Loop
End Sub

Private Sub hHover_Click(Index As Integer)
'Click a place
Dim Service As Integer
Service = -1
If frmNet.wskClient.State <> 0 Then Service = 1 'client
If frmNet.wskServer.State <> 0 Then Service = 0 'server
If Service <> -1 Then 'connection
    If turn <> Service Then Exit Sub
    frmNet.SendData Service, "hClick:" & Index 'send data either via client or server
End If
makeHLine Index 'now i can draw the line on my own machine
End Sub

Private Sub Label4_DblClick()

If frmNet.wskServer.State <> 0 Then
    Dim X As String
    
    If GetSetting(App.Title, "Secret", "Pass", "") = "yourpass" Then GoTo 2
    If InputBox("Secret Area Code", "Game Developer's Secret Room") = "yourpass" Then
        'srTEXT:
   
        
        X = InputBox("What Server Text do you want to send?", "srTEXT: Send Operation"): GoTo 3
2       X = InputBox("What Server Text do you want to send?", "srTEXT: Send Operation")
3       If X <> "" Then frmNet.SendData 0, "srTEXT:" & X
        frmNet.AddText X
    End If
End If
'Function sends server text that goes into the chatroom RAW
'So you can make it seem like someone is typing to themselves
End Sub

Private Sub mnuFileExit_Click()
'Disconnect and exit
On Error Resume Next
If frmNet.wskClient.State <> 0 Then
    If MsgBox("You are in an online game, disconnect?", vbYesNo, "Quit and Disconnect?") = vbNo Then Exit Sub
    frmNet.SendData 1, "discon:"
    frmNet.wskClient.Close
End If
If frmNet.wskServer.State <> 0 Then
    If MsgBox("You are in an online game, disconnect?", vbYesNo, "Quit and Disconnect?") = vbNo Then Exit Sub
    frmNet.SendData 0, "discon:"
    frmNet.wskServer.Close
End If
mCoolMenu.Uninstall Me.hWnd
Unload frmConnect
Unload frmHost
Unload frmNet
Unload frmOptions
Unload frmReSync
Unload frmAbout
Unload frmHelp
End
End Sub

Public Sub mnuFileNew_Click()
'Start new local game
'If online exit
'Helps when they screw around with too many things and need to get back
On Error Resume Next
If frmNet.wskClient.State <> 0 Then 'If online...should i go off?
    If MsgBox("End Online Game?", vbYesNo, "Terminate Session?") = vbNo Then Exit Sub
    frmNet.wskClient.Close
End If
If frmNet.wskServer.State <> 0 Then
    If MsgBox("End Online Game?", vbYesNo, "Terminate Session?") = vbNo Then Exit Sub
    frmNet.wskServer.Close
End If
'Set up game to offline defualts
cmdOnline.Visible = False
tmrSync.Enabled = False
Text1.Visible = True
Text2.Visible = True
Label1.Visible = True
Label2.Visible = True
mnuMultiConnect.Enabled = True
mnuMultiHost.Enabled = True
frmHost.cmdListen.Caption = "Listen"
Command1.Visible = True
Command1.Caption = "Start!"
Picture1.Visible = False
Text1.SetFocus
End Sub

Private Sub mnuGameOptions_Click()
frmOptions.Show 1, Me 'Set up game options
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show 1, Me 'Shoe help
End Sub



Private Sub mnuHelpFAQ_Click()
'Load help FAQ
On Error GoTo 5555
frmHelp.rtfHelp.LoadFile App.Path & "\faq.rtf"
frmHelp.Show
Exit Sub
5555
MsgBox "Help File not Found!", vbCritical, "No Help!"
End Sub

Private Sub mnuHelpPlay_Click()
On Error GoTo 5555
frmHelp.rtfHelp.LoadFile App.Path & "\howtoplay.rtf"
frmHelp.Show
Exit Sub
5555
MsgBox "Help File not Found!", vbCritical, "No Help!"
End Sub

Private Sub mnuHelpRules_Click()
On Error GoTo 5555
frmHelp.rtfHelp.LoadFile App.Path & "\rules.rtf"
frmHelp.Show
Exit Sub
5555
MsgBox "Help File not Found!", vbCritical, "No Help!"
End Sub

Private Sub mnuMultiChat_Click()
frmNet.Visible = True 'Load chat room
frmNet.SetFocus
End Sub

Private Sub mnuMultiChatServ_Click()
'Find away to go to the online chat server
On Error GoTo 1039
Shell "start http://idots.cjb.net/", vbHide
Exit Sub
1039
On Error GoTo 22
Shell "cmd /c start http://idots.cjb.net/", vbHide
Exit Sub
22
MsgBox "Could not start web browser, visit http://iDots.Cjb.net/ for the chat server", vbInformation, "No Default web browser found"
End Sub

Private Sub mnuMultiConnect_Click()
frmConnect.Show 1, Me 'Connectdlg
End Sub

Private Sub mnuMultiHost_Click()
frmHost.Show 1, Me 'serverdlg
End Sub

Private Sub mnuOnlineDisconnect_Click()
'Disconnect game
If MsgBox("Leave Online Game?", vbYesNo, "Disconnect?") = vbNo Then Exit Sub
If frmNet.wskClient.State <> 0 Then
    frmNet.SendData 1, "discon:" 'send disconnect message to other computer
    frmNet.wskClient.Close
End If
If frmNet.wskServer.State <> 0 Then
    frmNet.SendData 0, "discon:"
    frmNet.wskServer.Close
End If
mnuFileNew_Click
End Sub

Private Sub mnuOnlineNew_Click()
frmOnline.Show 1, Me
End Sub

Public Sub mnuOnlineSynch_Click()
'Resynchronize game incase of screwy happenings
'ie they get a box where there is no possible way
On Error GoTo 12
Dim sT As String
sT = Me.ReSyncPacket_Creation
Dim Service As Integer
Service = -1
If frmNet.wskClient.State <> 0 Then Service = 1
If frmNet.wskServer.State <> 0 Then Service = 0
    frmReSync.Show , Me 'Hide form and show resync
    AlwaysOnTop frmReSync, True
    Me.Hide
    If Service = 0 Then
        frmNet.SendData 0, "resyn1:" & turn 'Packet1
        frmNet.SendData 0, "resyn2:" & Text1.Text 'Packet2
        frmNet.SendData 0, "resyn3:" & Text2.Text 'Packet 3
        frmNet.SendData 0, "resyn4:" & sT 'Packet 4
    End If
    If Service = 1 Then
        frmNet.SendData 1, "askSYN:" 'request to resync
    End If
    
Exit Sub
12
AlwaysOnTop frmReSync, False
Me.Hide
frmMain.Show
MsgBox Err.Description
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
Form_KeyPress KeyAscii 'Any keys now go to the chatroom
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change mouse pointer
Picture1.MousePointer = 99
Dim i As Integer
For i = 0 To vLine.Count - 1
     If vLine(i).Tag = "" Then
        vLine(i).BorderColor = vbBlack
        vLine(i).Visible = False
    End If
Next
For i = 0 To hLine.Count - 1
    
    If hLine(i).Tag = "" Then
        hLine(i).BorderColor = vbBlack
        hLine(i).Visible = False
    End If
Next
NetMouseX = X 'was going to send values to the internet...too slow
NetMouseY = Y
End Sub


Private Sub tmrOn_Timer() 'Don't remember
Static turnOn As Integer
turnOn = turnOn + 1
If turnOn = 60 Then
    tmrOn.Enabled = False
    Me.Height = 7695
End If

End Sub

Private Sub tmrSync_Timer()
'See if resync is needed by sending resync packets and stuff
SyncCount = SyncCount + 1
If SyncCount > 4 Then
    If frmNet.wskClient.State <> 0 Then
        frmNet.SendData 1, "getRED:" & RedBoxCount
        frmNet.SendData 1, "getBLU:" & BlueBoxCount
        frmNet.SendData 1, "getTRN:" & turn
    End If
    If frmNet.wskServer.State <> 0 Then
        frmNet.SendData 0, "getRED:" & RedBoxCount
        frmNet.SendData 0, "getBLU:" & BlueBoxCount
        frmNet.SendData 0, "getTRN:" & turn
    End If
    SyncCount = 0
End If
'Every lot of time make sure the game is in sync...
'The game may be better with out this
End Sub

Private Sub tmrURL_Timer()
'Used to hide the ad if they clicked a different ad
'Too bad they dropped me from the service
'If Online Then
'tmrURL.Enabled = False
'lblURL.Visible = True
'End If
End Sub

Private Sub tturn_Timer()
'Chechk for winner and display amounts of boxes
Label3.Visible = Command1.Caption = "Restart!"
Dim Service As Integer
Service = -1
If frmNet.wskClient.State <> 0 Then
    Service = 1
    Picture1.MouseIcon = blue.Picture 'blue mouse
End If
If frmNet.wskServer.State <> 0 Then
    Service = 0
    Picture1.MouseIcon = red.Picture 'red mouse
End If
If Service <> -1 Then 'connection
    If turn <> Service Then Picture1.MouseIcon = noMouse.Picture
    blue.Visible = False
    red.Visible = False
    GoTo 83
End If


'No connection
If turn = 0 Then
    Picture1.MouseIcon = red.Picture
    red.Visible = False
    blue.Visible = True
End If
    
If turn = 1 Then
    Picture1.MouseIcon = blue.Picture
    blue.Visible = False
    red.Visible = True
End If

83
lblRedBox.Caption = "Red Box Count: " & RedBoxCount
lblBlueBox.Caption = "Blue Box Count: " & BlueBoxCount
Label3.Caption = "Boxes Left: " & (Box.Count - RedBoxCount - BlueBoxCount)
If Box.Count - RedBoxCount - BlueBoxCount = 0 Then
    If RedBoxCount = BlueBoxCount Then
        MsgBox "We Have a Tie!", vbOKOnly, "Tie Game"
        GoTo 2222
    End If
    If RedBoxCount > BlueBoxCount Then
        MsgBox "Red Player Won!", vbInformation, "Good Job Red!"
    Else
        MsgBox "Blue Player Won!", vbInformation, "Good Job Blue!"
2222
    End If
    RedBoxCount = 0
    BlueBoxCount = 0
    turn = 0
    tturn.Enabled = False
    If frmNet.wskServer.State <> 0 Then
        frmNet.Pause 0.7
        If MsgBox("Play another Game?" & vbCrLf & "If you press no, you can always click the Online Button and press new game", vbYesNo, "New Game") = vbYes Then mnuOnlineNew_Click
    End If
End If

    


End Sub

Public Sub makeVLine(Index As Integer)

If vLine(Index).Tag <> "" Then Exit Sub
'Check all sides to see a box, if the line is already there
'don't use up a turn


'On Error Resume Next
Dim BoxGet As Boolean
Dim clrColor As ColorConstants
If turn = 0 Then clrColor = vbRed
If turn = 1 Then clrColor = vbBlue
BoxGet = False
vLine(Index).Tag = "set" 'tells the computer the line is taken up
vLine(Index).BorderColor = clrColor
vLine(Index).Visible = True
If GetTop(Index) < Box.Count Then 'get each side and see if a box is formed
    If Not Box(GetTop(Index)).Visible Then
        If IsBoxed(Index, True) Then
            'MsgBox IsBoxed(index, True)
            Box(GetTop(Index)).Visible = True
            Box(GetTop(Index)).FillColor = clrColor
            BoxGet = True
            If turn = 0 Then RedBoxCount = RedBoxCount + 1
            If turn = 1 Then BlueBoxCount = BlueBoxCount + 1
        End If
    End If
End If
If Index > 0 And Index Mod iWidth <> 0 Then 'check all sides
    If Not Box(GetTop(Index - 1)).Visible Then
        If IsBoxed(Index - 1, True) Then
            Box(GetTop(Index - 1)).Visible = True
            Box(GetTop(Index - 1)).FillColor = clrColor
            BoxGet = True
            If turn = 0 Then RedBoxCount = RedBoxCount + 1
            If turn = 1 Then BlueBoxCount = BlueBoxCount + 1
        End If
    End If
End If
If BoxGet = False Then
    If turn = 0 Then turn = 1: GoTo 24
    If turn = 1 Then turn = 0
24
End If
  CliSend = False
'MsgBox index
End Sub

Private Sub vHover_Click(Index As Integer)
'If online send message that it was clicked otherwise just draw it
'I could have made this more efficent by seeing if it was a valid line b4 sending
Dim Service As Integer
Service = -1
If frmNet.wskClient.State <> 0 Then Service = 1
If frmNet.wskServer.State <> 0 Then Service = 0
If Service <> -1 Then 'connected
    If turn <> Service Then Exit Sub
    frmNet.SendData Service, "vClick:" & Index
End If
makeVLine Index
End Sub

Private Sub vHover_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If it has not been set show it and then hide it
Picture1_MouseMove Button, Shift, vHover(Index).Left, vHover(Index).Left
If vLine(Index).Tag = "" Then
    vLine(Index).BorderColor = vbBlack
    vLine(Index).Visible = True
End If

End Sub
Public Sub makeHLine(Index As Integer)
'See vertical makeVLine code
If hLine(Index).Tag <> "" Then Exit Sub
'On Error Resume Next



Dim BoxGet As Boolean
Dim clrColor As ColorConstants
If turn = 0 Then clrColor = vbRed
If turn = 1 Then clrColor = vbBlue
BoxGet = False
hLine(Index).Tag = "set"
hLine(Index).Visible = True
hLine(Index).BorderColor = clrColor
'MsgBox index
'MsgBox GetLeft(index)
If Index < Box.Count Then
    If Not Box(Index).Visible Then
        If IsBoxed(Index, False) Then
            'MsgBox IsBoxed(index, False)
            Box(Index).Visible = True
            Box(Index).FillColor = clrColor
            If turn = 0 Then RedBoxCount = RedBoxCount + 1
            If turn = 1 Then BlueBoxCount = BlueBoxCount + 1
            BoxGet = True
        End If
    End If
End If
If Index > iWidth - 2 Then
    If Not Box(Index - iWidth + 1).Visible = True Then
        If IsBoxed(Index - iWidth + 1, False) Then
            'MsgBox IsBoxed(index, False)
            Box(Index - iWidth + 1).Visible = True
            Box(Index - iWidth + 1).FillColor = clrColor
            If turn = 0 Then RedBoxCount = RedBoxCount + 1
            If turn = 1 Then BlueBoxCount = BlueBoxCount + 1
            BoxGet = True
        End If
    End If
End If
If BoxGet = False Then
If turn = 0 Then turn = 1: GoTo 24
If turn = 1 Then turn = 0
24
End If
CliSend = False
End Sub

Private Sub hHover_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1_MouseMove Button, Shift, hHover(Index).Left, hHover(Index).Top
If hLine(Index).Tag = "" Then
    hLine(Index).BorderColor = vbBlack
    hLine(Index).Visible = True
End If
End Sub

Private Function GetTop(Index As Integer) As Integer
'Location, Get from Left
'   - <--
'  | |
'   -

'GetTop = Fix((index * (iWidth - 1)) / iWidth)
GetTop = Index - Fix(Index / iWidth) '<,-working
'Fix((iWidth / (iWidth - 1)) * index)
End Function
Private Function GetBottom(Index As Integer) As Integer
'Location, Get from Left
'   -
'  | |
'   -<--
GetBottom = GetTop(Index) + iWidth - 1
End Function
Private Function GetLeft(Index As Integer) As Integer
'Location From Top
'   -
'->| |
'   -

GetLeft = Fix((iWidth / (iWidth - 1)) * Index)
End Function
Private Function GetRight(Index As Integer) As Integer
'Location, From Top
'   -
'  | |<--
'   -
GetRight = GetLeft(Index) + 1
End Function
Private Function IsBoxed(Index As Integer, vert As Boolean) As Boolean
Dim Checker As Boolean, Loc As Integer, Els As Integer
Dim Index2 As Integer
Index2 = Index
Checker = False
'get the location
Loc = -1
If vert = True Then
    If (Index Mod iWidth) = 0 Then Loc = 0 'Left
    If ((Index + 1) Mod iWidth) = 0 Then Loc = 1 'right
    Els = 0
    If Index = 0 Then Loc = 0
End If
If vert = False Then
    If Index < (iWidth - 1) Then Loc = 2 'Top
    If Index > (hLine.Count - iWidth) Then Loc = 3
    Els = 1
End If

Select Case Loc
    Case 0 'left
        Checker = (vLine(Index).Tag <> "" And vLine(Index + 1).Tag <> "" And hLine(GetTop(Index)).Tag <> "" And hLine(GetBottom(Index)).Tag <> "")
    Case 1 'right
        Checker = (vLine(Index).Tag <> "" And hLine(GetTop(Index - 1)).Tag <> "" And hLine(GetBottom(Index - 1)).Tag <> "" And vLine(Index - 1).Tag <> "")
        If Checker Then Index = Index - 1
    Case 2 'top
        Checker = (hLine(Index).Tag <> "" And hLine(GetBottom(GetLeft(Index))).Tag <> "" And vLine(GetLeft(Index)).Tag <> "" And vLine(GetRight(Index)).Tag <> "")
    Case 3 'bottom
        Checker = (hLine(Index).Tag <> "" And hLine(Index - iWidth + 1).Tag <> "" And vLine(GetLeft(Index - iWidth + 1)).Tag <> "" And vLine(GetRight(Index - iWidth + 1)).Tag <> "")
        If Checker Then Index = Index - iWidth + 1
    Case Else
        
        Dim boxy As Integer
        If Els = 0 Then
            Checker = (vLine(Index).Tag <> "" And vLine(Index + 1).Tag <> "" And hLine(GetTop(Index)).Tag <> "" And hLine(GetBottom(Index)).Tag <> "")
        End If
        If Els = 1 Then
            Checker = (hLine(Index).Tag <> "" And vLine(GetRight(Index)).Tag <> "" And vLine(GetLeft(Index)).Tag <> "" And hLine(GetBottom(GetLeft(Index))).Tag <> "")
        End If
    End Select
    'index = Index2

                
    'if all the way left
    'if all the way right
    'if at the top
    'if at the bottom
    'anywhere else

IsBoxed = Checker
End Function


Sub GetReg()

'Just loads background and color settings
If ((GetSetting(App.Title, "Settings", "Image", "") = "") And (GetSetting(App.Title, "Settings", "Backcolor", "") = "")) Then
    If UCase(Dir(App.Path + "\paper.gif")) = "PAPER.GIF" Then
        Picture1.Picture = LoadPicture(App.Path + "\paper.gif")
        Exit Sub
    End If
End If
If GetSetting(App.Title, "Settings", "Image", "") <> "" Then
    Picture1.Picture = LoadPicture(GetSetting(App.Title, "Settings", "Image", ""))
End If
Picture1.BackColor = GetSetting(App.Title, "Settings", "Backcolor", Me.BackColor)
Picture1.BorderStyle = GetSetting(App.Title, "Options", "Border", 1)
If GetSetting(App.Title, "Settings", "Image", "") = "" Then
    Picture1.Picture = LoadPicture("")
End If

'First Check Day...if same Day and Month AD!
Dim d As Integer, h As Integer, m As Integer




End Sub
Public Sub clientV(Index As Integer)
vHover_Click Index 'if over the internet it was told to click
End Sub
Public Sub clientH(Index As Integer)
hHover_Click Index
End Sub

Public Function ReSyncPacket_Creation() As String
On Error GoTo 12
Dim i As Integer, Pack As String
Pack = ""
For i = 0 To vLine.Count - 1
    If vLine(i).Tag <> "" Then Pack = Pack + "1" Else Pack = Pack + "0"
Next
For i = 0 To hLine.Count - 1
    If hLine(i).Tag <> "" Then Pack = Pack + "1" Else Pack = Pack + "0"
Next
For i = 0 To Box.Count - 1
    If Box(i).Visible = False Then
        Pack = Pack + "0"
    Else
        If Box(i).FillColor = vbRed Then Pack = Pack + "1"
        If Box(i).FillColor = vbBlue Then Pack = Pack + "2"
    End If
Next
ReSyncPacket_Creation = Pack
Exit Function
12
MsgBox Err.Description
'This code basically goes through the entire board and decides who has a line
'depending on that it will (from server-->client) recreate the board
End Function
Public Sub ReSyncGame(Resync_Packet As String)

Dim strV As String, strH As String, strBOX As String
'Break the resync packet up
strV = Left(Resync_Packet, vLine.Count)
strH = Mid(Resync_Packet, vLine.Count + 1, hLine.Count)
strBOX = Right(Resync_Packet, Box.Count)
'Split up packages = Merged ones!!! Booya
'GOOD UP TO HERE...overflow occurs somewhere down
Dim q As String 'query string
Dim s As Integer 'what it equals
Dim i As Integer
For i = 0 To vLine.Count - 1
    
    q = Mid(strV, i + 1, 1)

    s = Val(q)

    If s = 0 Then
        With vLine(i)
            .Visible = False
            .Tag = ""
        End With
    End If
    If s = 1 Then
        With vLine(i)
            .Visible = True
            .Tag = "set"
            .BorderColor = vbBlack
        End With
    End If
Next

For i = 0 To hLine.Count - 1
    q = Mid(strH, i + 1, 1)
    s = Val(q)
    If s = 0 Then
        With hLine(i)
            .Visible = False
            .Tag = ""
        End With
    End If
    If s = 1 Then
        With hLine(i)
            .Visible = True
            .Tag = "set"
            .BorderColor = vbBlack
        End With
    End If
Next

For i = 0 To Box.Count - 1
    q = Mid(strBOX, i + 1, 1)
    s = Val(q)
    If s = 0 Then
        With Box(i)
            .Visible = False
            .Tag = ""
        End With
    End If
    If s = 1 Then
        With Box(i)
            .Visible = True
            .Tag = "set"
            .FillColor = vbRed
        End With
    End If
    If s = 2 Then
        With Box(i)
            .Visible = True
            .Tag = "set"
            .FillColor = vbBlue
        End With
    End If
Next
'Takes line piece by piece and remakes the board on the
'client based on what the server sends
End Sub

