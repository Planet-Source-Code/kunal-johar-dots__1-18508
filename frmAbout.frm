VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Dots"
   ClientHeight    =   3915
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2702.203
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   5535
      TabIndex        =   8
      ToolTipText     =   "Extra thanks to DosFX for LaserDraw and PSC for Coolmenu"
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Timer tmrCredits 
      Interval        =   100
      Left            =   210
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   5475
      TabIndex        =   7
      Top             =   1800
      Width           =   5535
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2865
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   3315
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5521.624
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblDescription 
      Caption         =   "The addictive 2-player game, now brought to the computer.  Get more boxes than your friend!  Now even with online support!"
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   1050
      TabIndex        =   3
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "BootlegZ  R Us Presents: Dots!"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   2325
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "This product is now freeware!"
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   255
      TabIndex        =   4
      Top             =   2745
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'All of this is for the Laser Draw Function -NOT MINE
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Dim CreditLevel As Integer
Private Const SRCCOPY = &HCC0020
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Me.Hide
End Sub



Public Sub StartSysInfo() 'All from Microsoft's Defualt Help thing
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Private Sub Form_Activate()
tmrCredits.Enabled = True 'Start the Lazer Draw
tmrCredits.Interval = 100
CreditLevel = 0 'Start at 0th Message
End Sub

Private Sub Form_Deactivate()
tmrCredits.Enabled = False 'Helped trap errors, don't ask
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTitle.ForeColor = vbBlack
End Sub

Private Sub lblTitle_Click() 'Probably an inefficent way of making sure the website will load
On Error GoTo 1039
Shell "start http://www.BootlegZRUs.com", vbHide
Exit Sub
1039
On Error GoTo 22
Shell "cmd /c start http://www.BootlegZRUs.com", vbHide
Exit Sub
22
MsgBox "Could not start web browser, visit http://www.BootlegZRUs.com for the chat server", vbInformation, "No Default webbrowser found"
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTitle.ForeColor = vbBlue 'Make it look like Hyperlink
End Sub

Private Sub tmrCredits_Timer()
DoEvents 'Yield to otherevents
Picture1.Cls 'Clear old boxes
Picture2.Cls
Select Case CreditLevel
    Case 0
        tmrCredits.Interval = 100
        'WriteInPic (Made by Me) is just an easy way of drawing what will be lasered easily
        WriteInPic Picture1, "BootlegZ R Us", vbBlack, vbGreen
    Case 1
        WriteInPic Picture1, "and", vbWhite, vbBlack
    Case 2
        WriteInPic Picture1, "PartiXx Media Publishing", vbBlack, vbGreen
    Case 3
        WriteInPic Picture1, "Present", vbCyan, vbMagenta
    Case 4
        WriteInPic Picture1, "DOTS", vbBlack, vbBlack
        tmrCredits.Interval = 1000
    Case 5
        tmrCredits.Interval = 300
        WriteInPic Picture1, "Development Team", vbGreen, vbYellow
    Case 6
        WriteInPic Picture1, "Lead Programmer: Kunal Johar", vbRed, vbWhite
        tmrCredits.Interval = 2000
    Case 7
        tmrCredits.Interval = 300
        WriteInPic Picture1, "Beta Testers...", vbGreen, vbBlack
    Case 8
        WriteInPic Picture1, "James Wan", vbYellow, vbBlue
    Case 9
        WriteInPic Picture1, "Evan Hoberman", vbYellow, vbBlue
    Case 10
        WriteInPic Picture1, "Rebecca Schoer", vbYellow, vbBlue
    Case 11
        WriteInPic Picture1, "Mathew Porcu (yes 1 T)", vbYellow, vbBlue
        tmrCredits.Interval = 570
    Case 12
        WriteInPic Picture1, "Steven Lake", vbYellow, vbBlue
        tmrCredits.Interval = 300
    Case 13
        WriteInPic Picture1, "Samantha Schoer", vbYellow, vbBlue
    Case 14
        WriteInPic Picture1, "Craig Cohen", vbYellow, vbBlue
    Case 15
        WriteInPic Picture1, "Jon Mottahedeh", vbYellow, vbBlue
    Case 16
        WriteInPic Picture1, "Thank You for Playing!", vbRed, vbGreen
        tmrCredits.Interval = 700
    Case 17
        WriteInPic Picture1, "Check out our other Software!"
        tmrCredits.Interval = 1010
    Case 18
        WriteInPic Picture1, "Visit Us Online @", vbWhite, vbBlue
        tmrCredits.Interval = 900
    Case 19
        WriteInPic Picture1, "www.BootlegZRUs.com", vbBlue, vbWhite
        tmrCredits.Interval = 700
    Case 20
        tmrCredits.Interval = 8000
        WriteInPic Picture1, "Now go Play!", vbMagenta, vbWhite
        
End Select

CreditLevel = CreditLevel + 1 ' Go to the next Message
If CreditLevel > 20 Then CreditLevel = 0 'If hit end start over
LaserDraw Picture2, Picture1, , , , , , 0.02 'Laser Draw function, not by me
End Sub
Public Sub LaserDraw(picDest As PictureBox, picSource As PictureBox, Optional lngTrans As Long = -1, Optional lngWidth As Long, Optional lngHeight As Long, Optional lngLaserX As Long = 0&, Optional lngLaserY As Long = 0&, Optional sngTimeout As Single = 0.01)
DoEvents
  'Arguments
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'picDest
  '   Picture box to draw image on.
  'picSource
  '   Picture box where the image is located.
  'lngTrans (Optional)
  '   The transparent color. If this is not given,
  '   then the transparent color is the color
  '   of the pixel at 0,0
  'lngWidth (Optional)
  '   Width in pixels of picSource to draw.
  'lngHeight
  '   Height in pixels of picSource to draw.
  'lngLaserX
  '   X coordinate where the top of the laser starts.
  'lngLaserY
  '   Y coordinate where the top of the laser starts.
  'sngTimeout
  '   Delay in seconds to wait before moving on to the
  '   next row of pixels. This determines how fast it
  '   draws the picture.


  Dim lngX As Long, lngY As Long, lngColor As Long
  Dim sngTimer As Single
    
  picDest.ScaleMode = vbPixels
  picDest.AutoRedraw = True
  
  If lngWidth& = 0 Then lngWidth& = picDest.ScaleWidth - 1
  If lngHeight& = 0 Then lngHeight& = picDest.ScaleHeight - 1
   
  ''If no transparent color is specified, get the transparent color
  If lngTrans& < 0 Then lngTrans& = GetPixel(picSource.hdc, 0&, 0&)

     For lngY& = 1 To lngHeight&
          For lngX& = 1 To lngWidth&
             lngColor& = GetPixel(picSource.hdc, lngX&, lngY&)
             
                If lngColor& <> lngTrans& And lngColor& >= 0 Then
                    ''Draw line (laser)
                    picDest.Line (lngLaserX&, lngLaserY&)-(lngX&, lngY&), lngColor&
                End If
             
          Next lngX&
          
       ''Pause for sngTimeout second(s)
       sngTimer! = Timer
       
         Do While Timer - sngTimer! < sngTimeout!
            DoEvents
         Loop
       
       ''Clear picturebox
       Call picDest.Cls
       
       ''Copy part of image up to the pixel but not including lngY&
       Call BitBlt(picDest.hdc, 0&, 0&, lngWidth&, lngY& - 1, picSource.hdc, 0&, 0&, SRCCOPY)
     Next lngY&
     
End Sub
Private Sub WriteInPic(PicWrite As Object, TextWrite As String, Optional Color1 As ColorConstants = vbBlue, Optional Color2 As ColorConstants = vbRed)
DoEvents
   With PicWrite 'Set up a picture box to have a message in it
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .FontName = "Tahoma"
      .FontSize = 14
      .FontBold = True
      .CurrentY = 2
      .CurrentX = (.ScaleWidth / 2 - .TextWidth(TextWrite) / 2) + 2
      .ForeColor = Color1
      PicWrite.Print TextWrite
      .CurrentY = 2
      .CurrentX = (.ScaleWidth / 2 - .TextWidth(TextWrite) / 2) + .TextWidth("Laser ") + 2
      .ForeColor = vbBlue

      .CurrentY = 0
      .CurrentX = .ScaleWidth / 2 - .TextWidth(TextWrite) / 2
      .ForeColor = Color2
      PicWrite.Print TextWrite
      .CurrentY = 0
      .CurrentX = (.ScaleWidth / 2 - .TextWidth(TextWrite) / 2) + .TextWidth("Laser ")
      .ForeColor = vbBlack
   End With

End Sub
