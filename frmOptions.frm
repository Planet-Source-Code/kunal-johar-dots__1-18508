VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Options"
   ClientHeight    =   1995
   ClientLeft      =   2760
   ClientTop       =   1455
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFile 
      Caption         =   "Image..."
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Color..."
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox chkBorder 
      Caption         =   "Have a border around Game Board"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton optImage 
      Caption         =   "Use Background Image"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Use Background Color"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   1920
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Game Board Options:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Self Explanatory Sub
'Saves and loads the options from the registry



Private Sub chkBorder_Click()

frmMain.Picture1.BorderStyle = chkBorder.Value


End Sub





Private Sub cmdColor_Click()
On Error GoTo 4141
Dim a As Long
a = frmMain.Picture1.BackColor
dlgMain.ShowColor
dlgMain.CancelError = True

frmMain.Picture1.BackColor = dlgMain.Color
Exit Sub
4141
frmMain.Picture1.BackColor = a
End Sub

Private Sub cmdFile_Click()
On Error GoTo 4141
dlgMain.ShowOpen
dlgMain.CancelError = True
frmMain.Picture1.Picture = LoadPicture(dlgMain.FileName)


4141
End Sub

Private Sub Form_Activate()
If frmMain.Picture1.Picture.Height <> 0 Then optImage.Value = True Else optColor.Value = True
chkBorder.Value = frmMain.Picture1.BorderStyle
End Sub

Private Sub Form_Deactivate()
SaveSetting App.Title, "Options", "Backcolor", optColor.Value
SaveSetting App.Title, "Options", "Image", optImage.Value
SaveSetting App.Title, "Options", "Border", chkBorder.Value
SaveSetting App.Title, "Settings", "Backcolor", frmMain.Picture1.BackColor
SaveSetting App.Title, "Settings", "Image", dlgMain.FileName

End Sub

Private Sub Form_LostFocus()
SaveSetting App.Title, "Options", "Backcolor", optColor.Value
SaveSetting App.Title, "Options", "Image", optImage.Value
SaveSetting App.Title, "Options", "Border", chkBorder.Value
SaveSetting App.Title, "Settings", "Backcolor", frmMain.Picture1.BackColor
SaveSetting App.Title, "Settings", "Image", dlgMain.FileName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSetting App.Title, "Options", "Backcolor", optColor.Value
SaveSetting App.Title, "Options", "Image", optImage.Value
SaveSetting App.Title, "Options", "Border", chkBorder.Value
SaveSetting App.Title, "Settings", "Backcolor", frmMain.Picture1.BackColor
SaveSetting App.Title, "Settings", "Image", dlgMain.FileName

End Sub

Private Sub optColor_Click()
cmdColor.Visible = optColor.Value
cmdFile.Visible = optImage.Value
frmMain.Picture1.Picture = LoadPicture("")
End Sub

Private Sub optImage_Click()
cmdColor.Visible = optColor.Value

cmdFile.Visible = optImage.Value
End Sub
