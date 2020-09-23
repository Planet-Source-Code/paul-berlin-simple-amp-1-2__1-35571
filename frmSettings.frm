VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   5025
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   6060
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDevice 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   6120
      TabIndex        =   21
      Top             =   360
      Width           =   5775
      Visible         =   0   'False
      Begin VB.Frame Frame4 
         Caption         =   "Device Settings"
         Height          =   3015
         Left            =   50
         TabIndex        =   24
         Top             =   840
         Width           =   5655
         Begin VB.ComboBox cmbHz 
            Height          =   315
            ItemData        =   "frmSettings.frx":08CA
            Left            =   120
            List            =   "frmSettings.frx":08D4
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmbStereo 
            Height          =   315
            ItemData        =   "frmSettings.frx":08EC
            Left            =   1680
            List            =   "frmSettings.frx":08F6
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmbBits 
            Height          =   315
            ItemData        =   "frmSettings.frx":0908
            Left            =   3240
            List            =   "frmSettings.frx":0912
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cmbBuffer 
            Height          =   315
            ItemData        =   "frmSettings.frx":0927
            Left            =   120
            List            =   "frmSettings.frx":0937
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1200
            Width           =   1575
         End
         Begin MSComctlLib.Slider sdrPanning 
            Height          =   375
            Left            =   1800
            TabIndex        =   25
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   10
            Min             =   -100
            Max             =   100
            TickFrequency   =   10
         End
         Begin VB.Label Label9 
            Caption         =   $"frmSettings.frx":095A
            Height          =   480
            Left            =   120
            TabIndex        =   34
            Top             =   1920
            Width           =   5295
         End
         Begin VB.Label Label10 
            Caption         =   "Note: If mixing quality settings are unavailabe, try changing to an different sound device and restart the program."
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   2400
            Width           =   5415
         End
         Begin VB.Label Label6 
            Caption         =   "Mixing Quality:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Sound Buffer:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Sound Panning:"
            Height          =   255
            Left            =   1800
            TabIndex        =   30
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbDevices 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label5 
         Caption         =   "Sound Device:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame frmGeneral 
      BorderStyle     =   0  'None
      Caption         =   "General"
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5775
      Begin VB.Frame Frame1 
         Caption         =   "General Settings"
         Height          =   2055
         Left            =   50
         TabIndex        =   15
         Top             =   50
         Width           =   3975
         Begin VB.CheckBox chkSnap 
            Caption         =   "Snap &windows to screen edges"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   3735
         End
         Begin VB.CheckBox chkAlwaysTray 
            Caption         =   "Always show &tray icon"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   3735
         End
         Begin VB.CheckBox chkMinimizeTray 
            Caption         =   "&Minimize to tray instead of taskbar"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   3735
         End
         Begin VB.CheckBox chkAlwaysOntop 
            Caption         =   "&Always on top"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   3735
         End
         Begin VB.CheckBox chkNoID3 
            Caption         =   "Do &not mp3 read data when adding to playlist"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CheckBox chkStartIntray 
            Caption         =   "&Start in tray"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   3735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tray Icon"
         Height          =   2055
         Left            =   4110
         TabIndex        =   13
         Top             =   50
         Width           =   1575
         Begin MSComctlLib.Slider Slider1 
            Height          =   1455
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   2566
            _Version        =   393216
            Orientation     =   1
            LargeChange     =   2
            Min             =   1
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   840
            Picture         =   "frmSettings.frx":09F0
            Top             =   600
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   960
            Picture         =   "frmSettings.frx":12BA
            Stretch         =   -1  'True
            Top             =   1320
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Hotkey Actions"
         Height          =   1695
         Left            =   50
         TabIndex        =   3
         Top             =   2160
         Width           =   5655
         Begin VB.ComboBox cmdHotkey 
            Height          =   315
            ItemData        =   "frmSettings.frx":1B84
            Left            =   120
            List            =   "frmSettings.frx":1B97
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmdActive 
            Height          =   315
            ItemData        =   "frmSettings.frx":1BCD
            Left            =   1680
            List            =   "frmSettings.frx":1BE0
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox cmdInactive 
            Height          =   315
            ItemData        =   "frmSettings.frx":1C1F
            Left            =   3720
            List            =   "frmSettings.frx":1C2C
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtLocation 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1080
            Width           =   4095
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   4440
            TabIndex        =   4
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Hotkey:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Action when active:"
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Action when inactive:"
            Height          =   255
            Left            =   3720
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Program loaction:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1335
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sound &Device"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Program 
      Left            =   3600
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Program"
      Filter          =   "Programs (*.exe;*.com;*.bat)|*.exe;*.bat;*.com"
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActive_Click()
  Hotkey(cmdHotkey.ListIndex + 1).Active = cmdActive.ListIndex
End Sub

Private Sub cmdBrowse_Click()
  Program.ShowOpen
  If Program.FileName <> "" Then
    txtLocation = Program.FileName
    txtLocation.SelStart = Len(txtLocation)
    Hotkey(cmdHotkey.ListIndex + 1).Program = Program.FileName
  End If
End Sub

Private Sub cmdClose_Click()
  'Update variables
  AlwaysTray = CBool(chkAlwaysTray)
  MinimizeTray = CBool(chkMinimizeTray)
  OnTop = CBool(chkAlwaysOntop)
  NoID3 = CBool(chkNoID3)
  StartInTray = CBool(chkStartIntray)
  Snap = CBool(chkSnap)
  devNum = cmbDevices.ListIndex - 1
  dev44100 = CBool(cmbHz.ListIndex)
  devStereo = CBool(cmbStereo.ListIndex)
  dev16bits = CBool(cmbBits.ListIndex)
  Select Case cmbBuffer.ListIndex
    Case 0: devBuffer = 0.5
    Case 1: devBuffer = 1
    Case 2: devBuffer = 1.5
    Case 3: devBuffer = 2
  End Select
  devPanning = sdrPanning.Value
  
  frmMain.DMC.BufferLenInSeconds = devBuffer
  If OnTop Then
    AlwaysOnTop frmMain, True
    If frmPlaylist.Visible Then AlwaysOnTop frmPlaylist, True
  Else
    AlwaysOnTop frmMain, False
    If frmPlaylist.Visible Then AlwaysOnTop frmPlaylist, False
  End If
  If AlwaysTray Then
    frmMain.SetupSystray True
  Else
    frmMain.SetupSystray False
  End If
  
  frmMain.SaveSettings
  Hide
End Sub

Private Sub cmdHotkey_Click()
  cmdActive.ListIndex = Hotkey(cmdHotkey.ListIndex + 1).Active
  cmdInactive.ListIndex = Hotkey(cmdHotkey.ListIndex + 1).Inactive
  If Hotkey(cmdHotkey.ListIndex + 1).Inactive = 3 Then txtLocation = Hotkey(cmdHotkey.ListIndex + 1).Program
End Sub

Private Sub cmdInactive_Click()
  Hotkey(cmdHotkey.ListIndex + 1).Inactive = cmdInactive.ListIndex
  If cmdInactive.ListIndex = 2 Then
    txtLocation = Hotkey(cmdHotkey.ListIndex + 1).Program
    txtLocation.SelStart = Len(txtLocation)
    cmdBrowse.Enabled = True
  Else
    txtLocation = ""
    cmdBrowse.Enabled = False
  End If
End Sub

Private Sub Form_Activate()
  'Update controls
  Dim X As Long
  
  frmDevice.Top = frmGeneral.Top
  frmDevice.Left = frmGeneral.Left
  TabStrip1.Tabs(1).Selected = True
  
  If OnTop Then
    AlwaysOnTop Me, True
  Else
    AlwaysOnTop Me, False
  End If
  
  cmbDevices.Clear
  cmbDevices.AddItem "<Default Device>"
  For X = 0 To frmMain.DMC.Info_DeviceCount - 1
    cmbDevices.AddItem frmMain.DMC.Info_DeviceDescrip(X)
  Next X
  cmbDevices.ListIndex = devNum + 1
  
  cmbHz.ListIndex = Abs(dev44100)
  cmbStereo.ListIndex = Abs(devStereo)
  cmbBits.ListIndex = Abs(dev16bits)
  Select Case devBuffer
    Case 0.5: cmbBuffer.ListIndex = 0
    Case 1: cmbBuffer.ListIndex = 1
    Case 1.5: cmbBuffer.ListIndex = 2
    Case 2: cmbBuffer.ListIndex = 3
  End Select
  sdrPanning.Value = devPanning
  
  If frmMain.DMC.Info_HW16bitSupport And Not frmMain.DMC.Info_HW8bitSupport Then
    cmbBits.ListIndex = 1
    cmbBits.Enabled = False
  ElseIf frmMain.DMC.Info_HW8bitSupport And Not frmMain.DMC.Info_HW16bitSupport Then
    cmbBits.ListIndex = 0
    cmbBits.Enabled = False
  ElseIf Not frmMain.DMC.Info_HW8bitSupport And Not frmMain.DMC.Info_HW16bitSupport Then
    cmbBits.ListIndex = 1
    cmbBits.Enabled = False
  End If
  If Not frmMain.DMC.Info_SampRatesSupport Then
    cmbHz.Enabled = False
  End If
  If frmMain.DMC.Info_HWStereoSupport And Not frmMain.DMC.Info_HWMonoSupport Then
    cmbStereo.ListIndex = 1
    cmbStereo.Enabled = False
  ElseIf frmMain.DMC.Info_HWMonoSupport And Not frmMain.DMC.Info_HWStereoSupport Then
    cmbStereo.ListIndex = 0
    cmbStereo.Enabled = False
  ElseIf Not frmMain.DMC.Info_HWMonoSupport And Not frmMain.DMC.Info_HWStereoSupport Then
    cmbStereo.ListIndex = 1
    cmbStereo.Enabled = False
  End If
  
  chkAlwaysTray.Value = Abs(AlwaysTray)
  chkMinimizeTray.Value = Abs(MinimizeTray)
  chkAlwaysOntop.Value = Abs(OnTop)
  chkNoID3.Value = Abs(NoID3)
  chkStartIntray.Value = Abs(StartInTray)
  chkSnap.Value = Abs(Snap)
  Slider1.Value = TrayIcon
  Image1.Picture = frmMain.imlIcons.ListImages(TrayIcon).Picture
  Image2.Picture = frmMain.imlIcons.ListImages(TrayIcon).Picture
  
  cmdHotkey.ListIndex = 0

End Sub

Private Sub sdrPanning_Scroll()
  frmMain.DMC.StreamPan = sdrPanning.Value
End Sub

Private Sub Slider1_Scroll()
  'Update tray icon
  Image1.Picture = frmMain.imlIcons.ListImages(Slider1.Value).Picture
  Image2.Picture = frmMain.imlIcons.ListImages(Slider1.Value).Picture
  TrayIcon = Slider1.Value
  frmMain.SetupSystrayIcon TrayIcon
End Sub

Private Sub TabStrip1_Click()
  Select Case TabStrip1.SelectedItem.Index
  Case 1:
    frmGeneral.Visible = True
    frmDevice.Visible = False
  Case 2:
    frmGeneral.Visible = False
    frmDevice.Visible = True
  End Select
End Sub
