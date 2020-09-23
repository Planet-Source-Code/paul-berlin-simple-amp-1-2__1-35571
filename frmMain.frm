VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2C1EC115-F1BA-11D3-BF43-00A0CC32BE58}#9.1#0"; "DMC2.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Simple Amp 1.1"
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "SimpleAmp"
   MaxButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   Begin VB.Timer tmrCounter 
      Interval        =   1000
      Left            =   4080
      Top             =   2160
   End
   Begin VB.PictureBox Keys 
      Height          =   255
      Left            =   5280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox SpectrumBar 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
      Visible         =   0   'False
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2640
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5874
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrMain 
      Interval        =   10
      Left            =   3840
      Top             =   2160
   End
   Begin SimpleAmp.PicScroll Volume 
      Height          =   180
      Left            =   4095
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Volume"
      Top             =   1140
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   318
      Value           =   50
   End
   Begin DMC2.DMC DMC 
      Left            =   3240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox PictureLoader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
      Visible         =   0   'False
   End
   Begin SimpleAmp.PicScroll Position 
      Height          =   255
      Left            =   180
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Position"
      Top             =   1680
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   450
      Max             =   1
   End
   Begin VB.PictureBox PicSpectrum1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1455
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   450
      ScaleWidth      =   2340
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click to change scope"
      Top             =   1125
      Width           =   2340
   End
   Begin VB.Label lblComments 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   165
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   795
      UseMnemonic     =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgDropdown 
      Height          =   735
      Left            =   98
      OLEDropMode     =   1  'Manual
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lblGenre 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      Height          =   195
      Left            =   3825
      TabIndex        =   13
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgStereoMono 
      Height          =   255
      Left            =   840
      ToolTipText     =   "Stereo/Mono"
      Top             =   1365
      Width           =   495
   End
   Begin VB.Image imgShuffle 
      Height          =   255
      Left            =   4095
      ToolTipText     =   "Shuffle On/Off"
      Top             =   1365
      Width           =   255
   End
   Begin VB.Image imgRepeat 
      Height          =   255
      Left            =   3840
      ToolTipText     =   "Repeat On/Off"
      Top             =   1365
      Width           =   255
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Mp3 Info"
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Image imgPlaylist 
      Height          =   375
      Left            =   4560
      ToolTipText     =   "Show/Hide Playlist"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgNext 
      Height          =   375
      Left            =   2040
      ToolTipText     =   "Next"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgStop 
      Height          =   375
      Left            =   1440
      ToolTipText     =   "Stop"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgPlayPause 
      Height          =   375
      Left            =   840
      ToolTipText     =   "Play/Pause"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgPrev 
      Height          =   375
      Left            =   240
      ToolTipText     =   "Previous"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgMinimize 
      Height          =   255
      Left            =   4680
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   4965
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblYear 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4425
      TabIndex        =   7
      Top             =   405
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label lblAlbum 
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblTotalTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4425
      TabIndex        =   4
      ToolTipText     =   "Total Time"
      Top             =   1365
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   3
      ToolTipText     =   "Elapsed Time"
      Top             =   1365
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblArtistTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Artist & Title"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   405
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Image imgDropdown2 
      Height          =   2655
      Left            =   0
      OLEDropMode     =   1  'Manual
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'                      Simple Amp 1.2
'                    By Paul Berlin 2002
'                  berlin_paul@hotmail.com
'----------------------------------------------------------
' Some of the modules, classes & user controls in this
' program was NOT made by me, but some have been modified
' by me. Look at the credits in the about-box for more info.
'
' This program is Freeware. If you modify or inprove this
' program in any way it would be great if you could send it
' to me. Please mention me in your program credits if you
' use any of the code.
'
' Please vote at PlanetSourceCode if you like the program!
'
' Read readme.txt for quick help & more information.
'
' Note: This code can be a bit messy sometimes and some
'       areas do not have very much comments, but i hope
'       you will understand it anyway.

Option Explicit
Private WithEvents cTray As clsSysTray
Attribute cTray.VB_VarHelpID = -1
Dim ScopeArray(19) As Long
Dim ExitNow As Boolean
Dim MoveMain As Boolean 'True if in move mode
Dim MoveMainOldX As Long
Dim MoveMainOldY As Long

Public Sub cTray_LButtonDblClk()
  If Visible = False Then
    Show
    If OnTop Then AlwaysOnTop Me, True
    If frmPlaylist.TaG = "1" Then
      If OnTop Then AlwaysOnTop frmPlaylist, True
      frmPlaylist.Show
      frmPlaylist.TaG = ""
    End If
    If Not AlwaysTray Then cTray.RemoveFromSysTray
  End If
End Sub

Private Sub cTray_RButtonUp()
  PopupMenu frmMenus.menSystray
End Sub

Private Sub DMC_StreamStoped(ByVal paused As Boolean)
  If DMC.StreamLen <> -1 Then
    If DMC.StreamPos = DMC.StreamLen Then
      If frmPlaylist.lvwList.ListItems.Count > 0 Then
        PlayNext
      Else
        imgStop_Click
      End If
    End If
  End If
End Sub

Private Sub Form_GotFocus()
  Keys.SetFocus
End Sub

Private Sub Form_Load()
  Dim x As Long, ListAdd, tmp1 As Long, tmp2 As Boolean, tmp3 As Boolean
  Set cTray = New clsSysTray

  Caption = "Simple Amp " & App.Major & "." & App.Minor
  Load frmDDE
  
  'Load settings
  LoadSettings
  
  NumStarted = NumStarted + 1
  
  Volume.Value = CurrentVolume
  
  ReDim Shuff(0) As Long
  ReDim Playlist(0) As PlaylistData
  
  'Get windows work area size to type variable Scr
  SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
  
  'Setup systray
  Set cTray.SourceWindow = Me
  cTray.Icon = imlIcons.ListImages(TrayIcon).Picture  'Set icon
  cTray.ToolTip = "Simple Amp " & App.Major & "." & App.Minor 'Set tip text
  cTray.DefaultDblClk = False
  If AlwaysTray Then cTray.IconInSysTray
  
  'Loads skin
  LoadSkin CurrentSkin
  
  If StartInTray Then
    imgMinimize_Click
  Else
    'Sets Ontop
    If OnTop Then
      AlwaysOnTop Me, True
    End If
  End If
  
  'Init sound
  DMC.DeviceToUse = devNum
  
  If dev44100 Then
    tmp1 = 44100
  Else
    tmp1 = 22050
  End If
  If devStereo Then
    tmp2 = False
  Else
    tmp2 = True
  End If
  If dev16bits Then
    tmp3 = False
  Else
    tmp3 = True
  End If
  DMC.InitBASS Me.hwnd, tmp1, tmp2, tmp3
  DMC.BufferLenInSeconds = devBuffer
  DMC.StreamPan = devPanning
  
  'This loads the autosave playlist if there is one
  If FileExists(App.Path & "\current.playlist") Then
    If frmMenus.LoadPlaylist(App.Path & "\current.playlist") Then
      For x = 1 To UBound(Playlist)
      
        Set ListAdd = frmPlaylist.lvwList.ListItems.Add
      
        If Len(Playlist(x).Artist) > 0 And Len(Playlist(x).Title) > 0 Then
          ListAdd.Text = Playlist(x).Artist & " - " & Playlist(x).Title
        Else
          ListAdd.Text = Right(Playlist(x).Filename, Len(Playlist(x).Filename) - InStrRev(Playlist(x).Filename, "\"))
        End If
        ListAdd.SubItems(1) = Playlist(x).Album
        ListAdd.SubItems(2) = Playlist(x).Genre
        ListAdd.SubItems(3) = ConvertTime(Playlist(x).Length, True)
        ListAdd.SubItems(4) = Playlist(x).Filename
        ListAdd.TaG = x
        
        TotalPlaylistLength = TotalPlaylistLength + Playlist(x).Length
        
        If Playlist(x).Length > 0 Then tmp1 = tmp1 + 1
      
      Next x
    
      frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
      frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
      If tmp1 < UBound(Playlist) Then
        frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength, True) & "+"
      Else
        frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength, True)
      End If
      frmPlaylist.lblTotalNum = frmPlaylist.lvwList.ListItems.Count & " files."
      frmPlaylist.Scroll.Value = 1
      If frmPlaylist.lvwList.ListItems.Count > 1 Then
        frmPlaylist.Scroll.Max = frmPlaylist.lvwList.ListItems.Count
      End If
    End If
  End If
  
  If Len(Trim(Command)) > 0 Then
    Select Case Trim(Command)
      Case "-dev"
        Open App.Path & "\device.txt" For Output As #1
        Print #1, "Device output from Simple Amp v" & App.Major & "." & App.Minor & "." & App.Revision & ", using DMC² v1.03."
        Print #1, String(Len("Device output from Simple Amp v" & App.Major & "." & App.Minor & "." & App.Revision & ", using DMC² v1.03."), "-")
        Print #1, vbCr
        If devNum <> -1 Then
          Print #1, "Sound Device: " & devNum & ". " & DMC.Info_DeviceDescrip(devNum)
          Print #1, String(Len("Sound Device: " & devNum & ". " & DMC.Info_DeviceDescrip(devNum)), "-")
        Else
          Print #1, "Sound Device: " & devNum & ". Default / Autodetect"
          Print #1, String(Len("Sound Device: " & devNum & ". Default / Autodetect"), "-")
        End If
        Print #1, "Certified: " & CStr(DMC.Info_DrvIsCertified)
        Print #1, "Hardware Support: " & CStr(DMC.Info_HWSupport)
        Print #1, "Hardware 8-bit support: " & CStr(DMC.Info_HW8bitSupport)
        Print #1, "Hardware 16-bit support: " & CStr(DMC.Info_HW16bitSupport)
        Print #1, "Hardware Mono support: " & CStr(DMC.Info_HWMonoSupport)
        Print #1, "Hardware Stereo support: " & CStr(DMC.Info_HWStereoSupport)
        Print #1, "Sample rates support: " & CStr(DMC.Info_SampRatesSupport)
        Print #1, "Lower sample rate support: " & DMC.Info_SampRatesMin
        Print #1, "Upper sample rate support: " & DMC.Info_SampRatesMax
        Print #1, "Free sample slots: " & DMC.Info_HWSampSlotsFree
        Print #1, "Total Hardware memory: " & DMC.Info_HWMemTotal
        Print #1, "Free Hardware memory: " & DMC.Info_HWMemFree
        Print #1, vbCr
        Print #1, "Written: " & Date & ", " & Time
        Close #1
        MsgBox "Device info for currently selected device written to device.txt", vbInformation, "Device info written"
    End Select
  End If
  
  If frmPlaylist.lvwList.ListItems.Count > 0 Then Play
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.menMain
  Else
    'Get windows work area size to type variable Scr
    SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
    MoveMain = True
    MoveMainOldX = x
    MoveMainOldY = Y
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If MoveMain Then
    If Snap Then  'If Snap window option is on, snap it to screen edges
      'Snap to right or left edge of screen
      If (Left + (x - MoveMainOldX) + Width) / Screen.TwipsPerPixelX > Scr.Right - SnapWidth And (Left + (x - MoveMainOldX) + Width) / Screen.TwipsPerPixelX < Scr.Right + SnapWidth Then
        Left = (Scr.Right * Screen.TwipsPerPixelX) - Width
      ElseIf (Left + (x - MoveMainOldX)) / Screen.TwipsPerPixelX < Scr.Left + SnapWidth And (Left + (x - MoveMainOldX)) / Screen.TwipsPerPixelX > Scr.Left - SnapWidth Then
        Left = (Scr.Left * Screen.TwipsPerPixelX)
      Else
        Left = Left + (x - MoveMainOldX)
      End If
      'Snap to lower or upper edge of screen
      If (Top + (Y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY > Scr.Bottom - SnapWidth And (Top + (Y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY < Scr.Bottom + SnapWidth Then
        Top = (Scr.Bottom * Screen.TwipsPerPixelY) - Height
      ElseIf (Top + (Y - MoveMainOldY)) / Screen.TwipsPerPixelY < Scr.Top + SnapWidth And (Top + (Y - MoveMainOldY)) / Screen.TwipsPerPixelY > Scr.Top - SnapWidth Then
        Top = (Scr.Top * Screen.TwipsPerPixelY)
      Else
        Top = Top + (Y - MoveMainOldY)
      End If
    Else
      Left = Left + (x - MoveMainOldX)
      Top = Top + (Y - MoveMainOldY)
    End If
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  MoveMain = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If UBound(Playlist) > 0 Then
    If frmMenus.SavePlaylist(App.Path & "\current.playlist") Then DoEvents
  End If
  
  If DMC.StreamIsActive Then DMC.CloseStream
  
  'Save settings
  SaveSettings

  'make sure sound & forms is unloaded
  cTray.RemoveFromSysTray
  DMC.TerminateBASS
  Unload frmPitch
  Unload frmAbout
  Unload frmDDE
  Unload frmFind
  Unload frmId3
  Unload frmInfo
  Unload frmMenus
  Unload frmPlaylist
  Unload frmSettings
  Unload frmSkin

End Sub

Private Sub imgClose_Click()
  Unload Me
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgClose_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgClose.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgClose.Height * Screen.TwipsPerPixelY Then
      imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "CloseDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "CloseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "CloseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgDropdown_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.menMain
  Else
    'Get windows work area size to type variable Scr
    SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
    MoveMain = True
    MoveMainOldX = x
    MoveMainOldY = Y
  End If
  Keys.SetFocus
End Sub

Private Sub imgDropdown_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If MoveMain Then
    If Snap Then  'If Snap window option is on, snap it to screen edges
      'Snap to right or left edge of screen
      If (Left + (x - MoveMainOldX) + Width) / Screen.TwipsPerPixelX > Scr.Right - SnapWidth And (Left + (x - MoveMainOldX) + Width) / Screen.TwipsPerPixelX < Scr.Right + SnapWidth Then
        Left = (Scr.Right * Screen.TwipsPerPixelX) - Width
      ElseIf (Left + (x - MoveMainOldX)) / Screen.TwipsPerPixelX < Scr.Left + SnapWidth And (Left + (x - MoveMainOldX)) / Screen.TwipsPerPixelX > Scr.Left - SnapWidth Then
        Left = (Scr.Left * Screen.TwipsPerPixelX)
      Else
        Left = Left + (x - MoveMainOldX)
      End If
      'Snap to lower or upper edge of screen
      If (Top + (Y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY > Scr.Bottom - SnapWidth And (Top + (Y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY < Scr.Bottom + SnapWidth Then
        Top = (Scr.Bottom * Screen.TwipsPerPixelY) - Height
      ElseIf (Top + (Y - MoveMainOldY)) / Screen.TwipsPerPixelY < Scr.Top + SnapWidth And (Top + (Y - MoveMainOldY)) / Screen.TwipsPerPixelY > Scr.Top - SnapWidth Then
        Top = (Scr.Top * Screen.TwipsPerPixelY)
      Else
        Top = Top + (Y - MoveMainOldY)
      End If
    Else
      Left = Left + (x - MoveMainOldX)
      Top = Top + (Y - MoveMainOldY)
    End If
  End If
End Sub

Private Sub imgDropdown_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  MoveMain = False
End Sub

Private Sub imgDropdown_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
  'Adds mp3s dropped in window to playlist
  Dim i As Long, Num As Long
  Dim MP3 As New Mp3Info
  Dim ListAdd
  If Data.GetFormat(vbCFFiles) Then 'If data is list of files
    
    ReDim Playlist(0) As PlaylistData 'Clear playlist
    frmPlaylist.Scroll.Value = 1
    frmPlaylist.Scroll.Max = 2
    frmPlaylist.Scroll.Min = 1
    frmPlaylist.lvwList.ListItems.Clear
    TotalPlaylistLength = 0
    
    For i = 1 To Data.Files.Count 'add each dropped file
      If LCase(Right(Data.Files(i), 4)) = ".mp3" Then
        Num = Num + 1
        ReDim Preserve Playlist(Num) As PlaylistData
        
        'Get ID3v2. If there is none, get ID3v1
        Playlist(Num).Filename = Data.Files(i)
        If Not NoID3 Then
          If ReadID3v2(Data.Files(i)) Then
            Playlist(Num).Album = ID3v2Info.Album
            Playlist(Num).Artist = ID3v2Info.Artist
            Playlist(Num).Genre = TrimGenre(ID3v2Info.Genre)
            Playlist(Num).Title = ID3v2Info.Title
          Else
            If GetID3(Data.Files(i)) Then
              Playlist(Num).Title = Trim(CleanString(ID3v1Info.Title))
              Playlist(Num).Artist = Trim(CleanString(ID3v1Info.Artist))
              Playlist(Num).Album = Trim(CleanString(ID3v1Info.Album))
              Playlist(Num).Genre = DMC.GetGenreDescrip(ID3v1Info.Genre)
            Else
              Playlist(Num).Title = ""
              Playlist(Num).Album = ""
              Playlist(Num).Artist = ""
              Playlist(Num).Genre = ""
            End If
          End If
        End If
        
        'Get mp3 length
        MP3.Filename = Data.Files(i)
        MP3.GetMPEGInfo
        Playlist(Num).Length = MP3.Seconds
        Playlist(Num).Removed = False
        
        TotalPlaylistLength = TotalPlaylistLength + MP3.Seconds
        
        'Add to list
        Set ListAdd = frmPlaylist.lvwList.ListItems.Add
        
        If Len(Playlist(Num).Artist) > 0 And Len(Playlist(Num).Title) > 0 Then
          ListAdd.Text = Playlist(Num).Artist & " - " & Playlist(Num).Title
        Else
          ListAdd.Text = Right(Data.Files(i), Len(Data.Files(i)) - InStrRev(Data.Files(i), "\"))
        End If
        ListAdd.SubItems(1) = Playlist(Num).Album
        ListAdd.SubItems(2) = Playlist(Num).Genre
        ListAdd.SubItems(3) = ConvertTime(Playlist(Num).Length)
        ListAdd.SubItems(4) = Playlist(Num).Filename
        ListAdd.TaG = Num
      
      End If
    Next i
      
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength)
    frmPlaylist.lblTotalNum = frmPlaylist.lvwList.ListItems.Count & " files."
    If frmPlaylist.lvwList.ListItems.Count > 1 Then
      frmPlaylist.Scroll.Max = frmPlaylist.lvwList.ListItems.Count
    End If
    PlayNext
    
  End If
End Sub

Private Sub imgDropdown2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.menMain
  Else
    'Get windows work area size to type variable Scr
    SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
    MoveMain = True
    MoveMainOldX = x
    MoveMainOldY = Y
  End If
  Keys.SetFocus
End Sub

Private Sub imgDropdown2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If MoveMain Then
    If Snap Then  'If Snap window option is on, snap it to screen edges
      'Snap to right or left edge of screen
      If (Left + (x - MoveMainOldX) + Width) / Screen.TwipsPerPixelX > Scr.Right - SnapWidth And (Left + (x - MoveMainOldX) + Width) / Screen.TwipsPerPixelX < Scr.Right + SnapWidth Then
        Left = (Scr.Right * Screen.TwipsPerPixelX) - Width
      ElseIf (Left + (x - MoveMainOldX)) / Screen.TwipsPerPixelX < Scr.Left + SnapWidth And (Left + (x - MoveMainOldX)) / Screen.TwipsPerPixelX > Scr.Left - SnapWidth Then
        Left = (Scr.Left * Screen.TwipsPerPixelX)
      Else
        Left = Left + (x - MoveMainOldX)
      End If
      'Snap to lower or upper edge of screen
      If (Top + (Y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY > Scr.Bottom - SnapWidth And (Top + (Y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY < Scr.Bottom + SnapWidth Then
        Top = (Scr.Bottom * Screen.TwipsPerPixelY) - Height
      ElseIf (Top + (Y - MoveMainOldY)) / Screen.TwipsPerPixelY < Scr.Top + SnapWidth And (Top + (Y - MoveMainOldY)) / Screen.TwipsPerPixelY > Scr.Top - SnapWidth Then
        Top = (Scr.Top * Screen.TwipsPerPixelY)
      Else
        Top = Top + (Y - MoveMainOldY)
      End If
    Else
      Left = Left + (x - MoveMainOldX)
      Top = Top + (Y - MoveMainOldY)
    End If
  End If
End Sub

Private Sub imgDropdown2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  MoveMain = False
End Sub

Private Sub imgDropdown2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub imgMinimize_Click()
  If MinimizeTray Then
    cTray.IconInSysTray
    Hide
    If frmPlaylist.Visible Then
      If OnTop Then AlwaysOnTop frmPlaylist, False
      frmPlaylist.Hide
      frmPlaylist.TaG = "1"
    End If
  Else
    WindowState = vbMinimized
    If frmPlaylist.Visible Then
      frmPlaylist.TaG = "1"
      frmPlaylist.Hide
    End If
  End If
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgMinimize_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgMinimize.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgMinimize.Height * Screen.TwipsPerPixelY Then
      imgMinimize.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "MinimizeDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgMinimize.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "MinimizeUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgMinimize.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "MinimizeUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Public Sub imgNext_Click()
  PlayNext
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgNext_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgNext_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgNext.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgNext.Height * Screen.TwipsPerPixelY Then
      imgNext.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "NextDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgNext.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "NextUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgNext.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "NextUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgPlaylist_Click()
  If frmPlaylist.Visible Then
    If OnTop Then AlwaysOnTop frmPlaylist, False
    frmPlaylist.Hide
    imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
  Else
    If OnTop Then AlwaysOnTop frmPlaylist, True
    frmPlaylist.Show
    imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
  End If
  If OnTop Then AlwaysOnTop Me, True
End Sub

Public Sub LoadSkin(Skin As String)
  'Loads skin 'Skin' by reading from app.path & "\" & skin & ".ini"
  Dim Where As String 'Stores where in the ini-file the program is, to return an nice error
  Dim ErrorText As String
  On Error GoTo SkinError
  
  
  'First, set everythings size & position, start with size
  Where = "[Dimensions]"
  'Main window
  Height = val(ReadINI("Dimensions", "MainHeight", App.Path & "\skins\" & Skin & ".ini")) * Screen.TwipsPerPixelY
  Width = val(ReadINI("Dimensions", "MainWidth", App.Path & "\skins\" & Skin & ".ini")) * Screen.TwipsPerPixelX
  imgDropdown2.Height = Height
  imgDropdown2.Width = Width
  imgClose.Height = val(ReadINI("Dimensions", "CloseHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgClose.Width = val(ReadINI("Dimensions", "CloseWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgMinimize.Height = val(ReadINI("Dimensions", "MinimizeHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgMinimize.Width = val(ReadINI("Dimensions", "MinimizeWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgPrev.Height = val(ReadINI("Dimensions", "PreviousHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgPrev.Width = val(ReadINI("Dimensions", "PreviousWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgPlayPause.Height = val(ReadINI("Dimensions", "PlayPauseHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgPlayPause.Width = val(ReadINI("Dimensions", "PlayPauseWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgStop.Height = val(ReadINI("Dimensions", "StopHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgStop.Width = val(ReadINI("Dimensions", "StopWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgNext.Height = val(ReadINI("Dimensions", "NextHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgNext.Width = val(ReadINI("Dimensions", "NextWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgPlaylist.Height = val(ReadINI("Dimensions", "PlaylistHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgPlaylist.Width = val(ReadINI("Dimensions", "PlaylistWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgRepeat.Height = val(ReadINI("Dimensions", "RepeatHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgRepeat.Width = val(ReadINI("Dimensions", "RepeatWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgShuffle.Height = val(ReadINI("Dimensions", "ShuffleHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgShuffle.Width = val(ReadINI("Dimensions", "ShuffleWidth", App.Path & "\skins\" & Skin & ".ini"))
  imgStereoMono.Height = val(ReadINI("Dimensions", "StereoMonoHeight", App.Path & "\skins\" & Skin & ".ini"))
  imgStereoMono.Width = val(ReadINI("Dimensions", "StereoMonoWidth", App.Path & "\skins\" & Skin & ".ini"))
  PicSpectrum1.Height = val(ReadINI("Dimensions", "SpectrumHeight", App.Path & "\skins\" & Skin & ".ini"))
  PicSpectrum1.Width = val(ReadINI("Dimensions", "SpectrumWidth", App.Path & "\skins\" & Skin & ".ini"))
  Position.Height = val(ReadINI("Dimensions", "PositionBarHeight", App.Path & "\skins\" & Skin & ".ini"))
  Position.Width = val(ReadINI("Dimensions", "PositionBarWidth", App.Path & "\skins\" & Skin & ".ini"))
  Volume.Height = val(ReadINI("Dimensions", "VolumeBarHeight", App.Path & "\skins\" & Skin & ".ini"))
  Volume.Width = val(ReadINI("Dimensions", "VolumeBarWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.Height = val(ReadINI("Dimensions", "TextArtistTitleHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.Width = val(ReadINI("Dimensions", "TextArtistTitleWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.Height = val(ReadINI("Dimensions", "TextYearHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.Width = val(ReadINI("Dimensions", "TextYearWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.Height = val(ReadINI("Dimensions", "TextAlbumHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.Width = val(ReadINI("Dimensions", "TextAlbumWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.Height = val(ReadINI("Dimensions", "TextGenreHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.Width = val(ReadINI("Dimensions", "TextGenreWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.Height = val(ReadINI("Dimensions", "TextCommentsHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.Width = val(ReadINI("Dimensions", "TextCommentsWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.Height = val(ReadINI("Dimensions", "TextInfoHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.Width = val(ReadINI("Dimensions", "TextInfoWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.Height = val(ReadINI("Dimensions", "TextTimeHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.Width = val(ReadINI("Dimensions", "TextTimeWidth", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.Height = val(ReadINI("Dimensions", "TextTotalTimeHeight", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.Width = val(ReadINI("Dimensions", "TextTotalTimeWidth", App.Path & "\skins\" & Skin & ".ini"))
  'Playlist window
  frmPlaylist.Height = val(ReadINI("Dimensions", "PlaylistMainHeight", App.Path & "\skins\" & Skin & ".ini")) * Screen.TwipsPerPixelY
  frmPlaylist.Width = val(ReadINI("Dimensions", "PlaylistMainWidth", App.Path & "\skins\" & Skin & ".ini")) * Screen.TwipsPerPixelX
  frmPlaylist.imgDropdown.Height = frmPlaylist.Height
  frmPlaylist.imgDropdown.Width = frmPlaylist.Width
  frmPlaylist.imgClose.Height = val(ReadINI("Dimensions", "PlaylistCloseHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgClose.Width = val(ReadINI("Dimensions", "PlaylistCloseWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgAdd.Height = val(ReadINI("Dimensions", "PlaylistAddHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgAdd.Width = val(ReadINI("Dimensions", "PlaylistAddWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgRem.Height = val(ReadINI("Dimensions", "PlaylistRemoveHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgRem.Width = val(ReadINI("Dimensions", "PlaylistRemoveWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSelect.Height = val(ReadINI("Dimensions", "PlaylistSelectHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSelect.Width = val(ReadINI("Dimensions", "PlaylistSelectWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgList.Height = val(ReadINI("Dimensions", "PlaylistbListHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgList.Width = val(ReadINI("Dimensions", "PlaylistbListWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSubdirs.Height = val(ReadINI("Dimensions", "PlaylistSubdirsHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSubdirs.Width = val(ReadINI("Dimensions", "PlaylistSubdirsWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Height = val(ReadINI("Dimensions", "PlaylistColumnsHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Width = val(ReadINI("Dimensions", "PlaylistColumnsWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Height = val(ReadINI("Dimensions", "PlaylistColumnsHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Width = val(ReadINI("Dimensions", "PlaylistColumnsWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.Height = val(ReadINI("Dimensions", "PlaylistScrollHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.Width = val(ReadINI("Dimensions", "PlaylistScrollWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.Height = val(ReadINI("Dimensions", "PlaylistListHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.Width = val(ReadINI("Dimensions", "PlaylistListWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.Height = val(ReadINI("Dimensions", "PlaylistTextNumHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.Width = val(ReadINI("Dimensions", "PlaylistTextNumWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.Height = val(ReadINI("Dimensions", "PlaylistTextTimeHeight", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.Width = val(ReadINI("Dimensions", "PlaylistTextTimeWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.ColumnHeaders(1).Width = val(ReadINI("Dimensions", "ColumnArtistTitleWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.ColumnHeaders(2).Width = val(ReadINI("Dimensions", "ColumnAlbumWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.ColumnHeaders(3).Width = val(ReadINI("Dimensions", "ColumnGenreWidth", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.ColumnHeaders(4).Width = val(ReadINI("Dimensions", "ColumnTimeWidth", App.Path & "\skins\" & Skin & ".ini"))
  
  'Get Location
  Where = "[Locations]"
  'Main window
  imgClose.Left = val(ReadINI("Locations", "CloseX", App.Path & "\skins\" & Skin & ".ini"))
  imgClose.Top = val(ReadINI("Locations", "CloseY", App.Path & "\skins\" & Skin & ".ini"))
  imgMinimize.Left = val(ReadINI("Locations", "MinimizeX", App.Path & "\skins\" & Skin & ".ini"))
  imgMinimize.Top = val(ReadINI("Locations", "MinimizeY", App.Path & "\skins\" & Skin & ".ini"))
  imgPrev.Left = val(ReadINI("Locations", "PreviousX", App.Path & "\skins\" & Skin & ".ini"))
  imgPrev.Top = val(ReadINI("Locations", "PreviousY", App.Path & "\skins\" & Skin & ".ini"))
  imgPlayPause.Left = val(ReadINI("Locations", "PlayPauseX", App.Path & "\skins\" & Skin & ".ini"))
  imgPlayPause.Top = val(ReadINI("Locations", "PlayPauseY", App.Path & "\skins\" & Skin & ".ini"))
  imgStop.Left = val(ReadINI("Locations", "StopX", App.Path & "\skins\" & Skin & ".ini"))
  imgStop.Top = val(ReadINI("Locations", "StopY", App.Path & "\skins\" & Skin & ".ini"))
  imgNext.Left = val(ReadINI("Locations", "NextX", App.Path & "\skins\" & Skin & ".ini"))
  imgNext.Top = val(ReadINI("Locations", "NextY", App.Path & "\skins\" & Skin & ".ini"))
  imgPlaylist.Left = val(ReadINI("Locations", "PlaylistX", App.Path & "\skins\" & Skin & ".ini"))
  imgPlaylist.Top = val(ReadINI("Locations", "PlaylistY", App.Path & "\skins\" & Skin & ".ini"))
  imgRepeat.Left = val(ReadINI("Locations", "RepeatX", App.Path & "\skins\" & Skin & ".ini"))
  imgRepeat.Top = val(ReadINI("Locations", "RepeatY", App.Path & "\skins\" & Skin & ".ini"))
  imgShuffle.Left = val(ReadINI("Locations", "ShuffleX", App.Path & "\skins\" & Skin & ".ini"))
  imgShuffle.Top = val(ReadINI("Locations", "ShuffleY", App.Path & "\skins\" & Skin & ".ini"))
  imgStereoMono.Left = val(ReadINI("Locations", "StereoMonoX", App.Path & "\skins\" & Skin & ".ini"))
  imgStereoMono.Top = val(ReadINI("Locations", "StereoMonoY", App.Path & "\skins\" & Skin & ".ini"))
  PicSpectrum1.Left = val(ReadINI("Locations", "SpectrumX", App.Path & "\skins\" & Skin & ".ini"))
  PicSpectrum1.Top = val(ReadINI("Locations", "SpectrumY", App.Path & "\skins\" & Skin & ".ini"))
  Position.Left = val(ReadINI("Locations", "PositionBarX", App.Path & "\skins\" & Skin & ".ini"))
  Position.Top = val(ReadINI("Locations", "PositionBarY", App.Path & "\skins\" & Skin & ".ini"))
  Volume.Left = val(ReadINI("Locations", "VolumeBarX", App.Path & "\skins\" & Skin & ".ini"))
  Volume.Top = val(ReadINI("Locations", "VolumeBarY", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.Left = val(ReadINI("Locations", "TextArtistTitleX", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.Top = val(ReadINI("Locations", "TextArtistTitleY", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.Left = val(ReadINI("Locations", "TextYearX", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.Top = val(ReadINI("Locations", "TextYearY", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.Left = val(ReadINI("Locations", "TextAlbumX", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.Top = val(ReadINI("Locations", "TextAlbumY", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.Left = val(ReadINI("Locations", "TextGenreX", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.Top = val(ReadINI("Locations", "TextGenreY", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.Left = val(ReadINI("Locations", "TextCommentsX", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.Top = val(ReadINI("Locations", "TextCommentsY", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.Left = val(ReadINI("Locations", "TextInfoX", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.Top = val(ReadINI("Locations", "TextInfoY", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.Left = val(ReadINI("Locations", "TextTimeX", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.Top = val(ReadINI("Locations", "TextTimeY", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.Left = val(ReadINI("Locations", "TextTotalTimeX", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.Top = val(ReadINI("Locations", "TextTotalTimeY", App.Path & "\skins\" & Skin & ".ini"))
  'Playlist window
  frmPlaylist.imgClose.Left = val(ReadINI("Locations", "PlaylistCloseX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgClose.Top = val(ReadINI("Locations", "PlaylistCloseY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgAdd.Left = val(ReadINI("Locations", "PlaylistAddX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgAdd.Top = val(ReadINI("Locations", "PlaylistAddY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgRem.Left = val(ReadINI("Locations", "PlaylistRemoveX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgRem.Top = val(ReadINI("Locations", "PlaylistRemoveY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSelect.Left = val(ReadINI("Locations", "PlaylistSelectX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSelect.Top = val(ReadINI("Locations", "PlaylistSelectY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgList.Left = val(ReadINI("Locations", "PlaylistbListX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgList.Top = val(ReadINI("Locations", "PlaylistbListY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSubdirs.Left = val(ReadINI("Locations", "PlaylistSubdirsX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSubdirs.Top = val(ReadINI("Locations", "PlaylistSubdirsY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Left = val(ReadINI("Locations", "PlaylistColumnsX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Top = val(ReadINI("Locations", "PlaylistColumnsY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Left = val(ReadINI("Locations", "PlaylistColumnsX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Top = val(ReadINI("Locations", "PlaylistColumnsY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.Left = val(ReadINI("Locations", "PlaylistScrollX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.Top = val(ReadINI("Locations", "PlaylistScrollY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.Left = val(ReadINI("Locations", "PlaylistListX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.Top = val(ReadINI("Locations", "PlaylistListY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.Left = val(ReadINI("Locations", "PlaylistTextNumX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.Top = val(ReadINI("Locations", "PlaylistTextNumY", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.Left = val(ReadINI("Locations", "PlaylistTextTimeX", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.Top = val(ReadINI("Locations", "PlaylistTextTimeY", App.Path & "\skins\" & Skin & ".ini"))
  
  'Setup labels
  Where = "[Text]"
  lblArtistTitle.Font = ReadINI("Text", "ArtistTitleFont", App.Path & "\skins\" & Skin & ".ini")
  lblArtistTitle.FontSize = val(ReadINI("Text", "ArtistTitleSize", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.FontBold = CBool(ReadINI("Text", "ArtistTitleBold", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.FontItalic = CBool(ReadINI("Text", "ArtistTitleItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblArtistTitle.Alignment = val(ReadINI("Text", "ArtistTitleAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.Font = ReadINI("Text", "YearFont", App.Path & "\skins\" & Skin & ".ini")
  lblYear.FontSize = val(ReadINI("Text", "YearSize", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.FontBold = CBool(ReadINI("Text", "YearBold", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.FontItalic = CBool(ReadINI("Text", "YearItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblYear.Alignment = val(ReadINI("Text", "YearAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.Font = ReadINI("Text", "AlbumFont", App.Path & "\skins\" & Skin & ".ini")
  lblAlbum.FontSize = val(ReadINI("Text", "AlbumSize", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.FontBold = CBool(ReadINI("Text", "AlbumBold", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.FontItalic = CBool(ReadINI("Text", "AlbumItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblAlbum.Alignment = val(ReadINI("Text", "AlbumAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.Font = ReadINI("Text", "GenreFont", App.Path & "\skins\" & Skin & ".ini")
  lblGenre.FontSize = val(ReadINI("Text", "GenreSize", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.FontBold = CBool(ReadINI("Text", "GenreBold", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.FontItalic = CBool(ReadINI("Text", "GenreItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblGenre.Alignment = val(ReadINI("Text", "GenreAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.Font = ReadINI("Text", "CommentsFont", App.Path & "\skins\" & Skin & ".ini")
  lblComments.FontSize = val(ReadINI("Text", "CommentsSize", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.FontBold = CBool(ReadINI("Text", "CommentsBold", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.FontItalic = CBool(ReadINI("Text", "CommentsItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblComments.Alignment = val(ReadINI("Text", "CommentsAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.Font = ReadINI("Text", "InfoFont", App.Path & "\skins\" & Skin & ".ini")
  lblInfo.FontSize = val(ReadINI("Text", "InfoSize", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.FontBold = CBool(ReadINI("Text", "InfoBold", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.FontItalic = CBool(ReadINI("Text", "InfoItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblInfo.Alignment = val(ReadINI("Text", "InfoAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.Font = ReadINI("Text", "TimeFont", App.Path & "\skins\" & Skin & ".ini")
  lblTime.FontSize = val(ReadINI("Text", "TimeSize", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.FontBold = CBool(ReadINI("Text", "TimeBold", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.FontItalic = CBool(ReadINI("Text", "TimeItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblTime.Alignment = val(ReadINI("Text", "TimeAlignment", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.Font = ReadINI("Text", "TotalTimeFont", App.Path & "\skins\" & Skin & ".ini")
  lblTotalTime.FontSize = val(ReadINI("Text", "TotalTimeSize", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.FontBold = CBool(ReadINI("Text", "TotalTimeBold", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.FontItalic = CBool(ReadINI("Text", "TotalTimeItalic", App.Path & "\skins\" & Skin & ".ini"))
  lblTotalTime.Alignment = val(ReadINI("Text", "TotalTimeAlignment", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.Font = ReadINI("Text", "PlaylistNumFont", App.Path & "\skins\" & Skin & ".ini")
  frmPlaylist.lblTotalNum.FontSize = val(ReadINI("Text", "PlaylistNumSize", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.FontBold = CBool(ReadINI("Text", "PlaylistNumBold", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.FontItalic = CBool(ReadINI("Text", "PlaylistNumItalic", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalNum.Alignment = val(ReadINI("Text", "PlaylistNumAlignment", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.Font = ReadINI("Text", "PlaylistTimeFont", App.Path & "\skins\" & Skin & ".ini")
  frmPlaylist.lblTotalTime.FontSize = val(ReadINI("Text", "PlaylistTimeSize", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.FontBold = CBool(ReadINI("Text", "PlaylistTimeBold", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.FontItalic = CBool(ReadINI("Text", "PlaylistTimeItalic", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lblTotalTime.Alignment = val(ReadINI("Text", "PlaylistTimeAlignment", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.Font = ReadINI("Text", "ListFont", App.Path & "\skins\" & Skin & ".ini")
  
  'Set images
  Where = "[Images]"
  Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "MainBackground", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistBackground", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgColumns.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Columns", App.Path & "\skins\" & Skin & ".ini"))
    
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PositionBar", App.Path & "\skins\" & Skin & ".ini"))
  Position.PaintPicture1 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  Position.PaintPicture1 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PositionAfter", App.Path & "\skins\" & Skin & ".ini"))
  Position.PaintPicture2 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  Position.PaintPicture2 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PositionBefore", App.Path & "\skins\" & Skin & ".ini"))
  Position.PaintPicture3 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  Position.PaintPicture3 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
    
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "VolumeBar", App.Path & "\skins\" & Skin & ".ini"))
  Volume.PaintPicture1 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  Volume.PaintPicture1 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "VolumeAfter", App.Path & "\skins\" & Skin & ".ini"))
  Volume.PaintPicture2 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  Volume.PaintPicture2 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "VolumeBefore", App.Path & "\skins\" & Skin & ".ini"))
  Volume.PaintPicture3 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  Volume.PaintPicture3 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ScrollBar", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.PaintPicture1 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  frmPlaylist.Scroll.PaintPicture1 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ScrollAfter", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.PaintPicture2 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  frmPlaylist.Scroll.PaintPicture2 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  PictureLoader.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ScrollBefore", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.Scroll.PaintPicture3 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  frmPlaylist.Scroll.PaintPicture3 PictureLoader.Picture, 0, 0, , , 0, 0, PictureLoader.ScaleWidth, PictureLoader.ScaleHeight
  
  If Not FileExists(App.Path & "\skins\" & ReadINI("Images", "Spectrum1Background", App.Path & "\skins\" & Skin & ".ini")) Then Err.Raise 0
  If Not FileExists(App.Path & "\skins\" & ReadINI("Images", "Spectrum2Background", App.Path & "\skins\" & Skin & ".ini")) Then Err.Raise 0
  If Not FileExists(App.Path & "\skins\" & ReadINI("Images", "Spectrum3Background", App.Path & "\skins\" & Skin & ".ini")) Then Err.Raise 0
  If Not FileExists(App.Path & "\skins\" & ReadINI("Images", "Spectrum4Background", App.Path & "\skins\" & Skin & ".ini")) Then Err.Raise 0
  If Not FileExists(App.Path & "\skins\" & ReadINI("Images", "Spectrum5Background", App.Path & "\skins\" & Skin & ".ini")) Then Err.Raise 0
  If Not FileExists(App.Path & "\skins\" & ReadINI("Images", "SpectrumBackgroundOff", App.Path & "\skins\" & Skin & ".ini")) Then Err.Raise 0
  
  If Spectrum = 1 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum1Background", App.Path & "\skins\" & Skin & ".ini"))
  ElseIf Spectrum = 2 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum2Background", App.Path & "\skins\" & Skin & ".ini"))
  ElseIf Spectrum = 3 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum3Background", App.Path & "\skins\" & Skin & ".ini"))
  ElseIf Spectrum = 4 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum4Background", App.Path & "\skins\" & Skin & ".ini"))
  ElseIf Spectrum = 5 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum5Background", App.Path & "\skins\" & Skin & ".ini"))
  ElseIf Spectrum = 6 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SpectrumBackgroundOff", App.Path & "\skins\" & CurrentSkin & ".ini"))
  End If
  SpectrumBar.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum3Bar", App.Path & "\skins\" & Skin & ".ini"))
  
  frmPlaylist.imgAdd.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "AddUp", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgRem.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RemoveUp", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgSelect.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SelectUp", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.imgList.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ListUp", App.Path & "\skins\" & Skin & ".ini"))
  imgPrev.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PreviousUp", App.Path & "\skins\" & Skin & ".ini"))
  imgNext.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "NextUp", App.Path & "\skins\" & Skin & ".ini"))
  If DMC.StreamIsPaused Then
    imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayUp", App.Path & "\skins\" & Skin & ".ini"))
  Else
    imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseUp", App.Path & "\skins\" & Skin & ".ini"))
  End If
  imgStop.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "StopUp", App.Path & "\skins\" & Skin & ".ini"))
  If frmPlaylist.Visible Or frmPlaylist.TaG = "1" Then
    imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistOn", App.Path & "\skins\" & Skin & ".ini"))
  Else
    imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistUp", App.Path & "\skins\" & Skin & ".ini"))
  End If
  frmPlaylist.imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistCloseUp", App.Path & "\skins\" & Skin & ".ini"))
  imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "CloseUp", App.Path & "\skins\" & Skin & ".ini"))
  imgMinimize.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "MinimizeUp", App.Path & "\skins\" & Skin & ".ini"))
  
  If Repeat Then
    imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOn", App.Path & "\skins\" & Skin & ".ini"))
  Else
    imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOffUp", App.Path & "\skins\" & Skin & ".ini"))
  End If
  If Shuffle Then
    imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOn", App.Path & "\skins\" & Skin & ".ini"))
  Else
    imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOffUp", App.Path & "\skins\" & Skin & ".ini"))
  End If
  If SearchInSubdirs Then
    frmPlaylist.imgSubdirs.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SubdirsOn", App.Path & "\skins\" & Skin & ".ini"))
  Else
    frmPlaylist.imgSubdirs.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SubdirsOff", App.Path & "\skins\" & Skin & ".ini"))
  End If
  
  If DMC.StreamIsMono Then
    imgStereoMono.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Mono", App.Path & "\skins\" & Skin & ".ini"))
  Else
    imgStereoMono.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Stereo", App.Path & "\skins\" & Skin & ".ini"))
  End If
  
  'Get colors
  Where = "[Colors]"
  Color1L = Hex2VB(ReadINI("Colors", "Spectrum1Left", App.Path & "\skins\" & Skin & ".ini"))
  Color1R = Hex2VB(ReadINI("Colors", "Spectrum1Right", App.Path & "\skins\" & Skin & ".ini"))
  Color2 = Hex2VB(ReadINI("Colors", "Spectrum2", App.Path & "\skins\" & Skin & ".ini"))
  Color4 = Hex2VB(ReadINI("Colors", "Spectrum4", App.Path & "\skins\" & Skin & ".ini"))
  Color5 = Hex2VB(ReadINI("Colors", "Spectrum5", App.Path & "\skins\" & Skin & ".ini"))
  frmPlaylist.lvwList.BackColor = Hex2VB(ReadINI("Colors", "PlaylistBackground", App.Path & "\skins\" & Skin & ".ini"))
  
  If DMC.StreamIsActive Then
    lblArtistTitle.ForeColor = Hex2VB(ReadINI("Colors", "ArtistEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblAlbum.ForeColor = Hex2VB(ReadINI("Colors", "AlbumEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblYear.ForeColor = Hex2VB(ReadINI("Colors", "YearEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblComments.ForeColor = Hex2VB(ReadINI("Colors", "CommentsEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblGenre.ForeColor = Hex2VB(ReadINI("Colors", "GenreEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblInfo.ForeColor = Hex2VB(ReadINI("Colors", "InfoEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblTime.ForeColor = Hex2VB(ReadINI("Colors", "TimeEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "TotalTimeEnabledText", App.Path & "\skins\" & Skin & ".ini"))
  Else
    lblArtistTitle.ForeColor = Hex2VB(ReadINI("Colors", "ArtistDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblAlbum.ForeColor = Hex2VB(ReadINI("Colors", "AlbumDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblYear.ForeColor = Hex2VB(ReadINI("Colors", "YearDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblComments.ForeColor = Hex2VB(ReadINI("Colors", "CommentsDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblGenre.ForeColor = Hex2VB(ReadINI("Colors", "GenreDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblInfo.ForeColor = Hex2VB(ReadINI("Colors", "InfoDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblTime.ForeColor = Hex2VB(ReadINI("Colors", "TimeDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "TotalTimeDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumDisabledText", App.Path & "\skins\" & Skin & ".ini"))
  End If
  If frmPlaylist.lvwList.ListItems.Count > 0 Then
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & Skin & ".ini"))
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & Skin & ".ini"))
  Else
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeDisabledText", App.Path & "\skins\" & Skin & ".ini"))
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumDisabledText", App.Path & "\skins\" & Skin & ".ini"))
  End If
  
  frmPlaylist.lvwList.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistText", App.Path & "\skins\" & Skin & ".ini"))
  Exit Sub
  
SkinError:
  ErrorText = "There was an error when loading the skin '" & Skin & "', at sector " & Where & ". "
  Select Case Where
  Case "[Images]": ErrorText = ErrorText & "There could be images missing or the locations could be wrong. "
  Case "[Locations]", "[Dimensions]": ErrorText = ErrorText & "There could be an value missing or maybe the value isn't an number. "
  Case "[Text]": ErrorText = ErrorText & "The skins font could be missing or your system or any of the font settings could be unavailabe for the font. "
  Case "[Colors]": ErrorText = ErrorText & "There could be an missing value or one of the values is in the wrong format (Hex RGB, ex: FFAA00). "
  End Select
  ErrorText = ErrorText & "There could be spelling mistakes in the ini-file. Correct the error and try again." & vbNewLine & vbNewLine
  If Skin <> "Steel" Then
    MsgBox ErrorText & "The Default skin, 'Steel Amp' will be loaded.", vbCritical, "Skin Error"
    CurrentSkin = "Steel"
    LoadSkin CurrentSkin
  Else
    MsgBox ErrorText, vbCritical, "Skin Error"
    Unload Me
  End If
End Sub

Public Sub LoadSettings()
  'This sub loads program setting from ini-file
  Dim x As Long
  
  On Error GoTo LoadErr
    
    If Not FileExists(App.Path & "\settings.ini") Then GoTo LoadErr
  
    CurrentSkin = ReadINI("Simple Amp", "Skin", App.Path & "\settings.ini")
    If CurrentSkin = "" Then CurrentSkin = "Steel"
    Repeat = CBool(ReadINI("Simple Amp", "Repeat", App.Path & "\settings.ini"))
    Shuffle = CBool(ReadINI("Simple Amp", "Shuffle", App.Path & "\settings.ini"))
    SearchInSubdirs = CBool(ReadINI("Simple Amp", "Subdirs", App.Path & "\settings.ini"))
    CurrentVolume = val(ReadINI("Simple Amp", "Volume", App.Path & "\settings.ini"))
    Spectrum = val(ReadINI("Simple Amp", "Spectrum", App.Path & "\settings.ini"))
    Top = val(ReadINI("Simple Amp", "Windowx", App.Path & "\settings.ini"))
    Left = val(ReadINI("Simple Amp", "WindowY", App.Path & "\settings.ini"))
    frmPlaylist.Top = val(ReadINI("Simple Amp", "Playlistx", App.Path & "\settings.ini"))
    frmPlaylist.Left = val(ReadINI("Simple Amp", "PlaylistY", App.Path & "\settings.ini"))
    AlwaysTray = CBool(ReadINI("Simple Amp", "AlwaysTray", App.Path & "\settings.ini"))
    MinimizeTray = CBool(ReadINI("Simple Amp", "MinimizeTray", App.Path & "\settings.ini"))
    OnTop = CBool(ReadINI("Simple Amp", "OnTop", App.Path & "\settings.ini"))
    TrayIcon = val(ReadINI("Simple Amp", "TrayIcon", App.Path & "\settings.ini"))
    NoID3 = CBool(ReadINI("Simple Amp", "NoID3", App.Path & "\settings.ini"))
    StartInTray = CBool(ReadINI("Simple Amp", "StartInTray", App.Path & "\settings.ini"))
    Snap = CBool(ReadINI("Simple Amp", "Snap", App.Path & "\settings.ini"))
    CurrentlyPlaying = val(ReadINI("Simple Amp", "Playing", App.Path & "\settings.ini"))
    NumStarted = val(ReadINI("Simple Amp", "TotalStart", App.Path & "\settings.ini"))
    TimePlayed = val(ReadINI("Simple Amp", "TotalTime", App.Path & "\settings.ini"))
    If CBool(ReadINI("Simple Amp", "PlaylistOn", App.Path & "\settings.ini")) Then
      If StartInTray Then
        frmPlaylist.TaG = "1"
      ElseIf Not StartInTray And OnTop Then
        AlwaysOnTop frmPlaylist, True
      Else
        frmPlaylist.Show
      End If
    End If
    For x = 1 To 5
      Hotkey(x).Inactive = val(ReadINI("Simple Amp", "Hotkey" & x & "Inactive", App.Path & "\settings.ini"))
      Hotkey(x).Active = val(ReadINI("Simple Amp", "Hotkey" & x & "Active", App.Path & "\settings.ini"))
      Hotkey(x).Program = ReadINI("Simple Amp", "Hotkey" & x & "Location", App.Path & "\settings.ini")
    Next x
    devNum = val(ReadINI("Device", "DeviceNum", App.Path & "\settings.ini"))
    dev44100 = CBool(ReadINI("Device", "Mixing44100", App.Path & "\settings.ini"))
    devStereo = CBool(ReadINI("Device", "MixingStereo", App.Path & "\settings.ini"))
    dev16bits = CBool(ReadINI("Device", "Mixing16bits", App.Path & "\settings.ini"))
    devBuffer = CSng(ReadINI("Device", "Buffer", App.Path & "\settings.ini"))
    devPanning = val(ReadINI("Device", "Panning", App.Path & "\settings.ini"))
    If devBuffer > 2 Or devBuffer < 0.3 Then
      devBuffer = 1
    End If
  
  Exit Sub

LoadErr:
  MsgBox "There was an error when trying to load settings from " & App.Path & "\settings.ini. Default will be used.", vbExclamation, "Load Error"
  devNum = -1
  dev44100 = True
  devStereo = True
  dev16bits = True
  devBuffer = 1
  CurrentSkin = "Steel"
  Repeat = True
  CurrentVolume = 100
  Spectrum = 1
  MinimizeTray = True
  TrayIcon = 1
End Sub

Public Sub SaveSettings()
  'This sub saves program settings to ini-file
  Dim x As Long
  
  TimePlayed = TimePlayed + Int(TimePlayedNow / 60)
  
  'Program settings
  WriteINI "Simple Amp", "TotalStart", CStr(NumStarted), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "TotalTime", CStr(TimePlayed), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Skin", CurrentSkin, App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Repeat", CStr(Repeat), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Shuffle", CStr(Shuffle), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Volume", CStr(CurrentVolume), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Subdirs", CStr(SearchInSubdirs), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Windowx", CStr(Top), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "WindowY", CStr(Left), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Playlistx", CStr(frmPlaylist.Top), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "PlaylistY", CStr(frmPlaylist.Left), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Spectrum", CStr(Spectrum), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "AlwaysTray", CStr(AlwaysTray), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "MinimizeTray", CStr(MinimizeTray), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "TrayIcon", CStr(TrayIcon), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "OnTop", CStr(OnTop), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "NoID3", CStr(NoID3), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "StartInTray", CStr(StartInTray), App.Path & "\settings.ini"
  WriteINI "Simple Amp", "Snap", CStr(Snap), App.Path & "\settings.ini"
  If frmPlaylist.TaG = "1" Then
    WriteINI "Simple Amp", "PlaylistOn", "True", App.Path & "\settings.ini"
  Else
    WriteINI "Simple Amp", "PlaylistOn", CStr(frmPlaylist.Visible), App.Path & "\settings.ini"
  End If
  For x = 1 To 5
    WriteINI "Simple Amp", "Hotkey" & x & "Inactive", CStr(Hotkey(x).Inactive), App.Path & "\settings.ini"
    WriteINI "Simple Amp", "Hotkey" & x & "Active", CStr(Hotkey(x).Active), App.Path & "\settings.ini"
    If Len(Hotkey(x).Program) > 0 Then WriteINI "Simple Amp", "Hotkey" & x & "Location", Hotkey(x).Program, App.Path & "\settings.ini"
  Next x
  WriteINI "Simple Amp", "Playing", CStr(CurrentlyPlaying), App.Path & "\settings.ini"
  'Sound device settings
  WriteINI "Device", "DeviceNum", CStr(devNum), App.Path & "\settings.ini"
  WriteINI "Device", "Mixing44100", CStr(dev44100), App.Path & "\settings.ini"
  WriteINI "Device", "MixingStereo", CStr(devStereo), App.Path & "\settings.ini"
  WriteINI "Device", "Mixing16bits", CStr(dev16bits), App.Path & "\settings.ini"
  WriteINI "Device", "Buffer", CStr(devBuffer), App.Path & "\settings.ini"
  WriteINI "Device", "Panning", CStr(devPanning), App.Path & "\settings.ini"
 
End Sub

Private Sub imgPlaylist_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgPlaylist_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgPlaylist_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgPlaylist.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgPlaylist.Height * Screen.TwipsPerPixelY Then
      imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      If frmPlaylist.Visible Then
        imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
      Else
        imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
      End If
    End If
  End If
End Sub

Public Sub imgPlayPause_Click()
  If CurrentlyPlaying > 0 And frmPlaylist.lvwList.ListItems.Count >= CurrentlyPlaying Then
    If DMC.StreamIsPaused Then
      If DMC.StreamPos <> -1 Then
        imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
        DMC.ResumeStream
      Else
        Play
        imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
      End If
    ElseIf DMC.StreamIsActive Then
      imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
      DMC.PauseStream
    Else
      Play
      imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  Else
    PlayNext
  End If
End Sub

Private Sub imgPlayPause_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgPlayPause_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgPlayPause_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgPlayPause.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgPlayPause.Height * Screen.TwipsPerPixelY Then
      If DMC.StreamIsActive Then
        If DMC.StreamIsPaused Then
          imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
        Else
          imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
        End If
      Else
        imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
      End If
    Else
      If DMC.StreamIsActive Then
        If DMC.StreamIsPaused Then
          imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
        Else
          imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
        End If
      Else
        imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
      End If
    End If
  End If
End Sub

Public Sub imgPrev_Click()
  PlayPrev
End Sub

Private Sub imgPrev_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgPrev_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgPrev_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgPrev.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgPrev.Height * Screen.TwipsPerPixelY Then
      imgPrev.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PreviousDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgPrev.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PreviousUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgPrev_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgPrev.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PreviousUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Public Sub imgRepeat_Click()
  If Repeat Then
    Repeat = False
    imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOffUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
  Else
    Repeat = True
    imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
  End If
End Sub

Private Sub imgRepeat_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgRepeat_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgRepeat_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgRepeat.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgRepeat.Height * Screen.TwipsPerPixelY Then
      imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOffDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      If Repeat Then
        imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
      Else
        imgRepeat.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RepeatOffUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
      End If
    End If
  End If
End Sub

Public Sub imgShuffle_Click()
  If Shuffle Then
    Shuffle = False
    imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOffUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
  Else
    Shuffle = True
    imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
  End If
End Sub

Private Sub imgShuffle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgShuffle_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgShuffle_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgShuffle.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgShuffle.Height * Screen.TwipsPerPixelY Then
      imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOffDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      If Shuffle Then
        imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
      Else
        imgShuffle.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ShuffleOffUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
      End If
    End If
  End If
End Sub

Public Sub imgStop_Click()
  'If DMC.StreamIsActive Or DMC.StreamIsPaused Then
    DMC.CloseStream
    lblTime = "00:00"
    PicSpectrum1.Cls
    Position.Value = Position.Min
    imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlayUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblArtistTitle.ForeColor = Hex2VB(ReadINI("Colors", "ArtistDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblAlbum.ForeColor = Hex2VB(ReadINI("Colors", "AlbumDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblYear.ForeColor = Hex2VB(ReadINI("Colors", "YearDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblComments.ForeColor = Hex2VB(ReadINI("Colors", "CommentsDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblGenre.ForeColor = Hex2VB(ReadINI("Colors", "GenreDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblInfo.ForeColor = Hex2VB(ReadINI("Colors", "InfoDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTime.ForeColor = Hex2VB(ReadINI("Colors", "TimeDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "TotalTimeDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
  'End If
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgStop_MouseMove Button, Shift, x, Y
End Sub

Private Sub imgStop_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 0 Then
    If x >= 0 And x <= imgStop.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgStop.Height * Screen.TwipsPerPixelY Then
      imgStop.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "StopDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgStop.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "StopUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgStop.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "StopUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub Keys_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyZ
      imgPrev_Click
    Case vbKeyX
      imgPlayPause_Click
    Case vbKeyC
      imgStop_Click
    Case vbKeyV
      imgNext_Click
    Case vbKeyP
      imgPlaylist_Click
    Case vbKeyR
      imgRepeat_Click
    Case vbKeyS
      imgShuffle_Click
  End Select
End Sub

Private Sub lblComments_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Public Sub PicSpectrum1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then Spectrum = Spectrum + 1
  If Button = 2 Then Spectrum = Spectrum - 1
  If Spectrum = 0 Then Spectrum = 6
  If Spectrum = 7 Then Spectrum = 1
  If Spectrum = 1 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum1Background", App.Path & "\skins\" & CurrentSkin & ".ini"))
  ElseIf Spectrum = 2 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum2Background", App.Path & "\skins\" & CurrentSkin & ".ini"))
  ElseIf Spectrum = 3 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum3Background", App.Path & "\skins\" & CurrentSkin & ".ini"))
  ElseIf Spectrum = 4 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum4Background", App.Path & "\skins\" & CurrentSkin & ".ini"))
  ElseIf Spectrum = 5 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Spectrum5Background", App.Path & "\skins\" & CurrentSkin & ".ini"))
  ElseIf Spectrum = 6 Then
    PicSpectrum1.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SpectrumBackgroundOff", App.Path & "\skins\" & CurrentSkin & ".ini"))
  End If
End Sub

Private Sub PicSpectrum1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub Position_Change()
  If DMC.StreamIsActive Then
    DMC.StreamPos = Position.Value
  End If
End Sub

Private Sub tmrCounter_Timer()
  If DMC.StreamIsActive Then
    TimePlayedNow = TimePlayedNow + 1
  End If
End Sub

Private Sub tmrMain_Timer()
  If DMC.StreamIsActive And WindowState = vbNormal And Visible Then
    Position.Max = DMC.StreamLen
    Position.Value = DMC.StreamPos
    lblTime = ConvertTime(DMC.StreamPosInSeconds)
    lblTotalTime = ConvertTime(DMC.StreamLenInSeconds)
    DrawScope Spectrum
  End If
End Sub

Private Sub Volume_Change()
  CurrentVolume = Volume.Value
  DMC.StreamVol = Volume.Value
End Sub

Public Sub Play()
  'This will start playing CurrentlyPlaying file
  Dim MP3 As New Mp3Info
  Dim sTemp As String, x As Long
  
  If CurrentlyPlaying = 0 Then PlayNext
  
  'If the file exists
  If frmPlaylist.lvwList.ListItems.Count >= CurrentlyPlaying And FileExists(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename) Then
    
    For x = 1 To frmPlaylist.lvwList.ListItems.Count
      If frmPlaylist.lvwList.ListItems.Item(x).Bold Then frmPlaylist.lvwList.ListItems.Item(x).Bold = False
    Next x
    frmPlaylist.lvwList.ListItems.Item(CurrentlyPlaying).Bold = True
    frmPlaylist.lvwList.ListItems.Item(CurrentlyPlaying).EnsureVisible
    frmPlaylist.Scroll.Value = CurrentlyPlaying
    
    
    'Load it
    DMC.OpenStream Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename
    DMC.StreamVol = CurrentVolume
    
    'Setup label colors
    lblArtistTitle.ForeColor = Hex2VB(ReadINI("Colors", "ArtistEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblAlbum.ForeColor = Hex2VB(ReadINI("Colors", "AlbumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblYear.ForeColor = Hex2VB(ReadINI("Colors", "YearEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblComments.ForeColor = Hex2VB(ReadINI("Colors", "CommentsEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblGenre.ForeColor = Hex2VB(ReadINI("Colors", "GenreEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblInfo.ForeColor = Hex2VB(ReadINI("Colors", "InfoEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTime.ForeColor = Hex2VB(ReadINI("Colors", "TimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "TotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblComments.ToolTipText = ""
    
    'Updates mp3 info, first trying ID3v2 then ID3v1
    If ReadID3v2(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename) Then
      If Len(ID3v2Info.Artist) > 0 And Len(ID3v2Info.Title) > 0 Then
        lblArtistTitle = ID3v2Info.Artist & " - " & ID3v2Info.Title
        cTray.ToolTip = lblArtistTitle
        Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Artist = ID3v2Info.Artist
        Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Title = ID3v2Info.Title
      Else
        sTemp = Right(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename, Len(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename) - InStrRev(Playlist(CurrentlyPlaying).Filename, "\"))
        lblArtistTitle = Left(sTemp, Len(sTemp) - 4)
        cTray.ToolTip = lblArtistTitle
      End If
      If Len(ID3v2Info.Album) > 0 Then
        lblAlbum = ID3v2Info.Album
        Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Album = ID3v2Info.Album
      Else
        lblAlbum = "Album N/A"
      End If
      If Len(ID3v2Info.Year) > 0 Then
        lblYear = ID3v2Info.Year
      Else
        lblYear = "N/A"
      End If
      If Len(ID3v2Info.Comments) > 0 Then
        lblComments = ID3v2Info.Comments
        If TextWidth(lblComments) > lblComments.Width Then
          lblComments.ToolTipText = ID3v2Info.Comments
        End If
      Else
        lblComments = "Comments N/A"
      End If
      If Len(ID3v2Info.Genre) > 0 Then
        lblGenre = TrimGenre(ID3v2Info.Genre)
        Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Genre = TrimGenre(ID3v2Info.Genre)
      Else
        lblGenre = "N/A"
      End If
    Else  'ID3v1
      If GetID3(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename) Then
        If Len(Trim(CleanString(ID3v1Info.Artist))) > 0 And Len(CleanString(Trim(ID3v1Info.Title))) > 0 Then
          lblArtistTitle = Trim(CleanString(ID3v1Info.Artist)) & " - " & Trim(CleanString(ID3v1Info.Title))
          cTray.ToolTip = Trim(CleanString(ID3v1Info.Artist)) & " - " & Trim(CleanString(ID3v1Info.Title))
          Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Artist = Trim(CleanString(ID3v1Info.Artist))
          Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Title = Trim(CleanString(ID3v1Info.Title))
        Else
          sTemp = Right(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename, Len(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename) - InStrRev(Playlist(CurrentlyPlaying).Filename, "\"))
          lblArtistTitle = Left(sTemp, Len(sTemp) - 4)
          cTray.ToolTip = lblArtistTitle
        End If
        If Len(Trim(CleanString(ID3v1Info.Album))) > 0 Then
          lblAlbum = Trim(CleanString(ID3v1Info.Album))
          Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Album = Trim(CleanString(ID3v1Info.Album))
        Else
          lblAlbum = "Album N/A"
        End If
        If Len(Trim(CleanString(ID3v1Info.Year))) > 0 Then
          lblYear = Trim(CleanString(ID3v1Info.Year))
        Else
          lblYear = "N/A"
        End If
        If Len(Trim(CleanString(ID3v1Info.Comments))) > 0 Then
          If ID3v1Info.IsTrack <> 0 Then
            lblComments = Trim(CleanString(ID3v1Info.Comments) & Chr(ID3v1Info.IsTrack) & Chr(ID3v1Info.Tracknumber))
          Else
            lblComments = Trim(CleanString(ID3v1Info.Comments))
          End If
        Else
          lblComments = "Comments N/A"
        End If
        If Len(DMC.GetGenreDescrip(ID3v1Info.Genre)) > 0 Then
          lblGenre = DMC.GetGenreDescrip(ID3v1Info.Genre)
          Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Genre = DMC.GetGenreDescrip(ID3v1Info.Genre)
        Else
          lblGenre = "N/A"
        End If
      Else  'If neither ID3v1 nor ID3v2 was found
        sTemp = Right(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename, Len(Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename) - InStrRev(Playlist(CurrentlyPlaying).Filename, "\"))
        lblArtistTitle = Left(sTemp, Len(sTemp) - 4)
        cTray.ToolTip = lblArtistTitle
        lblAlbum = "Album N/A"
        lblComments = "Comments N/A"
        lblYear = "N/A"
        lblGenre = "N/A"
      End If
    End If
    
    If NoID3 Then
      frmPlaylist.lvwList.ListItems(CurrentlyPlaying).Text = lblArtistTitle
      If lblAlbum <> "Album N/A" Then frmPlaylist.lvwList.ListItems(CurrentlyPlaying).SubItems(1) = lblAlbum
      If lblGenre <> "N/A" Then frmPlaylist.lvwList.ListItems(CurrentlyPlaying).SubItems(2) = lblGenre
    End If
    
    'Get bitrate & frequency
    MP3.Filename = Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Filename
    MP3.GetMPEGInfo
    lblInfo = MP3.BitRate & " kbps " & Left(MP3.Frequency, 2) & " khz"
    
    'Set Stereo/Mono image
    If DMC.StreamIsMono Then
      imgStereoMono.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Mono", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgStereoMono.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "Stereo", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
    'Show total time
    lblTotalTime = ConvertTime(DMC.StreamLenInSeconds)
    
    'Update play & pause button to pause
    imgPlayPause.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PauseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    
    'Play!
    DMC.PlayStream False
    
    Playlist(frmPlaylist.lvwList.ListItems(CurrentlyPlaying).TaG).Length = DMC.StreamLenInSeconds
    frmPlaylist.lvwList.ListItems(CurrentlyPlaying).SubItems(3) = ConvertTime(DMC.StreamLenInSeconds, True)
    If NoID3 Then
      sTemp = 0
      TotalPlaylistLength = 0
      For x = 1 To UBound(Playlist)
        If Playlist(x).Length > 0 Then
          TotalPlaylistLength = TotalPlaylistLength + Playlist(x).Length
          sTemp = sTemp + 1
        End If
      Next x
      If sTemp < UBound(Playlist) Then
        frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength, True) & "+"
      Else
        frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength, True)
      End If
    End If
    
    
    If frmPitch.Visible Then frmPitch.Update
  End If
End Sub

Public Sub PlayNext()
  'This plays the next song in the list
  Dim x As Long
  
  If frmPlaylist.lvwList.ListItems.Count > 0 Then
    If Shuffle Then
      Randomize Timer
      If frmPlaylist.lvwList.ListItems.Count > 1 Then
ReRnd:
        x = Int(Rnd * frmPlaylist.lvwList.ListItems.Count) + 1
        If x = 0 Then GoTo ReRnd
        If x <> CurrentlyPlaying Then
          CurrentlyPlaying = x
        Else
          GoTo ReRnd
        End If
      Else
        CurrentlyPlaying = 1
      End If
      Play
    Else
      If CurrentlyPlaying < frmPlaylist.lvwList.ListItems.Count Then
        CurrentlyPlaying = CurrentlyPlaying + 1
        Play
      Else
        If Repeat Then
          CurrentlyPlaying = 1
          Play
        Else
          If DMC.StreamPos = DMC.StreamLen Then imgStop_Click
        End If
      End If
    End If
  End If
End Sub

Public Sub PlayPrev()
  'Plays previous item in playlist
  Dim x As Long
  
  If frmPlaylist.lvwList.ListItems.Count > 0 Then
    If Shuffle Then
      Randomize Timer
      If frmPlaylist.lvwList.ListItems.Count > 1 Then
ReRnd:
        x = Int(Rnd * frmPlaylist.lvwList.ListItems.Count) + 1
        If x = 0 Then GoTo ReRnd
        If x <> CurrentlyPlaying Then
          CurrentlyPlaying = x
        Else
          GoTo ReRnd
        End If
      Else
        CurrentlyPlaying = 1
      End If
      Play
    Else
      If CurrentlyPlaying > 1 Then
        CurrentlyPlaying = CurrentlyPlaying - 1
        Play
      Else
        If Repeat Then
          CurrentlyPlaying = frmPlaylist.lvwList.ListItems.Count
          Play
        Else
          If DMC.StreamPos = DMC.StreamLen Then imgStop_Click
        End If
      End If
    End If
    For x = 1 To frmPlaylist.lvwList.ListItems.Count
      If frmPlaylist.lvwList.ListItems.Item(x).Bold Then frmPlaylist.lvwList.ListItems.Item(x).Bold = False
    Next x
    frmPlaylist.lvwList.ListItems.Item(CurrentlyPlaying).Bold = True
    frmPlaylist.lvwList.ListItems.Item(CurrentlyPlaying).EnsureVisible
    frmPlaylist.Scroll.Value = CurrentlyPlaying
  End If
End Sub

Private Sub DrawScope(Scope)
  'Draws the Scopes
  Static x As Long, h As Long, i As Long
  Dim SampleData(1000) As Integer, nDataSize As Integer
  Dim posL As Long, posR As Long, Pos As Long
  
  Select Case Scope
  Case 1

    If DMC.StreamIsMono Then
      nDataSize = 500
    Else
      nDataSize = 1000
    End If
    DMC.StreamData SampleData, nDataSize

    PicSpectrum1.Cls

    If DMC.StreamIsMono Then
      'mono channel
      For i = 0 To 499 Step 1
        h = ((SampleData(i) + 32768) / 65535 * PicSpectrum1.ScaleHeight)
        x = (PicSpectrum1.ScaleWidth * i * 2) / nDataSize
        If i = 0 Then PicSpectrum1.PSet (0, h)
        PicSpectrum1.Line -(x, h), Color1L
      Next
    Else
      'left channel
      For i = 0 To 499 Step 1
        h = ((SampleData(i) + 32768) / 65535 * PicSpectrum1.ScaleHeight)
        x = (PicSpectrum1.ScaleWidth * i * 2) / nDataSize
        If i = 0 Then PicSpectrum1.PSet (0, h)
        PicSpectrum1.Line -(x, h), Color1L
      Next
      'right channel
      For i = 1 To 499 Step 2
        h = ((SampleData(i) + 32768) / 65535 * PicSpectrum1.ScaleHeight)
        x = (PicSpectrum1.ScaleWidth * i) / 500
        If i = 1 Then PicSpectrum1.PSet (0, h)
        PicSpectrum1.Line -(x, h), Color1R
      Next
    End If
  
  Case 2
  
    If DMC.StreamIsMono Then
      nDataSize = 500
    Else
      nDataSize = 1000
    End If

    DMC.StreamData SampleData, nDataSize
    PicSpectrum1.Cls

    For i = 0 To 499 Step 1
      h = (PicSpectrum1.ScaleHeight / 2) + (SampleData(i) / 72)
      x = (PicSpectrum1.ScaleWidth * i * 2) / nDataSize
      If i = 0 Then PicSpectrum1.PSet (0, h)
      PicSpectrum1.Line (x, PicSpectrum1.ScaleHeight / 2)-(x, h), Color2, BF
    Next
    
  Case 3
  
    PicSpectrum1.Cls
  
    posL = Int((PicSpectrum1.ScaleWidth / 128) * DMC.StreamLeftLevel)
    posR = Int((PicSpectrum1.ScaleWidth / 128) * DMC.StreamRightLevel)
  
    If posL > 0 Then
      PicSpectrum1.PaintPicture SpectrumBar.Picture, 30, 30, , , 0, 0, posL
    End If
    If posR > 0 Then
      PicSpectrum1.PaintPicture SpectrumBar.Picture, 30, 240, , , 0, 0, posR
    End If
  
  Case 4
  
    If DMC.StreamIsMono Then
      nDataSize = 500
    Else
      nDataSize = 1000
    End If

    DMC.StreamData SampleData, nDataSize
    PicSpectrum1.Cls
  
    i = Int(PicSpectrum1.ScaleWidth / 20)
    For x = 0 To 19
      If SampleData(25 * x) > ScopeArray(x) Then
        ScopeArray(x) = SampleData(25 * x)
      Else
        ScopeArray(x) = ScopeArray(x) - 750
      End If
      PicSpectrum1.Line ((x * i) + 20, PicSpectrum1.ScaleHeight)-((x * i) + (i - 30), PicSpectrum1.ScaleHeight - (ScopeArray(x) / 52)), Color4, BF
    Next x
    
  Case 5
  
    Pos = Int(((PicSpectrum1.ScaleHeight - 30) / 256) * (DMC.StreamLeftLevel + DMC.StreamRightLevel))
    If Pos < 1 Then Pos = 1

    PictureLoader.Width = PicSpectrum1.Width
    PictureLoader.Height = PicSpectrum1.Height
    PictureLoader.Cls
    BitBlt PictureLoader.hdc, 0, 0, 154, 29, PicSpectrum1.hdc, 1, 1, vbSrcCopy
    PicSpectrum1.Cls
    BitBlt PicSpectrum1.hdc, 1, 1, 153, 29, PictureLoader.hdc, 1, 0, vbSrcCopy
    PicSpectrum1.Line (PicSpectrum1.ScaleWidth - 30, PicSpectrum1.ScaleHeight - 30)-(PicSpectrum1.ScaleWidth - 30, PicSpectrum1.ScaleHeight - Pos - 30), Color5
    
  End Select
    
End Sub

Public Sub SetupSystray(Mode As Boolean)
  If Mode Then
    cTray.IconInSysTray
  Else
    cTray.RemoveFromSysTray
  End If
End Sub

Public Sub SetupSystrayIcon(Icon As Long)
  cTray.Icon = imlIcons.ListImages(Icon).Picture
End Sub
