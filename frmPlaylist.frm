VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlaylist 
   BorderStyle     =   0  'None
   Caption         =   "Playlist"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog ListSaveLoad 
      Left            =   5520
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select File"
      Filter          =   "Simple Amp Playlist (*.playlist)|*.playlist"
   End
   Begin MSComDlg.CommonDialog Open 
      Left            =   4920
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add file..."
      Filter          =   "Supported files (*.mp3;*.mp2;*.mp1)|*.mp3;*.mp2;*.mp1"
   End
   Begin SimpleAmp.PicVScroll Scroll 
      Height          =   4440
      Left            =   8490
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   645
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   7832
      Min             =   1
      Max             =   2
      Value           =   1
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   4440
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7832
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Artist & Title"
         Object.Width           =   7276
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Album"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Genre"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Length"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Filename"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image imgSubdirs 
      Height          =   255
      Left            =   1920
      ToolTipText     =   "Search in subdirs"
      Top             =   5520
      Width           =   2100
   End
   Begin VB.Image imgColumns 
      Height          =   255
      Left            =   135
      ToolTipText     =   "List Columns"
      Top             =   405
      Width           =   8370
   End
   Begin VB.Image imgList 
      Height          =   375
      Left            =   8280
      ToolTipText     =   "List Options"
      Top             =   5475
      Width           =   495
   End
   Begin VB.Image imgSelect 
      Height          =   375
      Left            =   1320
      ToolTipText     =   "Select"
      Top             =   5475
      Width           =   495
   End
   Begin VB.Label lblTotalTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7155
      TabIndex        =   2
      ToolTipText     =   "Total length of files in playlist"
      Top             =   5100
      Width           =   1215
   End
   Begin VB.Label lblTotalNum 
      BackStyle       =   0  'Transparent
      Caption         =   "0 files."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   150
      TabIndex        =   1
      ToolTipText     =   "Total files in playlist"
      Top             =   5100
      Width           =   2055
   End
   Begin VB.Image imgRem 
      Height          =   375
      Left            =   720
      ToolTipText     =   "Remove"
      Top             =   5475
      Width           =   495
   End
   Begin VB.Image imgAdd 
      Height          =   375
      Left            =   120
      ToolTipText     =   "Add"
      Top             =   5475
      Width           =   495
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   8640
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgDropdown 
      Height          =   5895
      Left            =   0
      OLEDropMode     =   1  'Manual
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MovePlaylist As Boolean 'True if in move mode
Dim MovePlaylistOldX As Long
Dim MovePlaylistOldY As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  lvwList_KeyDown KeyCode, Shift
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Get windows work area size to type variable Scr
  SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
  MovePlaylist = True
  MovePlaylistOldX = X
  MovePlaylistOldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If MovePlaylist Then
    If Snap Then  'If Snap window option is on, snap it to screen edges
      'Snap to right or left edge of screen
      If (Left + (X - MovePlaylistOldX) + Width) / Screen.TwipsPerPixelX > Scr.Right - SnapWidth And (Left + (X - MovePlaylistOldX) + Width) / Screen.TwipsPerPixelX < Scr.Right + SnapWidth Then
        Left = (Scr.Right * Screen.TwipsPerPixelX) - Width
      ElseIf (Left + (X - MovePlaylistOldX)) / Screen.TwipsPerPixelX < Scr.Left + SnapWidth And (Left + (X - MovePlaylistOldX)) / Screen.TwipsPerPixelX > Scr.Left - SnapWidth Then
        Left = (Scr.Left * Screen.TwipsPerPixelX)
      Else
        Left = Left + (X - MovePlaylistOldX)
      End If
      'Snap to lower or upper edge of screen
      If (Top + (Y - MovePlaylistOldY) + Height) / Screen.TwipsPerPixelY > Scr.Bottom - SnapWidth And (Top + (Y - MovePlaylistOldY) + Height) / Screen.TwipsPerPixelY < Scr.Bottom + SnapWidth Then
        Top = (Scr.Bottom * Screen.TwipsPerPixelY) - Height
      ElseIf (Top + (Y - MovePlaylistOldY)) / Screen.TwipsPerPixelY < Scr.Top + SnapWidth And (Top + (Y - MovePlaylistOldY)) / Screen.TwipsPerPixelY > Scr.Top - SnapWidth Then
        Top = (Scr.Top * Screen.TwipsPerPixelY)
      Else
        Top = Top + (Y - MovePlaylistOldY)
      End If
    Else
      Left = Left + (X - MovePlaylistOldX)
      Top = Top + (Y - MovePlaylistOldY)
    End If
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MovePlaylist = False
End Sub

Private Sub imgAdd_Click()
  PopupMenu frmMenus.menAdd
End Sub

Private Sub imgAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgAdd_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    If X >= 0 And X <= imgAdd.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgAdd.Height * Screen.TwipsPerPixelY Then
      imgAdd.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "AddDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgAdd.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "AddUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgAdd.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "AddUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgClose_Click()
  Hide
  frmMain.imgPlaylist.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgClose_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    If X >= 0 And X <= imgClose.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgClose.Height * Screen.TwipsPerPixelY Then
      imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistCloseDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistCloseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgClose.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "PlaylistCloseUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgDropdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Get windows work area size to type variable Scr
  SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
  MovePlaylist = True
  MovePlaylistOldX = X
  MovePlaylistOldY = Y
End Sub

Private Sub imgDropdown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If MovePlaylist Then
    If Snap Then  'If Snap window option is on, snap it to screen edges
      'Snap to right or left edge of screen
      If (Left + (X - MovePlaylistOldX) + Width) / Screen.TwipsPerPixelX > Scr.Right - SnapWidth And (Left + (X - MovePlaylistOldX) + Width) / Screen.TwipsPerPixelX < Scr.Right + SnapWidth Then
        Left = (Scr.Right * Screen.TwipsPerPixelX) - Width
      ElseIf (Left + (X - MovePlaylistOldX)) / Screen.TwipsPerPixelX < Scr.Left + SnapWidth And (Left + (X - MovePlaylistOldX)) / Screen.TwipsPerPixelX > Scr.Left - SnapWidth Then
        Left = (Scr.Left * Screen.TwipsPerPixelX)
      Else
        Left = Left + (X - MovePlaylistOldX)
      End If
      'Snap to lower or upper edge of screen
      If (Top + (Y - MovePlaylistOldY) + Height) / Screen.TwipsPerPixelY > Scr.Bottom - SnapWidth And (Top + (Y - MovePlaylistOldY) + Height) / Screen.TwipsPerPixelY < Scr.Bottom + SnapWidth Then
        Top = (Scr.Bottom * Screen.TwipsPerPixelY) - Height
      ElseIf (Top + (Y - MovePlaylistOldY)) / Screen.TwipsPerPixelY < Scr.Top + SnapWidth And (Top + (Y - MovePlaylistOldY)) / Screen.TwipsPerPixelY > Scr.Top - SnapWidth Then
        Top = (Scr.Top * Screen.TwipsPerPixelY)
      Else
        Top = Top + (Y - MovePlaylistOldY)
      End If
    Else
      Left = Left + (X - MovePlaylistOldX)
      Top = Top + (Y - MovePlaylistOldY)
    End If
  End If
End Sub

Private Sub imgDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MovePlaylist = False
End Sub

Private Sub imgDropdown_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Adds mp3s dropped in window to playlist
  Dim i As Long, Upper As Long
  Dim MP3 As New Mp3Info
  Dim ListAdd
  Dim AddNew As Boolean
  
  If Data.GetFormat(vbCFFiles) Then 'If data is list of files
    
    If lvwList.ListItems.Count > 0 Then
      AddNew = False
    Else
      AddNew = True
    End If
    
    For i = 1 To Data.Files.Count 'add each dropped file
      If LCase(Right(Data.Files(i), 4)) = ".mp3" Then
        
        Upper = UBound(Playlist)
        ReDim Preserve Playlist(Upper + 1) As PlaylistData
        
        'Get ID3v2. If there is none, get ID3v1
        Playlist(Upper + 1).FileName = Data.Files(i)
        If Not NoID3 Then
          If ReadID3v2(Data.Files(i)) Then
            Playlist(Upper + 1).Album = ID3v2Info.Album
            Playlist(Upper + 1).Artist = ID3v2Info.Artist
            Playlist(Upper + 1).Genre = TrimGenre(ID3v2Info.Genre)
            Playlist(Upper + 1).Title = ID3v2Info.Title
          Else
            If GetID3(Data.Files(i)) Then
              Playlist(Upper + 1).Title = Trim(CleanString(ID3v1Info.Title))
              Playlist(Upper + 1).Artist = Trim(CleanString(ID3v1Info.Artist))
              Playlist(Upper + 1).Album = Trim(CleanString(ID3v1Info.Album))
              Playlist(Upper + 1).Genre = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre)
            Else
              Playlist(Upper + 1).Title = ""
              Playlist(Upper + 1).Album = ""
              Playlist(Upper + 1).Artist = ""
              Playlist(Upper + 1).Genre = ""
            End If
          End If
        End If
        
        'Get mp3 length
        MP3.FileName = Data.Files(i)
        MP3.GetMPEGInfo
        Playlist(Upper + 1).Length = MP3.Seconds
        Playlist(Upper + 1).Removed = False
        
        TotalPlaylistLength = TotalPlaylistLength + MP3.Seconds
        
        'Add to list
        Set ListAdd = lvwList.ListItems.Add
        
        If Len(Playlist(Upper + 1).Artist) > 0 And Len(Playlist(Upper + 1).Title) > 0 Then
          ListAdd.Text = Playlist(Upper + 1).Artist & " - " & Playlist(Upper + 1).Title
        Else
          ListAdd.Text = Right(Data.Files(i), Len(Data.Files(i)) - InStrRev(Data.Files(i), "\"))
        End If
        ListAdd.SubItems(1) = Playlist(Upper + 1).Album
        ListAdd.SubItems(2) = Playlist(Upper + 1).Genre
        ListAdd.SubItems(3) = ConvertTime(Playlist(Upper + 1).Length)
        ListAdd.SubItems(4) = Playlist(Upper + 1).FileName
        ListAdd.TaG = Upper + 1
      
      End If
    Next i
      
    lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTotalTime = ConvertTime(TotalPlaylistLength)
    lblTotalNum = lvwList.ListItems.Count & " files."
    If lvwList.ListItems.Count > 1 Then
      Scroll.Max = lvwList.ListItems.Count
    End If
    If AddNew Then frmMain.PlayNext
    
  End If
End Sub

Private Sub imgList_Click()
  PopupMenu frmMenus.menList
End Sub

Private Sub imgList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgList_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    If X >= 0 And X <= imgList.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgList.Height * Screen.TwipsPerPixelY Then
      imgList.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ListDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgList.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ListUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgList.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "ListUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgRem_Click()
  PopupMenu frmMenus.menDel
End Sub

Private Sub imgRem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgRem_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgRem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    If X >= 0 And X <= imgRem.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgRem.Height * Screen.TwipsPerPixelY Then
      imgRem.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RemoveDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgRem.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RemoveUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgRem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgRem.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "RemoveUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgSelect_Click()
  PopupMenu frmMenus.menSelect
End Sub

Private Sub imgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgSelect_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    If X >= 0 And X <= imgSelect.Width * Screen.TwipsPerPixelX And Y >= 0 And Y <= imgSelect.Height * Screen.TwipsPerPixelY Then
      imgSelect.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SelectDown", App.Path & "\skins\" & CurrentSkin & ".ini"))
    Else
      imgSelect.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SelectUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
    End If
  End If
End Sub

Private Sub imgSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgSelect.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SelectUp", App.Path & "\skins\" & CurrentSkin & ".ini"))
End Sub

Private Sub imgSubdirs_Click()
  If SearchInSubdirs Then
    SearchInSubdirs = False
    imgSubdirs.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SubdirsOff", App.Path & "\skins\" & CurrentSkin & ".ini"))
  Else
    SearchInSubdirs = True
    imgSubdirs.Picture = LoadPicture(App.Path & "\skins\" & ReadINI("Images", "SubdirsOn", App.Path & "\skins\" & CurrentSkin & ".ini"))
  End If
End Sub

Private Sub lvwList_DblClick()
  Dim X As Integer
  
  If lvwList.ListItems.Count > 0 Then
    CurrentlyPlaying = lvwList.SelectedItem.Index
    For X = 1 To lvwList.ListItems.Count
      If lvwList.ListItems.Item(X).Bold Then lvwList.ListItems.Item(X).Bold = False
    Next X
    lvwList.SelectedItem.Bold = True
    frmMain.Play
  End If
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Scroll.Value = Item.Index
  If frmInfo.Visible Then frmInfo.UpdateInfo
End Sub

Private Sub lvwList_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDelete
      frmMenus.menDeleteFile_Click
    Case vbKeyReturn
      lvwList_DblClick
    Case vbKeyInsert
      If Shift = 1 Then
        frmMenus.menAddDir_Click
      ElseIf Shift = 0 Then
        frmMenus.menAddFile_Click
      End If
    Case vbKeyA
      If Shift = 2 Then frmMenus.menSelectAll_Click
    Case vbKeyI
      If Shift = 2 Then frmMenus.menSelectInvert_Click
    Case vbKeyN
      If Shift = 2 Then frmMenus.menSelectNone_Click
  End Select
End Sub

Private Sub lvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.menRClick
  End If
End Sub

Private Sub lvwList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Adds mp3s dropped in window to playlist
  Dim i As Long, Upper As Long
  Dim MP3 As New Mp3Info
  Dim ListAdd
  Dim AddNew As Boolean
  
  If Data.GetFormat(vbCFFiles) Then 'If data is list of files
    
    If lvwList.ListItems.Count > 0 Then
      AddNew = False
    Else
      AddNew = True
    End If
    
    For i = 1 To Data.Files.Count 'add each dropped file
      If LCase(Right(Data.Files(i), 4)) = ".mp3" Then
        
        Upper = UBound(Playlist)
        ReDim Preserve Playlist(Upper + 1) As PlaylistData
        
        'Get ID3v2. If there is none, get ID3v1
        Playlist(Upper + 1).FileName = Data.Files(i)
        If Not NoID3 Then
          If ReadID3v2(Data.Files(i)) Then
            Playlist(Upper + 1).Album = ID3v2Info.Album
            Playlist(Upper + 1).Artist = ID3v2Info.Artist
            Playlist(Upper + 1).Genre = TrimGenre(ID3v2Info.Genre)
            Playlist(Upper + 1).Title = ID3v2Info.Title
          Else
            If GetID3(Data.Files(i)) Then
              Playlist(Upper + 1).Title = Trim(CleanString(ID3v1Info.Title))
              Playlist(Upper + 1).Artist = Trim(CleanString(ID3v1Info.Artist))
              Playlist(Upper + 1).Album = Trim(CleanString(ID3v1Info.Album))
              Playlist(Upper + 1).Genre = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre)
            Else
              Playlist(Upper + 1).Title = ""
              Playlist(Upper + 1).Album = ""
              Playlist(Upper + 1).Artist = ""
              Playlist(Upper + 1).Genre = ""
            End If
          End If
        End If
        
        'Get mp3 length
        MP3.FileName = Data.Files(i)
        MP3.GetMPEGInfo
        Playlist(Upper + 1).Length = MP3.Seconds
        Playlist(Upper + 1).Removed = False
        
        TotalPlaylistLength = TotalPlaylistLength + MP3.Seconds
        
        'Add to list
        Set ListAdd = lvwList.ListItems.Add
        
        If Len(Playlist(Upper + 1).Artist) > 0 And Len(Playlist(Upper + 1).Title) > 0 Then
          ListAdd.Text = Playlist(Upper + 1).Artist & " - " & Playlist(Upper + 1).Title
        Else
          ListAdd.Text = Right(Data.Files(i), Len(Data.Files(i)) - InStrRev(Data.Files(i), "\"))
        End If
        ListAdd.SubItems(1) = Playlist(Upper + 1).Album
        ListAdd.SubItems(2) = Playlist(Upper + 1).Genre
        ListAdd.SubItems(3) = ConvertTime(Playlist(Upper + 1).Length)
        ListAdd.SubItems(4) = Playlist(Upper + 1).FileName
        ListAdd.TaG = Upper + 1
      
      End If
    Next i
      
    lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    lblTotalTime = ConvertTime(TotalPlaylistLength)
    lblTotalNum = lvwList.ListItems.Count & " files."
    If lvwList.ListItems.Count > 1 Then
      Scroll.Max = lvwList.ListItems.Count
    End If
    If AddNew Then frmMain.PlayNext
    
  End If
End Sub

Private Sub Scroll_Change()
  If lvwList.ListItems.Count > 1 Then
    If Scroll.Value <= lvwList.ListItems.Count Then
      lvwList.ListItems.Item(Scroll.Value).EnsureVisible
    End If
  End If
End Sub
