VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "Stores menus"
   ClientHeight    =   1575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   6180
   Begin VB.Label Label1 
      Caption         =   "This for contains all the rightclick menus and most code for the commands in them. "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Menu menRClick 
      Caption         =   "RClickMenu"
      Begin VB.Menu menPlay 
         Caption         =   "&Play Item"
      End
      Begin VB.Menu menID3 
         Caption         =   "&Edit ID3..."
      End
      Begin VB.Menu menFileInfo 
         Caption         =   "&File Info..."
      End
      Begin VB.Menu menAdd 
         Caption         =   "&Add"
         Begin VB.Menu menAddDir 
            Caption         =   "&Dir..."
         End
         Begin VB.Menu menAddFile 
            Caption         =   "&File..."
         End
      End
      Begin VB.Menu menDel 
         Caption         =   "&Delete"
         Begin VB.Menu menDeleteFile 
            Caption         =   "&File(s)"
         End
         Begin VB.Menu menDeleteAll 
            Caption         =   "&All"
         End
         Begin VB.Menu menCrop 
            Caption         =   "&Crop"
         End
      End
      Begin VB.Menu menSelect 
         Caption         =   "&Select"
         Begin VB.Menu menSelectAll 
            Caption         =   "&All"
         End
         Begin VB.Menu menSelectNone 
            Caption         =   "&None"
         End
         Begin VB.Menu menSelectInvert 
            Caption         =   "&Invert"
         End
      End
      Begin VB.Menu menList 
         Caption         =   "&List"
         Begin VB.Menu menListLoad 
            Caption         =   "&Load..."
         End
         Begin VB.Menu menListSave 
            Caption         =   "&Save..."
         End
         Begin VB.Menu menSort 
            Caption         =   "Sort"
            Begin VB.Menu menSortArtistTitle 
               Caption         =   "Sort list by Artist && &Title"
            End
            Begin VB.Menu menSortAlbum 
               Caption         =   "Sort list by &Album"
            End
            Begin VB.Menu menSortGenre 
               Caption         =   "Sort list by &Genre"
            End
            Begin VB.Menu menSortTime 
               Caption         =   "Sort list by &Time"
            End
            Begin VB.Menu menSortFilename 
               Caption         =   "Sort list by &Filename"
            End
            Begin VB.Menu op 
               Caption         =   "-"
            End
            Begin VB.Menu menReverse 
               Caption         =   "&Reverse list"
            End
         End
      End
   End
   Begin VB.Menu menMain 
      Caption         =   "MainMenu"
      Begin VB.Menu menAbout 
         Caption         =   "&About Simple Amp..."
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu menSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu menSkin 
         Caption         =   "&Change Skin..."
      End
      Begin VB.Menu menScope 
         Caption         =   "S&witch Scope"
      End
      Begin VB.Menu menPitch 
         Caption         =   "P&itch Control..."
      End
      Begin VB.Menu menline 
         Caption         =   "-"
      End
      Begin VB.Menu menPrev 
         Caption         =   "&Previous"
      End
      Begin VB.Menu menPlayPause 
         Caption         =   "P&lay/Pause"
      End
      Begin VB.Menu menStop 
         Caption         =   "S&top"
      End
      Begin VB.Menu menNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu menline2 
         Caption         =   "-"
      End
      Begin VB.Menu menExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menSystray 
      Caption         =   "SystrayMenu"
      Begin VB.Menu menSysOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu menSysline 
         Caption         =   "-"
      End
      Begin VB.Menu menSysPrev 
         Caption         =   "&Previous"
      End
      Begin VB.Menu menSysPlayPause 
         Caption         =   "P&lay/Pause"
      End
      Begin VB.Menu menSysStop 
         Caption         =   "S&top"
      End
      Begin VB.Menu menSysNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu menSysline2 
         Caption         =   "-"
      End
      Begin VB.Menu menSysExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub menAbout_Click()
  'Shows about window
  frmAbout.Show , frmMain
End Sub

Public Sub menAddDir_Click()
  'Adds all files in selected dir
  Dim MP3 As New Mp3Info
  Dim Upper As Long
  Dim ReturnValue As String 'Keeps up the return
  Dim ListAdd, i As Long
  Dim AddNew As Boolean
  
  'Shows browse for dir window
  ReturnValue = BrowseForFolder(Me.hwnd, "Select an folder to search in:")

  If ReturnValue <> "" Then
    ReturnValue = ts.sAppend(ReturnValue, "\")
  
    If Len(cDir) > 0 Then
      ListAdd = ""
      For i = 1 To Len(cDir)
        If Left(cDir, i) = Left(ReturnValue, i) Then
          ListAdd = Left(cDir, i)
        Else
          cDir = ListAdd
          Exit For
        End If
      Next i
    Else
      cDir = ReturnValue
    End If
    
    If frmPlaylist.lvwList.ListItems.Count > 0 Then
      AddNew = False
    Else
      AddNew = True
    End If
  
    'Update forms
    frmPlaylist.MousePointer = vbHourglass
    If OnTop Then
      AlwaysOnTop frmFind, True
    Else
      AlwaysOnTop frmFind, False
    End If
    frmFind.Show , frmPlaylist
    frmFind.lblSearch = "Searching for files..."
    
    'Search for files (really fast!)
    cFiles.LoadFiles ReturnValue & "*.mp3", SearchInSubdirs, frmFind.prbProgress

    'Show found files
    frmFind.lblSearch = "Found " & cFiles.Count & " files."
    
    'If there was any found
    If cFiles.Count > 0 Then
      
      'Update form & variables
      frmFind.prbProgress.Max = cFiles.Count
      Upper = UBound(Playlist)
      
      'Redim array to hold playlist
      ReDim Preserve Playlist(cFiles.Count + Upper) As PlaylistData
      
      'Loop through all new playlist entries
      For i = 1 To cFiles.Count
        'Update & refresh forms
        frmFind.prbProgress.Value = i
        frmFind.lblInfo = "Getting data from " & cFiles(i).sNameAndExtension & "."
        frmFind.Caption = "Working... [" & Int((frmFind.prbProgress.Value / frmFind.prbProgress.Max) * 100) & "%]"
        frmFind.Refresh
        
        'Add item to array
        'Get ID3v2. If there is none, get ID3v1
        Playlist(i + Upper).Filename = cFiles(i).sFilename
        If Not NoID3 Then
          If ReadID3v2(cFiles(i).sFilename) Then
            Playlist(i + Upper).Album = ID3v2Info.Album
            Playlist(i + Upper).Artist = ID3v2Info.Artist
            Playlist(i + Upper).Genre = TrimGenre(ID3v2Info.Genre)
            Playlist(i + Upper).Title = ID3v2Info.Title
          Else
            If GetID3(cFiles(i).sFilename) Then 'ID3v1
              Playlist(i + Upper).Title = Trim(CleanString(ID3v1Info.Title))
              Playlist(i + Upper).Artist = Trim(CleanString(ID3v1Info.Artist))
              Playlist(i + Upper).Album = Trim(CleanString(ID3v1Info.Album))
              Playlist(i + Upper).Genre = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre)
            Else
              Playlist(i + Upper).Title = ""
              Playlist(i + Upper).Album = ""
              Playlist(i + Upper).Artist = ""
              Playlist(i + Upper).Genre = ""
            End If
          End If
          
          'Get mp3 length
          MP3.Filename = cFiles(i).sFilename
          MP3.GetMPEGTime
          Playlist(i + Upper).Length = MP3.Seconds
          Playlist(i + Upper).Removed = False
          
          'Add this mp3s length to total
          TotalPlaylistLength = TotalPlaylistLength + MP3.Seconds
        End If
      Next i
      
      'Loop through each found mp3 and this time, add them to the list
      For i = 1 To cFiles.Count
        'Update & refresh form
        frmFind.prbProgress.Value = i
        frmFind.lblInfo = "Adding " & cFiles(i).sNameAndExtension & " to list..."
        frmFind.Caption = "Working... [" & Int((frmFind.prbProgress.Value / frmFind.prbProgress.Max) * 100) & "%]"
        frmFind.Refresh
        
        'Setup Listadd to simplify adding
        Set ListAdd = frmPlaylist.lvwList.ListItems.Add
          
          'Add item to each field of the list
          If Len(Playlist(i + Upper).Artist) > 0 And Len(Playlist(i + Upper).Title) > 0 Then
            ListAdd.Text = Playlist(i + Upper).Artist & " - " & Playlist(i + Upper).Title
          Else
            ListAdd.Text = cFiles(i).sNameAndExtension
          End If
          ListAdd.SubItems(1) = Playlist(i + Upper).Album
          ListAdd.SubItems(2) = Playlist(i + Upper).Genre
          If NoID3 Then
            ListAdd.SubItems(3) = ConvertTime(Playlist(i + Upper).Length, True)
          Else
            ListAdd.SubItems(3) = ConvertTime(Playlist(i + Upper).Length)
          End If
          ListAdd.SubItems(4) = Playlist(i + Upper).Filename
          ListAdd.TaG = i + Upper
      Next i
      
    End If
    
    'updates form
    frmFind.Hide
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength)
    frmPlaylist.lblTotalNum = frmPlaylist.lvwList.ListItems.Count & " files."
    If frmPlaylist.lvwList.ListItems.Count > 1 Then
      frmPlaylist.Scroll.Max = frmPlaylist.lvwList.ListItems.Count
    End If
    frmPlaylist.MousePointer = vbDefault
    If AddNew Then
      CurrentlyPlaying = 0
      frmMain.PlayNext
    End If
  End If
End Sub

Public Sub menAddFile_Click()
  'Adds one file
  Dim Upper As Long, i As Long
  Dim MP3 As New Mp3Info
  Dim ListAdd
  Dim AddNew As Boolean
  Dim sTemp As String
  
  'Show open dialog box
  frmPlaylist.Open.ShowOpen
  If Len(frmPlaylist.Open.Filename) > 0 Then
    If FileExists(frmPlaylist.Open.Filename) Then
      'Setup variables & redim
      Upper = UBound(Playlist)
      ReDim Preserve Playlist(Upper + 1) As PlaylistData
      
      sTemp = Left(frmPlaylist.Open.Filename, InStrRev(frmPlaylist.Open.Filename, "\"))
      If Len(cDir) > 0 Then
        ListAdd = ""
        For i = 1 To Len(cDir)
          If Left(cDir, i) = Left(sTemp, i) Then
            ListAdd = Left(cDir, i)
          Else
            cDir = ListAdd
            Exit For
          End If
        Next i
      Else
        cDir = sTemp
      End If
      
      If frmPlaylist.lvwList.ListItems.Count > 0 Then
        AddNew = False
      Else
        AddNew = True
      End If
      
      'Get ID3v2. If there is none, get ID3v1
      Playlist(Upper + 1).Filename = frmPlaylist.Open.Filename
      If ReadID3v2(frmPlaylist.Open.Filename) Then
        Playlist(Upper + 1).Album = ID3v2Info.Album
        Playlist(Upper + 1).Artist = ID3v2Info.Artist
        Playlist(Upper + 1).Genre = TrimGenre(ID3v2Info.Genre)
        Playlist(Upper + 1).Title = ID3v2Info.Title
      Else
        If GetID3(frmPlaylist.Open.Filename) Then
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

      'Get mp3 length
      MP3.Filename = frmPlaylist.Open.Filename
      MP3.GetMPEGInfo
      Playlist(Upper + 1).Length = MP3.Seconds
      Playlist(Upper + 1).Removed = False
      
      'Add this mp3s length to total
      TotalPlaylistLength = TotalPlaylistLength + MP3.Seconds
      
      'Add to list
      Set ListAdd = frmPlaylist.lvwList.ListItems.Add
        
      If Len(Playlist(Upper + 1).Artist) > 0 And Len(Playlist(Upper + 1).Title) > 0 Then
        ListAdd.Text = Playlist(Upper + 1).Artist & " - " & Playlist(Upper + 1).Title
      Else
        ListAdd.Text = Right(frmPlaylist.Open.Filename, Len(frmPlaylist.Open.Filename) - InStrRev(frmPlaylist.Open.Filename, "\"))
      End If
      ListAdd.SubItems(1) = Playlist(Upper + 1).Album
      ListAdd.SubItems(2) = Playlist(Upper + 1).Genre
      ListAdd.SubItems(3) = ConvertTime(Playlist(Upper + 1).Length)
      ListAdd.SubItems(4) = Playlist(Upper + 1).Filename
      ListAdd.TaG = Upper + 1
      
      'Update form
      frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
      frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
      frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength)
      frmPlaylist.lblTotalNum = frmPlaylist.lvwList.ListItems.Count & " files."
      If frmPlaylist.lvwList.ListItems.Count > 1 Then
        frmPlaylist.Scroll.Max = frmPlaylist.lvwList.ListItems.Count
      End If
      If AddNew Then
        CurrentlyPlaying = 0
        frmMain.PlayNext
      End If
    End If
  End If
End Sub

Private Sub menCrop_Click()
  'Removes files NOT selected
  Dim x As Long
  Dim Count As Long
  
  'loop through all items in the list last to first
  For x = frmPlaylist.lvwList.ListItems.Count To 1 Step -1
    'If the current item ISN'T selected
    If Not frmPlaylist.lvwList.ListItems.Item(x).Selected Then
      'Set to be removed
      Playlist(frmPlaylist.lvwList.ListItems.Item(x).TaG).Removed = True
      'Reduce playlist length
      TotalPlaylistLength = TotalPlaylistLength - Playlist(frmPlaylist.lvwList.ListItems.Item(x).TaG).Length
      'Remove item from list
      frmPlaylist.lvwList.ListItems.Remove x
      'Add one to cound to keep track of how many was removed
      Count = Count + 1
    End If
  Next x

  'Cleans up the playlist from unused entries
  CleanUpPlaylist Count

  'Updates form
  If frmPlaylist.lvwList.ListItems.Count > 0 Then
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength)
    frmPlaylist.lblTotalNum = frmPlaylist.lvwList.ListItems.Count & " files."
  Else
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalTime = "00:00"
    frmPlaylist.lblTotalNum = "0 files."
    cDir = ""
  End If
  'Update scroller
  If frmPlaylist.lvwList.ListItems.Count > 1 Then
    frmPlaylist.Scroll.Value = 1
    frmPlaylist.Scroll.Max = frmPlaylist.lvwList.ListItems.Count
  Else
    frmPlaylist.Scroll.Value = 1
    frmPlaylist.Scroll.Max = 2
  End If
End Sub

Public Sub menDeleteAll_Click()
  'Removes all files
  'Clears srray & list
  ReDim Playlist(0) As PlaylistData
  TotalPlaylistLength = 0
  frmPlaylist.lvwList.ListItems.Clear
  cDir = ""
  'Update form
  frmPlaylist.Scroll.Value = 1
  frmPlaylist.Scroll.Max = 2
  frmPlaylist.Scroll.Min = 1
  frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
  frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
  frmPlaylist.lblTotalTime = "00:00"
  frmPlaylist.lblTotalNum = "0 files."
End Sub

Public Sub menDeleteFile_Click()
  'Removes Selected files
  Dim x As Long
  Dim Count As Long
  
  'loop through all files in list
  For x = frmPlaylist.lvwList.ListItems.Count To 1 Step -1  'reverse
    If frmPlaylist.lvwList.ListItems.Item(x).Selected Then  'if selected
      'Update array
      Playlist(frmPlaylist.lvwList.ListItems.Item(x).TaG).Removed = True
      'Update total playlist length
      TotalPlaylistLength = TotalPlaylistLength - Playlist(frmPlaylist.lvwList.ListItems.Item(x).TaG).Length
      'Remove from list
      frmPlaylist.lvwList.ListItems.Remove x
      Count = Count + 1
    End If
  Next x

  'Cleans up the playlist from unused entries
  CleanUpPlaylist Count

  'updates form
  If frmPlaylist.lvwList.ListItems.Count > 0 Then
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalTime = ConvertTime(TotalPlaylistLength)
    frmPlaylist.lblTotalNum = frmPlaylist.lvwList.ListItems.Count & " files."
  Else
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumDisabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalTime = "00:00"
    frmPlaylist.lblTotalNum = "0 files."
    cDir = ""
  End If
  'updates scroller
  If frmPlaylist.lvwList.ListItems.Count > 1 Then
    frmPlaylist.Scroll.Value = frmPlaylist.lvwList.SelectedItem.Index
    frmPlaylist.Scroll.Max = frmPlaylist.lvwList.ListItems.Count
  Else
    frmPlaylist.Scroll.Value = 1
    frmPlaylist.Scroll.Max = 2
  End If
End Sub

Private Sub menExit_Click()
  Unload frmMain
End Sub

Private Sub menFileInfo_Click()
  frmInfo.Show , frmPlaylist
End Sub

Private Sub menID3_Click()
  frmPlaylist.Enabled = False
  frmId3.Show , frmPlaylist
End Sub

Private Sub menListLoad_Click()
  'Loads playlist
  Dim x As Long, ListAdd, Num As Long
  
  TotalPlaylistLength = 0
  
  'Shows open dialog box
  frmPlaylist.ListSaveLoad.ShowSave
  frmPlaylist.ListSaveLoad.Filename = sAppend(LCase(frmPlaylist.ListSaveLoad.Filename), ".playlist")
  
  If Len(frmPlaylist.ListSaveLoad.Filename) > 0 And FileExists(frmPlaylist.ListSaveLoad.Filename) Then
    If Not LoadPlaylist(frmPlaylist.ListSaveLoad.Filename) Then
      MsgBox "Loading " & frmPlaylist.ListSaveLoad.Filename & " was unsuccessful. The king relinquish his throne!", vbCritical, "Load Error"
      Exit Sub
    End If
    
    'Clear playlist & fill it again
    frmPlaylist.lvwList.ListItems.Clear
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
      
      If Playlist(x).Length > 0 Then Num = Num + 1
      
    Next x
    
    
    'Updates form
    frmPlaylist.lblTotalTime.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalTimeEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    frmPlaylist.lblTotalNum.ForeColor = Hex2VB(ReadINI("Colors", "PlaylistTotalNumEnabledText", App.Path & "\skins\" & CurrentSkin & ".ini"))
    If Num < UBound(Playlist) Then
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
End Sub

Private Sub menListSave_Click()
  'Saves playlist
 
  If frmPlaylist.lvwList.ListItems.Count > 0 Then
    'Shows save dialog box
    frmPlaylist.ListSaveLoad.ShowSave
    If Len(frmPlaylist.ListSaveLoad.Filename) > 0 Then
      frmPlaylist.ListSaveLoad.Filename = sAppend(LCase(frmPlaylist.ListSaveLoad.Filename), ".playlist")
      If Not SavePlaylist(frmPlaylist.ListSaveLoad.Filename) Then
        MsgBox "The save in " & frmPlaylist.ListSaveLoad.Filename & " was unsuccessful. The snake rattles in the darkness.", vbCritical, "Saving Error"
      End If
    End If
  End If
End Sub

Private Sub menNext_Click()
  frmMain.PlayNext
End Sub

Private Sub menPitch_Click()
  frmPitch.Show , frmMain
End Sub

Private Sub menPlay_Click()
  Dim x As Long
  
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    If frmPlaylist.lvwList.ListItems.Item(x).Selected Then
      CurrentlyPlaying = x
      frmMain.Play
      Exit For
    End If
  Next x
End Sub

Private Sub menPlayPause_Click()
  frmMain.imgPlayPause_Click
End Sub

Private Sub menPrev_Click()
  frmMain.PlayPrev
End Sub

Private Sub menReverse_Click()
  If frmPlaylist.lvwList.SortOrder = lvwAscending Then
    frmPlaylist.lvwList.SortOrder = lvwDescending
  Else
    frmPlaylist.lvwList.SortOrder = lvwAscending
  End If
  frmPlaylist.lvwList.Sorted = True
  frmPlaylist.lvwList.Sorted = False
End Sub

Private Sub menScope_Click()
  frmMain.PicSpectrum1_MouseDown 1, 0, 0, 0
End Sub

Public Sub menSelectAll_Click()
  'Selects all items in list
  Dim x As Long
  
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    frmPlaylist.lvwList.ListItems.Item(x).Selected = True
  Next x
End Sub

Public Sub menSelectInvert_Click()
  'Inverts selection in list
  Dim x As Long
  
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    If frmPlaylist.lvwList.ListItems.Item(x).Selected Then
      frmPlaylist.lvwList.ListItems.Item(x).Selected = False
    Else
      frmPlaylist.lvwList.ListItems.Item(x).Selected = True
    End If
  Next x
End Sub

Public Sub menSelectNone_Click()
  'Deselect all items in list
  Dim x As Long
  
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    frmPlaylist.lvwList.ListItems.Item(x).Selected = False
  Next x
End Sub

Private Sub menSettings_Click()
  frmSettings.Show , frmMain
End Sub

Private Sub menSkin_Click()
  frmSkin.Show , frmMain
End Sub

Private Sub menSortAlbum_Click()
  frmPlaylist.lvwList.SortKey = 1
  frmPlaylist.lvwList.SortOrder = lvwAscending
  frmPlaylist.lvwList.Sorted = True
  frmPlaylist.lvwList.Sorted = False
End Sub

Private Sub menSortArtistTitle_Click()
  frmPlaylist.lvwList.SortKey = 0
  frmPlaylist.lvwList.SortOrder = lvwAscending
  frmPlaylist.lvwList.Sorted = True
  frmPlaylist.lvwList.Sorted = False
End Sub

Private Sub menSortFilename_Click()
  frmPlaylist.lvwList.SortKey = 4
  frmPlaylist.lvwList.SortOrder = lvwAscending
  frmPlaylist.lvwList.Sorted = True
  frmPlaylist.lvwList.Sorted = False
End Sub

Private Sub menSortGenre_Click()
  frmPlaylist.lvwList.SortKey = 2
  frmPlaylist.lvwList.SortOrder = lvwAscending
  frmPlaylist.lvwList.Sorted = True
  frmPlaylist.lvwList.Sorted = False
End Sub

Private Sub menSortTime_Click()
  frmPlaylist.lvwList.SortKey = 3
  frmPlaylist.lvwList.SortOrder = lvwAscending
  frmPlaylist.lvwList.Sorted = True
  frmPlaylist.lvwList.Sorted = False
End Sub

Private Sub menStop_Click()
  frmMain.imgStop_Click
End Sub

Private Sub menSysExit_Click()
  menExit_Click
End Sub

Private Sub menSysNext_Click()
  menNext_Click
End Sub

Public Sub menSysOpen_Click()
  frmMain.cTray_LButtonDblClk
End Sub

Private Sub menSysPlayPause_Click()
  menPlayPause_Click
End Sub

Private Sub menSysPrev_Click()
  menPrev_Click
End Sub

Private Sub menSysStop_Click()
  menStop_Click
End Sub

Public Sub CleanUpPlaylist(Num As Long)
  'This sub removes unused entries in the array Playlist()
  Dim x As Long, z As Long, Upper As Long, tempNum As Long
  
  tempNum = Num
  Upper = UBound(Playlist)
  
  If Upper > 0 Then  'make sure there are any entries
    For x = Upper To 1 Step -1 'go tha Reverse style, yo, man cool
      If Playlist(x).Removed Then 'If the entry was marked as not used
        For z = x To Upper - 1 'moves all following entries one step back, thus overwriting this one
          Playlist(z).Album = Playlist(z + 1).Album
          Playlist(z).Artist = Playlist(z + 1).Artist
          Playlist(z).Filename = Playlist(z + 1).Filename
          Playlist(z).Genre = Playlist(z + 1).Genre
          Playlist(z).Length = Playlist(z + 1).Length
          Playlist(z).Removed = Playlist(z + 1).Removed
          Playlist(z).Title = Playlist(z + 1).Title
        Next z
        'loop through every item in the list and subtracts 1 to tag if tag is higher than x
        If frmPlaylist.lvwList.ListItems.Count > 0 Then
          For z = 1 To frmPlaylist.lvwList.ListItems.Count
            If frmPlaylist.lvwList.ListItems(z).TaG > x Then frmPlaylist.lvwList.ListItems(z).TaG = frmPlaylist.lvwList.ListItems(z).TaG - 1
          Next z
        End If
        Num = Num - 1
      End If
      If Num = 0 Then Exit For 'If num is 0 then exit for because there are no more unused entries
    Next x
    
    'redim array to save memory
    ReDim Preserve Playlist(Upper - tempNum) As PlaylistData
  End If
End Sub

Public Function SavePlaylist(FName As String) As Boolean
  'Saves playlist to file FName, returns true if successful
  Dim x As Long
  Dim File As New clsDatafile
  
  
  'Dim Pos As Long 'holds the current file position
  
  On Error GoTo SaveError
  If FileExists(FName) Then Kill FName
  SavePlaylist = True
  
  File.Filename = FName
  File.OpenFile
  
  'Write header
  File.WriteStrFixed "SAMPEXT"
  'Write total number of items in list
  File.WriteLong UBound(Playlist)
  'Write common dir
  File.WriteStr cDir
  For x = 1 To UBound(Playlist)
    'Write filename & reduce its size with common dir if possible
    If Left(Playlist(x).Filename, Len(cDir)) = cDir <> 0 Then
      File.WriteStr CStr(Right(Playlist(x).Filename, Len(Playlist(x).Filename) - Len(cDir)))
    Else
      File.WriteStr Playlist(x).Filename
    End If
    'Write mp3 length in seconds
    File.WriteLong Playlist(x).Length
    'Write ID3 title
    File.WriteStr Playlist(x).Title
    'Write ID3 artist
    File.WriteStr Playlist(x).Artist
    'Write ID3 album
    File.WriteStr Playlist(x).Album
    'Write ID3 genre
    File.WriteStr Playlist(x).Genre
  Next x
  
  Exit Function
SaveError:
  SavePlaylist = False
End Function

Public Function LoadPlaylist(FName As String) As Boolean
  'Loads playlist FName, returns true if successful
  Dim x As Long
  Dim File As New clsDatafile
  
  On Error GoTo LoadError
  LoadPlaylist = True
  
  File.Filename = FName
  File.OpenFile
  
  'Reads header & goes to error if it is not 'SAMPEXT'
  If File.ReadStrFixed(7) <> "SAMPEXT" Then GoTo LoadError
  'Reads total num items
  ReDim Playlist(File.ReadLong) As PlaylistData
  'Reads common dir
  cDir = File.ReadStr
  'Starts getting info for each item
  For x = 1 To UBound(Playlist)
    'Start with filename
    Playlist(x).Filename = cDir & File.ReadStr
    'length in seconds 4 bytes
    Playlist(x).Length = File.ReadLong
    'title
    Playlist(x).Title = File.ReadStr
    'artist
    Playlist(x).Artist = File.ReadStr
    'album
    Playlist(x).Album = File.ReadStr
    'genre
    Playlist(x).Genre = File.ReadStr
  Next x
  
  Exit Function
LoadError:
  ReDim Playlist(0) As PlaylistData
  LoadPlaylist = False
End Function
