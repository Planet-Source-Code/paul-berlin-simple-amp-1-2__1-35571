VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmId3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit ID3"
   ClientHeight    =   6330
   ClientLeft      =   1545
   ClientTop       =   825
   ClientWidth     =   9780
   ControlBox      =   0   'False
   Icon            =   "frmId3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmv1 
      BorderStyle     =   0  'None
      Caption         =   "ID3v1"
      Height          =   5175
      Left            =   5040
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "Multiple File Options"
         Height          =   1095
         Left            =   120
         TabIndex        =   51
         Top             =   3960
         Width           =   4455
         Begin VB.CheckBox chkGrabTitle 
            Caption         =   "&Grab title from filename (between the last ""-"" to "".mp3"")."
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   4215
         End
         Begin VB.CheckBox chkMulti 
            Caption         =   "&Do not write empty fields."
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkTrackerize 
            Caption         =   "&Add track # starting at first selected mp3 in list."
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   4170
         End
      End
      Begin VB.CommandButton cmdv2v1 
         Caption         =   "Copy &from ID3v2"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdv1v2 
         Caption         =   "Copy &to ID3v2"
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CheckBox chkID3v1 
         Caption         =   "&ID3v1 tag"
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   135
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtTrack 
         Height          =   315
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   16
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox txtComments 
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   22
         Top             =   2130
         Width           =   3375
      End
      Begin VB.ComboBox cmbGenre 
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1695
         Width           =   2055
      End
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   20
         Top             =   1695
         Width           =   615
      End
      Begin VB.TextBox txtAlbum 
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   19
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtArtist 
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   18
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   17
         Top             =   450
         Width           =   3375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Track #:"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   1770
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Album:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdRevert 
      Caption         =   "&Revert"
      Height          =   375
      Left            =   1200
      TabIndex        =   28
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame frmv2 
      BorderStyle     =   0  'None
      Caption         =   "ID3v2"
      Height          =   5175
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Width           =   4695
      Visible         =   0   'False
      Begin VB.TextBox txtEncoded 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   4695
         Width           =   3375
      End
      Begin VB.TextBox txtURL 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   4275
         Width           =   3375
      End
      Begin VB.TextBox txtCopyright 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   3855
         Width           =   3375
      End
      Begin VB.TextBox txtOriginalArtist 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   3450
         Width           =   3375
      End
      Begin VB.TextBox txtComposer 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   3030
         Width           =   3375
      End
      Begin VB.TextBox txtTitle2 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   450
         Width           =   3375
      End
      Begin VB.TextBox txtArtist2 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtAlbum2 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtYear2 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1695
         Width           =   615
      End
      Begin VB.ComboBox cmbGenre2 
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1695
         Width           =   2055
      End
      Begin VB.TextBox txtComments2 
         Height          =   795
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2130
         Width           =   3375
      End
      Begin VB.TextBox txtTrack2 
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   30
         Width           =   495
      End
      Begin VB.CheckBox chkID3v2 
         Caption         =   "&ID3v2 tag"
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   135
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Encoded by:"
         Height          =   255
         Left            =   75
         TabIndex        =   50
         Top             =   4770
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4350
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3930
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Orig. Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3525
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Composer:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3105
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Album:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   1770
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Track #:"
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   30
      Top             =   5880
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tabTags 
      Height          =   5655
      Left            =   70
      TabIndex        =   1
      Top             =   50
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9975
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ID3v&1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ID3v&2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmId3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sel As Long

Private Sub chkID3v1_Click()
  If chkID3v1.Value = 1 Then
    txtTitle.Enabled = True
    txtAlbum.Enabled = True
    txtArtist.Enabled = True
    txtTrack.Enabled = True
    txtComments.Enabled = True
    cmbGenre.Enabled = True
    txtYear.Enabled = True
    cmdv1v2.Enabled = True
    tabTags.Tabs(1).Caption = "ID3v&1*"
  Else
    txtTitle.Enabled = False
    txtAlbum.Enabled = False
    txtArtist.Enabled = False
    txtTrack.Enabled = False
    txtComments.Enabled = False
    cmbGenre.Enabled = False
    txtYear.Enabled = False
    cmdv1v2.Enabled = False
    tabTags.Tabs(1).Caption = "ID3v&1"
  End If
End Sub

Private Sub chkID3v2_Click()
  If chkID3v2.Value = 1 Then
    txtTitle2.Enabled = True
    txtAlbum2.Enabled = True
    txtArtist2.Enabled = True
    txtTrack2.Enabled = True
    txtComments2.Enabled = True
    cmbGenre2.Enabled = True
    txtYear2.Enabled = True
    txtComposer.Enabled = True
    txtOriginalArtist.Enabled = True
    txtCopyright.Enabled = True
    txtURL.Enabled = True
    txtEncoded.Enabled = True
    cmdv2v1.Enabled = True
    tabTags.Tabs(2).Caption = "ID3v&2*"
  Else
    txtTitle2.Enabled = False
    txtAlbum2.Enabled = False
    txtArtist2.Enabled = False
    txtTrack2.Enabled = False
    txtComments2.Enabled = False
    cmbGenre2.Enabled = False
    txtYear2.Enabled = False
    txtComposer.Enabled = False
    txtOriginalArtist.Enabled = False
    txtCopyright.Enabled = False
    txtURL.Enabled = False
    txtEncoded.Enabled = False
    cmdv2v1.Enabled = False
    tabTags.Tabs(2).Caption = "ID3v&2"
  End If
End Sub

Private Sub cmdOk_Click()
  frmPlaylist.Enabled = True
  Unload Me
End Sub

Private Sub cmdRevert_Click()
  'Reloads tags
  GetTags
End Sub

Private Sub cmdSave_Click()
  'This saves changes
  Dim tmpCom As String * 30, sTemp As String
  Dim x As Long, Y As Long, SelAdd As Long
  
  Dim oFile As clsFile
  Set oFile = New clsFile
  
  MousePointer = vbHourglass
  
  'Check file statuses
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    If frmPlaylist.lvwList.ListItems(x).Selected Then 'Perform write-only & playing check
    
      'Check if playing
      If CurrentlyPlaying = x And frmMain.DMC.StreamLen <> -1 Then
        If Sel = 1 Then
          MsgBox "The selected MP3 is currently loaded by the player. Tags cannot be changed on MP3's while playing them.", vbExclamation, "Warning"
          MousePointer = vbDefault
          Exit Sub
        Else
          Y = MsgBox("One of the selected MP3's are currently loaded by the player. Tags cannot be changed on MP3's while playing them." & vbNewLine & vbNewLine & "Press OK to continue saving the other MP3's or Cancel to abort.", vbExclamation + vbOKCancel, "Warning")
          If Y = vbCancel Then
            MousePointer = vbDefault
            Exit Sub
          End If
        End If
      End If

again:
      'Check if write-only
      oFile.sFilename = Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName
      If oFile.eAttributes And efaREADONLY Then
        If Sel = 1 Then
          Y = MsgBox("The file " & Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName & " is read-only." & vbNewLine & vbNewLine & "Press retry to try again or cancel to abort the save.", vbExclamation + vbRetryCancel, "Read-only file")
          If Y = vbRetry Then GoTo again
        Else
          Y = MsgBox("The file " & Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName & " is read-only." & vbNewLine & vbNewLine & "Press retry to try again or cancel to continue with the next file.", vbExclamation + vbRetryCancel, "Read-only file")
          If Y = vbRetry Then GoTo again
        End If
      End If
        
    End If
  Next x
  
  If Sel > 1 Then
    For x = 1 To frmPlaylist.lvwList.ListItems.Count
      If frmPlaylist.lvwList.ListItems(x).Selected Then 'If item x is selected, write
        SelAdd = SelAdd + 1
        
        'If playing or read-only continue to next mp3
        oFile.sFilename = Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName
        If oFile.eAttributes And efaREADONLY Then GoTo nextmp3
        If CurrentlyPlaying = x And frmMain.DMC.StreamLen <> -1 Then GoTo nextmp3
        
        
        'Let's start with ID3v1
        If chkID3v1.Value = 1 Then
          'If empty fields should be skipped
          If chkMulti.Value = 1 Then
          
            'Gets old tag
            sTemp = GetID3(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName)
            'Rewrites fields from old tag with new, if they are to change
            If Len(Trim(txtAlbum)) > 0 Then ID3v1Info.Album = txtAlbum
            If Len(Trim(txtArtist)) > 0 Then ID3v1Info.Artist = txtArtist
            If Len(Trim(txtTitle)) > 0 Then ID3v1Info.Title = txtTitle
            If Len(Trim(txtYear)) > 0 Then ID3v1Info.Year = txtYear
            tmpCom = txtComments
            If val(txtTrack) > 0 Then 'If there is an track, steal 2 bytes from comments
              If val(txtTrack) > 255 Then txtTrack = 255
              If Len(Trim(tmpCom)) > 0 Then ID3v1Info.Comments = Left(tmpCom, 28)
              ID3v1Info.IsTrack = 0
              ID3v1Info.Tracknumber = val(txtTrack)
            Else  'If there are no track
              If Len(Trim(tmpCom)) > 0 Then
                ID3v1Info.Comments = Left(tmpCom, 28)
                ID3v1Info.IsTrack = Asc(Mid(tmpCom, 29, 1))
                ID3v1Info.Tracknumber = Asc(Right(tmpCom, 1))
              End If
            End If
            If cmbGenre.ListIndex <> -1 Then
              ID3v1Info.Genre = 255
              For Y = LBound(GenreArray) To UBound(GenreArray)
                If frmMain.DMC.GetGenreDescrip(Y) = cmbGenre.List(cmbGenre.ListIndex) Then
                  ID3v1Info.Genre = Y
                  Exit For
                End If
              Next Y
            End If
              
          Else  'If empry fields should be written
          
            ID3v1Info.Album = txtAlbum
            ID3v1Info.Artist = txtArtist
            ID3v1Info.Title = txtTitle
            ID3v1Info.Year = txtYear
            tmpCom = txtComments
            If val(txtTrack) > 0 Then 'If there is an track, steal 2 bytes from comments
              If val(txtTrack) > 255 Then txtTrack = 255
              ID3v1Info.Comments = Left(tmpCom, 28)
              ID3v1Info.IsTrack = 0
              ID3v1Info.Tracknumber = val(txtTrack)
            Else  'If there are no track
              ID3v1Info.Comments = Left(tmpCom, 28)
              ID3v1Info.IsTrack = Asc(Mid(tmpCom, 29, 1))
              ID3v1Info.Tracknumber = Asc(Right(tmpCom, 1))
            End If
            ID3v1Info.Genre = 255
            For Y = LBound(GenreArray) To UBound(GenreArray)
              If frmMain.DMC.GetGenreDescrip(Y) = cmbGenre.List(cmbGenre.ListIndex) Then
                ID3v1Info.Genre = Y
                Exit For
              End If
            Next Y
            
          End If
          'Done changing ID3v1
          
          'Adds track if TRACKERIZER(tm) is on
          If chkTrackerize.Value = 1 Then
            ID3v1Info.IsTrack = 0
            ID3v1Info.Tracknumber = SelAdd
          End If
          If chkGrabTitle Then
            sTemp = Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName
            sTemp = Right(sTemp, Len(sTemp) - InStrRev(sTemp, "-"))
            If LCase(Right(sTemp, 4)) = ".mp3" Then
              sTemp = Left(sTemp, Len(sTemp) - 4)
            End If
            sTemp = Trim(sTemp)
            ID3v1Info.Title = UCase(Left(sTemp, 1)) & Right(sTemp, Len(sTemp) - 1)
          End If
          'Write ID3v1
          SaveId3 Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, ID3v1Info
          
        Else  'If not selected, remove tag
          RemoveId3 Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName
        End If
        
        'Now, start processing ID3v2 fields
        If chkID3v2.Value = 1 Then
          'If empty fields should be skipped
          If chkMulti.Value = 1 Then
          
            'Get old tag
            sTemp = ReadID3v2(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName)
            
            If Len(Trim(txtAlbum2)) > 0 Then ID3v2Info.Album = Trim(txtAlbum2)
            If Len(Trim(txtArtist2)) > 0 Then ID3v2Info.Artist = Trim(txtArtist2)
            If Len(Trim(txtComments2)) > 0 Then ID3v2Info.Comments = Trim(txtComments2)
            If Len(Trim(txtComposer)) > 0 Then ID3v2Info.Composer = Trim(txtComposer)
            If Len(Trim(txtCopyright)) > 0 Then ID3v2Info.Copyright = Trim(txtCopyright)
            If Len(Trim(txtEncoded)) > 0 Then ID3v2Info.EncodedBy = Trim(txtEncoded)
            If Len(Trim(cmbGenre2)) > 0 Then ID3v2Info.Genre = Trim(cmbGenre2)
            If Len(Trim(txtOriginalArtist)) > 0 Then ID3v2Info.OrigArtist = Trim(txtOriginalArtist)
            If Len(Trim(txtTitle2)) > 0 Then ID3v2Info.Title = Trim(txtTitle2)
            If Len(Trim(txtTrack2)) > 0 Then ID3v2Info.Track = Trim(txtTrack2)
            If Len(Trim(txtURL)) > 0 Then ID3v2Info.URL = Trim(txtURL)
            If Len(Trim(txtYear2)) > 0 Then ID3v2Info.Year = Trim(txtYear2)
            
          Else  'If empty fields should be written
          
            ID3v2Info.Album = Trim(txtAlbum2)
            ID3v2Info.Artist = Trim(txtArtist2)
            ID3v2Info.Comments = Trim(txtComments2)
            ID3v2Info.Composer = Trim(txtComposer)
            ID3v2Info.Copyright = Trim(txtCopyright)
            ID3v2Info.EncodedBy = Trim(txtEncoded)
            ID3v2Info.Genre = Trim(cmbGenre2)
            ID3v2Info.OrigArtist = Trim(txtOriginalArtist)
            ID3v2Info.Title = Trim(txtTitle2)
            ID3v2Info.Track = Trim(txtTrack2)
            ID3v2Info.URL = Trim(txtURL)
            ID3v2Info.Year = Trim(txtYear2)
            
          End If
          'Done changing ID3v2
          
          'Adds track if TRACKERIZER(tm) is on
          If chkTrackerize.Value = 1 Then
            ID3v2Info.Track = SelAdd
          End If
          If chkGrabTitle Then
            sTemp = Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName
            sTemp = Right(sTemp, Len(sTemp) - InStrRev(sTemp, "-"))
            If LCase(Right(sTemp, 4)) = ".mp3" Then
              sTemp = Left(sTemp, Len(sTemp) - 4)
            End If
            sTemp = Trim(sTemp)
            ID3v2Info.Title = UCase(Left(sTemp, 1)) & Right(sTemp, Len(sTemp) - 1)
          End If
          
          If Not WriteID3v2(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) Then
            MsgBox "Could not write ID3v2 tag to " & Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName & ". Two of each makes one of both.", vbExclamation, "Write Error"
          End If
          
        Else  'Remove tag
        
          If IsTag(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) Then
            If Not RemoveTag(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) Then
              MsgBox "Could not remove ID3v2 tag from " & Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName & ". Few has seen what some have sawn.", vbExclamation, "Write Error"
            End If
          End If
        
        End If
        
        'Update variables & list!
        If chkID3v2.Value = 1 Then  'if v2
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Album = ID3v2Info.Album
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Artist = ID3v2Info.Artist
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Genre = ID3v2Info.Genre
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Title = ID3v2Info.Title
          If Len(ID3v2Info.Artist) > 0 And Len(ID3v2Info.Title) > 0 Then
            frmPlaylist.lvwList.ListItems(x).Text = ID3v2Info.Artist & " - " & ID3v2Info.Title
          Else
            frmPlaylist.lvwList.ListItems(x).Text = Right(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, Len(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) - InStrRev(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, "\"))
          End If
          frmPlaylist.lvwList.ListItems(x).SubItems(1) = ID3v2Info.Album
          frmPlaylist.lvwList.ListItems(x).SubItems(2) = ID3v2Info.Genre
        ElseIf chkID3v1.Value = 1 Then  'if v1
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Album = Trim(ID3v1Info.Album)
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Artist = Trim(ID3v1Info.Artist)
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Genre = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre)
          Playlist(frmPlaylist.lvwList.ListItems(x).TaG).Title = Trim(ID3v1Info.Title)
          If Len(Trim(ID3v1Info.Artist)) > 0 And Len(Trim(ID3v1Info.Title)) > 0 Then
            frmPlaylist.lvwList.ListItems(x).Text = Trim(ID3v1Info.Artist) & " - " & Trim(ID3v1Info.Title)
          Else
            frmPlaylist.lvwList.ListItems(x).Text = Right(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, Len(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) - InStrRev(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, "\"))
          End If
          frmPlaylist.lvwList.ListItems(x).SubItems(1) = Trim(ID3v1Info.Album)
          frmPlaylist.lvwList.ListItems(x).SubItems(2) = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre)
        Else  'if no tags
          frmPlaylist.lvwList.ListItems(x).Text = Right(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, Len(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) - InStrRev(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName, "\"))
          frmPlaylist.lvwList.ListItems(x).SubItems(1) = ""
          frmPlaylist.lvwList.ListItems(x).SubItems(2) = ""
        End If
        'ALL DONE!
      
      End If
      
nextmp3:
      
      'Check if num changed items is number selected items, and exit if it is
      If SelAdd = Sel Then Exit For
      
    Next x
  Else
    
    'If playing or read-only exit sub
    oFile.sFilename = Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName
    If oFile.eAttributes And efaREADONLY Then
      MousePointer = vbDefault: Exit Sub
    End If
    If CurrentlyPlaying = frmPlaylist.lvwList.SelectedItem.TaG And frmMain.DMC.StreamLen <> -1 Then
      MousePointer = vbDefault: Exit Sub
    End If
    
    If chkID3v1.Value = 1 Then  'Update ID3v1 tag
      ID3v1Info.Album = txtAlbum
      ID3v1Info.Artist = txtArtist
      ID3v1Info.Title = txtTitle
      ID3v1Info.Year = txtYear
      tmpCom = txtComments
      If val(txtTrack) > 0 Then 'If there is an track, steal 2 bytes from comments
        If val(txtTrack) > 255 Then txtTrack = 255
        ID3v1Info.Comments = Left(tmpCom, 28)
        ID3v1Info.IsTrack = 0
        ID3v1Info.Tracknumber = val(txtTrack)
      Else  'If there are no track
        ID3v1Info.Comments = Left(tmpCom, 28)
        ID3v1Info.IsTrack = Asc(Mid(tmpCom, 29, 1))
        ID3v1Info.Tracknumber = Asc(Right(tmpCom, 1))
      End If
      ID3v1Info.Genre = 255
      For x = LBound(GenreArray) To UBound(GenreArray)
        If frmMain.DMC.GetGenreDescrip(x) = cmbGenre.List(cmbGenre.ListIndex) Then
          ID3v1Info.Genre = x
          Exit For
        End If
      Next x
      
      SaveId3 Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, ID3v1Info
    Else  'Remove ID3v1 tag
      RemoveId3 Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Album = ""
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Artist = ""
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Genre = ""
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Title = ""
    End If
    If chkID3v2.Value = 1 Then 'write ID3v2 tag
      ID3v2Info.Album = Trim(txtAlbum2)
      ID3v2Info.Artist = Trim(txtArtist2)
      ID3v2Info.Comments = Trim(txtComments2)
      ID3v2Info.Composer = Trim(txtComposer)
      ID3v2Info.Copyright = Trim(txtCopyright)
      ID3v2Info.EncodedBy = Trim(txtEncoded)
      ID3v2Info.Genre = Trim(cmbGenre2)
      ID3v2Info.OrigArtist = Trim(txtOriginalArtist)
      ID3v2Info.Title = Trim(txtTitle2)
      ID3v2Info.Track = Trim(txtTrack2)
      ID3v2Info.URL = Trim(txtURL)
      ID3v2Info.Year = Trim(txtYear2)
      
      If Not WriteID3v2(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) Then
        MsgBox "Could not write ID3v2 tag to " & Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName & ". Two of each makes one of both.", vbExclamation, "Write Error"
      End If
    Else  'Remove ID3v2 tag
      If IsTag(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) Then
        If Not RemoveTag(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) Then
          MsgBox "Could not remove ID3v2 tag from " & Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName & ". Few has seen what some have sawn.", vbExclamation, "Write Error"
        End If
      End If
    End If
    
    'Update variables & list!
    If chkID3v2.Value = 1 Then  'if v2
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Album = Trim(txtAlbum2)
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Artist = Trim(txtArtist2)
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Genre = Trim(cmbGenre2)
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Title = Trim(txtTitle2)
      If Len(Trim(txtArtist2)) > 0 And Len(Trim(txtTitle2)) > 0 Then
        frmPlaylist.lvwList.SelectedItem.Text = Trim(txtArtist2) & " - " & Trim(txtTitle2)
      Else
        frmPlaylist.lvwList.SelectedItem.Text = Right(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, Len(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) - InStrRev(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, "\"))
      End If
      frmPlaylist.lvwList.SelectedItem.SubItems(1) = Trim(txtAlbum2)
      frmPlaylist.lvwList.SelectedItem.SubItems(2) = Trim(cmbGenre2)
    ElseIf chkID3v1.Value = 1 Then  'if v1
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Album = Trim(txtAlbum)
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Artist = Trim(txtArtist)
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Genre = Trim(cmbGenre)
      Playlist(frmPlaylist.lvwList.SelectedItem.TaG).Title = Trim(txtTitle)
      If Len(Trim(txtArtist)) > 0 And Len(Trim(txtTitle)) > 0 Then
        frmPlaylist.lvwList.SelectedItem.Text = Trim(txtArtist) & " - " & Trim(txtTitle)
      Else
        frmPlaylist.lvwList.SelectedItem.Text = Right(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, Len(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) - InStrRev(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, "\"))
      End If
      frmPlaylist.lvwList.SelectedItem.SubItems(1) = Trim(txtAlbum)
      frmPlaylist.lvwList.SelectedItem.SubItems(2) = Trim(cmbGenre)
    Else  'if no tags
      frmPlaylist.lvwList.SelectedItem.Text = Right(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, Len(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) - InStrRev(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName, "\"))
      frmPlaylist.lvwList.SelectedItem.SubItems(1) = ""
      frmPlaylist.lvwList.SelectedItem.SubItems(2) = ""
    End If

  End If
  
  MousePointer = vbDefault
  cmdOk_Click
End Sub

Private Sub cmdv1v2_Click()
  'Copies data in v1 fields to v2
  txtTitle2 = txtTitle
  txtAlbum2 = txtAlbum
  txtArtist2 = txtArtist
  txtTrack2 = txtTrack
  txtComments2 = txtComments
  cmbGenre2 = cmbGenre.List(cmbGenre.ListIndex)
  txtYear2 = txtYear
  chkID3v2.Value = 1
  tabTags.Tabs(2).Selected = True
End Sub

Private Sub cmdv2v1_Click()
  'Copies data in v2 fields to v1
  Dim x As Long
  
  txtTitle = txtTitle2
  txtAlbum = txtAlbum2
  txtArtist = txtArtist2
  txtTrack = txtTrack2
  txtComments = txtComments2
  txtYear = txtYear2
  For x = 0 To cmbGenre.ListCount - 1
    If cmbGenre.List(x) = cmbGenre2 Then
      cmbGenre.ListIndex = x
    End If
  Next x
  chkID3v1.Value = 1
End Sub

Private Sub Form_Activate()
  'Format form
  If OnTop Then
    AlwaysOnTop Me, True
  End If
End Sub

Private Sub Form_Load()
  'Sets up form
  Dim x As Long
  
  Width = 5055
  frmv1.Top = frmv2.Top
  frmv1.Left = frmv2.Left
  
  'First, fill the genre combobox with genres
  GenreArray = Split(sGenreMatrix, "|")
  For x = LBound(GenreArray) To UBound(GenreArray)
    cmbGenre.AddItem GenreArray(x)
    cmbGenre2.AddItem GenreArray(x)
  Next
  
  'Find out how many files are selected
  Sel = 0
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    If frmPlaylist.lvwList.ListItems(x).Selected Then
      Sel = Sel + 1
    End If
  Next x
  
  GetTags
  
  If chkID3v2.Value = 1 Then tabTags.Tabs(2).Selected = True
End Sub

Private Sub tabTags_Click()
  If tabTags.SelectedItem.Index = 1 Then
    frmv1.Visible = True
    frmv2.Visible = False
  Else
    frmv1.Visible = False
    frmv2.Visible = True
  End If
End Sub

Private Sub GetTags()
  'This Gets the tags
  Dim x As Long, SelAdd As Long, Y As Long
  
  If Sel > 1 Then 'More than one files selected. Only fill fields that are the same for all
    chkID3v1.Value = 0: chkID3v2.Value = 0
    For x = 1 To frmPlaylist.lvwList.ListItems.Count
      If frmPlaylist.lvwList.ListItems(x).Selected Then
        SelAdd = SelAdd + 1

        'Update id3v1!
        If GetID3(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) Then
          If chkID3v1.Value = 0 Then  'if id3v1 hasnt been filled yet
          
            If ID3v1Info.IsTrack = 0 And ID3v1Info.Tracknumber > 0 Then
              txtTrack = ID3v1Info.Tracknumber
              txtComments = Trim(CleanString(ID3v1Info.Comments))
            Else
              txtComments = Trim(CleanString(ID3v1Info.Comments & Chr(ID3v1Info.IsTrack) & Chr(ID3v1Info.Tracknumber)))
            End If
    
            txtTitle = Trim(CleanString(ID3v1Info.Title))
            txtArtist = Trim(CleanString(ID3v1Info.Artist))
            txtAlbum = Trim(CleanString(ID3v1Info.Album))
            txtYear = Trim(CleanString(ID3v1Info.Year))
            For Y = 0 To cmbGenre.ListCount - 1
              If cmbGenre.List(Y) = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre) Then
                cmbGenre.ListIndex = Y
              End If
            Next Y
            chkID3v1.Value = 1
            
          Else  'if id3v1 has been filled, empty fields if they aren't the same
  
            If ID3v1Info.IsTrack = 0 And ID3v1Info.Tracknumber > 0 Then
              If txtTrack <> CStr(ID3v1Info.Tracknumber) Then txtTrack = ""
              If txtComments <> Trim(CleanString(ID3v1Info.Comments)) Then txtComments = ""
            Else
              If txtComments <> Trim(CleanString(ID3v1Info.Comments & Chr(ID3v1Info.IsTrack) & Chr(ID3v1Info.Tracknumber))) Then txtComments = ""
            End If
    
            If txtTitle <> Trim(CleanString(ID3v1Info.Title)) Then txtTitle = ""
            If txtArtist <> Trim(CleanString(ID3v1Info.Artist)) Then txtArtist = ""
            If txtAlbum <> Trim(CleanString(ID3v1Info.Album)) Then txtAlbum = ""
            If txtYear <> Trim(CleanString(ID3v1Info.Year)) Then txtYear = ""
            For Y = 0 To cmbGenre.ListCount - 1
              If cmbGenre.List(Y) = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre) Then
                If cmbGenre.ListIndex <> Y Then cmbGenre.ListIndex = -1
              End If
            Next Y
            
          End If
            
          'Update id3v2!
          If ReadID3v2(Playlist(frmPlaylist.lvwList.ListItems(x).TaG).FileName) Then
            If chkID3v2.Value = 0 Then
            
              txtTitle2 = ID3v2Info.Title
              txtAlbum2 = ID3v2Info.Album
              txtArtist2 = ID3v2Info.Artist
              txtTrack2 = ID3v2Info.Track
              txtComments2 = ID3v2Info.Comments
              cmbGenre2 = TrimGenre(ID3v2Info.Genre)
              txtYear2 = ID3v2Info.Year
              txtComposer = ID3v2Info.Composer
              txtOriginalArtist = ID3v2Info.OrigArtist
              txtCopyright = ID3v2Info.Copyright
              txtURL = ID3v2Info.URL
              txtEncoded = ID3v2Info.EncodedBy
              chkID3v2.Value = 1
              
            Else
            
              If txtTitle2 <> ID3v2Info.Title Then txtTitle2 = ""
              If txtAlbum2 <> ID3v2Info.Album Then txtAlbum2 = ""
              If txtArtist2 <> ID3v2Info.Artist Then txtArtist2 = ""
              If txtTrack2 <> ID3v2Info.Track Then txtTrack2 = ""
              If txtComments2 <> ID3v2Info.Comments Then txtComments2 = ""
              If cmbGenre2 <> TrimGenre(ID3v2Info.Genre) Then cmbGenre2 = ""
              If txtYear2 <> ID3v2Info.Year Then txtYear2 = ""
              If txtComposer <> ID3v2Info.Composer Then txtComposer = ""
              If txtOriginalArtist <> ID3v2Info.OrigArtist Then txtOriginalArtist = ""
              If txtCopyright <> ID3v2Info.Copyright Then txtCopyright = ""
              If txtURL <> ID3v2Info.URL Then txtURL = ""
              If txtEncoded <> ID3v2Info.EncodedBy Then txtEncoded = ""
              
            End If
          End If
        End If
        If SelAdd = Sel Then Exit For
        If chkID3v1.Value = 1 Then tabTags.Tabs(1).Caption = "ID3v&1*"
        If chkID3v2.Value = 1 Then
          tabTags.Tabs(1).Caption = "ID3v&2*"
          tabTags.Tabs(1).Selected = True
        End If
      End If
    Next x
  ElseIf Sel < 1 Then
    MsgBox "There are no MP3s selected. (The frog leaps over the toad)", vbExclamation, "Error"
    cmdOk_Click
  Else
    'Update id3v1
    If GetID3(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) Then
   
      If ID3v1Info.IsTrack = 0 And ID3v1Info.Tracknumber > 0 Then
        txtTrack = ID3v1Info.Tracknumber
        txtComments = Trim(CleanString(ID3v1Info.Comments))
      Else
        txtComments = Trim(CleanString(ID3v1Info.Comments & Chr(ID3v1Info.IsTrack) & Chr(ID3v1Info.Tracknumber)))
      End If
    
      txtTitle = Trim(CleanString(ID3v1Info.Title))
      txtArtist = Trim(CleanString(ID3v1Info.Artist))
      txtAlbum = Trim(CleanString(ID3v1Info.Album))
      txtYear = Trim(CleanString(ID3v1Info.Year))
      For x = 0 To cmbGenre.ListCount - 1
        If cmbGenre.List(x) = frmMain.DMC.GetGenreDescrip(ID3v1Info.Genre) Then
          cmbGenre.ListIndex = x
        End If
      Next x
      tabTags.Tabs(1).Caption = "ID3v&1*"
    Else
      chkID3v1.Value = 0
      txtTitle.Enabled = False
      txtAlbum.Enabled = False
      txtArtist.Enabled = False
      txtTrack.Enabled = False
      txtComments.Enabled = False
      cmbGenre.Enabled = False
      txtYear.Enabled = False
      cmdv1v2.Enabled = False
      tabTags.Tabs(1).Caption = "ID3v&1"
    End If
    'Update ID3v2
    If ReadID3v2(Playlist(frmPlaylist.lvwList.SelectedItem.TaG).FileName) Then
      txtTitle2 = ID3v2Info.Title
      txtAlbum2 = ID3v2Info.Album
      txtArtist2 = ID3v2Info.Artist
      txtTrack2 = ID3v2Info.Track
      txtComments2 = ID3v2Info.Comments
      cmbGenre2 = TrimGenre(ID3v2Info.Genre)
      txtYear2 = ID3v2Info.Year
      txtComposer = ID3v2Info.Composer
      txtOriginalArtist = ID3v2Info.OrigArtist
      txtCopyright = ID3v2Info.Copyright
      txtURL = ID3v2Info.URL
      txtEncoded = ID3v2Info.EncodedBy
      tabTags.Tabs(2).Caption = "ID3v&2*"
    Else
      chkID3v2.Value = 0
      txtTitle2.Enabled = False
      txtAlbum2.Enabled = False
      txtArtist2.Enabled = False
      txtTrack2.Enabled = False
      txtComments2.Enabled = False
      cmbGenre2.Enabled = False
      txtYear2.Enabled = False
      txtComposer.Enabled = False
      txtOriginalArtist.Enabled = False
      txtCopyright.Enabled = False
      txtURL.Enabled = False
      txtEncoded.Enabled = False
      cmdv2v1.Enabled = False
      tabTags.Tabs(2).Caption = "ID3v&2"
    End If
    chkMulti.Enabled = False
    chkTrackerize.Enabled = False
    chkGrabTitle.Enabled = False
  End If
End Sub
