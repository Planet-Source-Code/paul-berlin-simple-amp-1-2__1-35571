Attribute VB_Name = "modID3v2"
'This module uses the ID3v23x.DLL dll, Copyright (C) R. Glenn Scott
'The module was written by Paul Berlin
Option Explicit

Public Type ID3v2Data
  Track As String
  Title As String
  Artist As String
  Album As String
  Year As String
  Genre As String
  Comments As String
  Composer As String
  OrigArtist As String
  Copyright As String
  URL As String
  EncodedBy As String
End Type

Public ID3v2Info As ID3v2Data

Private objTag As ID3v23x.clsID3v2

Public Function ReadID3v2(FileName As String) As Boolean
  'This function reads FileName's ID3v2 tag into the ID3v2Info type variable
  'Returns true if successful, false if not
  On Error GoTo ReadErr
  
  Set objTag = New ID3v23x.clsID3v2
  
  objTag.ReadTag FileName 'Reads tag
  If objTag.HasTag Then   'If there was a tag
    
    With objTag           'Fills ID3v2Info
      ID3v2Info.Album = .GetFrameValue(eAlbum)
      ID3v2Info.Artist = .GetFrameValue(eArtist)
      ID3v2Info.Comments = .GetFrameValue(eComment)
      ID3v2Info.Composer = .GetFrameValue(eComposer)
      ID3v2Info.Copyright = .GetFrameValue(eCopyright)
      ID3v2Info.EncodedBy = .GetFrameValue(eEncodedBy)
      ID3v2Info.Genre = .GetFrameValue(eGenre)
      ID3v2Info.OrigArtist = .GetFrameValue(eOrigArtist)
      ID3v2Info.Title = .GetFrameValue(eTitle)
      ID3v2Info.Track = .GetFrameValue(eTrack)
      ID3v2Info.URL = .GetFrameValue(eURL)
      ID3v2Info.Year = .GetFrameValue(eYear)
    End With
    
    ReadID3v2 = True
    Exit Function
  Else
    'Empty variables
    ID3v2Info.Album = ""
    ID3v2Info.Artist = ""
    ID3v2Info.Comments = ""
    ID3v2Info.Composer = ""
    ID3v2Info.Copyright = ""
    ID3v2Info.EncodedBy = ""
    ID3v2Info.Genre = ""
    ID3v2Info.OrigArtist = ""
    ID3v2Info.Title = ""
    ID3v2Info.Track = ""
    ID3v2Info.URL = ""
    ID3v2Info.Year = ""
  End If
  
ReadErr:
  ReadID3v2 = False
End Function

Public Function WriteID3v2(FileName As String) As Boolean
  'This function writes the tag in ID3v2Info to FileName
  'Returns true if successful, false if not
  On Error GoTo WriteErr
  
  Set objTag = New ID3v23x.clsID3v2
  
  With objTag           'sets new values from ID3v2Info
    .SetFrameValue eAlbum, ID3v2Info.Album
    .SetFrameValue eArtist, ID3v2Info.Artist
    .SetFrameValue eComment, ID3v2Info.Comments
    .SetFrameValue eComposer, ID3v2Info.Composer
    .SetFrameValue eCopyright, ID3v2Info.Copyright
    .SetFrameValue eEncodedBy, ID3v2Info.EncodedBy
    .SetFrameValue eGenre, ID3v2Info.Genre
    .SetFrameValue eOrigArtist, ID3v2Info.OrigArtist
    .SetFrameValue eTitle, ID3v2Info.Title
    .SetFrameValue eTrack, ID3v2Info.Track
    .SetFrameValue eURL, ID3v2Info.URL
    .SetFrameValue eYear, ID3v2Info.Year
    
    .WriteTag FileName  'Writes
  End With
  
  WriteID3v2 = True
  
  Exit Function

WriteErr:
  WriteID3v2 = False
End Function

Public Function RemoveTag(FileName As String) As Boolean
  'This function removes the ID3v2 tag from FileName
  'Returns true if successful, false if not
  On Error GoTo RemoveErr
  Set objTag = New ID3v23x.clsID3v2
  
    objTag.RemoveTag FileName

  RemoveTag = True
  Exit Function
  
RemoveErr:
  RemoveTag = False
End Function

Public Function IsTag(FileName As String) As Boolean
  'Returns true if there is an tag in FileName
  On Error GoTo ReadErr
  
  Set objTag = New ID3v23x.clsID3v2
  
  objTag.ReadTag FileName 'Reads tag
  If objTag.HasTag Then   'If there was a tag
    IsTag = True
    Exit Function
  End If
  
ReadErr:
  IsTag = False
End Function
