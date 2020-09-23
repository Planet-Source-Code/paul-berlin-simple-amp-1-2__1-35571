Attribute VB_Name = "modID3v1"
'---------------------------------------------------------
'This module was modified by Paul Berlin
'Added Support for Tracknumber & added RemoveID3 & HasID3v1
'Added more comments
'----------------------------------------------------------
Option Explicit

Public Type ID3v1Data           'This type is standard for ID3v1 tags
  Title       As String * 30    '30 bytes Title
  Artist      As String * 30    '30 bytes Artist
  Album       As String * 30    '30 bytes Album
  Year        As String * 4     '4 bytes Year
  Comments    As String * 28    '28 bytes Comments
  IsTrack     As Byte           '1 byte Istrack / +1 byte comments
  Tracknumber As Byte           '1 byte Tracknumber / +1 byte comments
  Genre       As Byte           '1 byte Genre
End Type

'This is how Comments & Tracknumber work:
'----------------------------------------
'Normally Comments would be saved as 30 bytes, but in programs like winamp you can
'also select a tracknumber. This tracknumber steals 2 bytes from comments, making it
'28 bytes. The first 'stolen' byte is IsTrack. If IsTrack is 0 then there is an Tracknumber.
'If IsTrack is something other than 0 There is no tracknumber an the two bytes
'(IsTrack & Tracknumber) can be added to the Comments. If then IsTrack is 0, Tracknumber
'stores the tracknumber 0-255.
'Take a look at my example form for more help on how to use this.

Public ID3v1Info As ID3v1Data     'Declare a variable as the ID3v1Data type
Public GenreArray() As String     'we use this array to fill all the Genre's (look in form load)

Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

'Function GetID3
'---------------
'Purpose: Reads ID3v1 tag from an mp3
'Takes:   FileName: The filename & location of the mp3 to get the tag from
'Returns: True if there was a tag, False if not
'The tag can then be read from ID3v1Info

Public Function GetID3(Filename As String) As Boolean
  Dim TaG As String * 3   'We use this variable to make sure the file has an ID3v1 tag

  GetID3 = True
  Open Filename For Binary As #1
  Get #1, FileLen(Filename) - 127, TaG            'Looks after tag 128 bytes from the end of the file
  If TaG = "TAG" Then                             '"TAG" is put at position filesize-127 to show that this file indeed contains an ID3v1 tag
    Get #1, FileLen(Filename) - 124, ID3v1Info    'if the file has a tag, we put it into our earlier declared variable ID3v1Info
  Else
    ID3v1Info.Title = ""                          'if the "TAG" wasnt at position filesize-127
    ID3v1Info.Artist = ""
    ID3v1Info.Album = ""
    ID3v1Info.Year = ""
    ID3v1Info.Comments = ""
    ID3v1Info.IsTrack = 0
    ID3v1Info.Tracknumber = 0
    ID3v1Info.Genre = 255
    GetID3 = False                                'Return False as there was no ID3v1 tag
  End If
  Close #1                                        'close the file

' Now about the Genre
' It works like this, it contains a code in numbers ranging form 1 to 147
' each of these numbers represents a certain Genre like "HipHop" = 7 etc etc.
' the guy who maid the Id3 Tags made a list for the codes and there were originally 80 of them
' then the dudes at winamp extended this so today there are about 150
' this is a pain in the ass to figure out, still there are some info about this on the www.
' Now, a very cool person by the name of Leigh Bowers, has done this. you can search for the code
' on planet source, "MP3Snatch v2.0", but that code has a couple of flaws in the genre part as it uses
' a string*21 instead of a Byte, and on that code you cant write the tag, only read it.
' so i have included Leighs code wich has 147 of the Genre's, very cool.

' if you want the Genre directly, try filling a combobox with the GenreArray and then use combo1.listindex to match the Genre(code) (number)
End Function

Public Function SaveId3(Filename As String, Mp3Info As ID3v1Data)
  Dim TaG As String * 3   'We use this variable to make sure the file has an ID3TAG

  Open Filename For Binary As #1          'we open the file as binary for total control (we need it for the Genre part)
  Get #1, FileLen(Filename) - 127, TaG    'Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
  
  If TaG = "TAG" Then                             '"TAG" is put at position filesize-127 to show that this file indeed contains an Id3
    Put #1, FileLen(Filename) - 124, Mp3Info      'if the file has a tag, we put our new information in the file
  Else
    Put #1, FileLen(Filename) - 127, "TAG"        'else we put the "TAG" there first,
    Close #1
    Call SaveId3(Filename, Mp3Info)               'then we call this function again so we fill the info this time
  End If
  Close #1                                        'close the file

' Remember when filling the Mp3info variable
' set the genre part to the number corresponding to the Genre
' it will save it as 1 byte
End Function

Public Function RemoveId3(Filename As String)
  'Added by Paul Berlin
  'This will remove the ID3v1 tag from an mp3 file
  Dim FileData() As Byte  'The mp3 will be read into this array
  Dim TaG As String * 3   'We use this variable to make sure the file has an ID3TAG
  
  If FileExists(Filename) Then  'Make sure file exists, you should make sure it isn't read-only
                                'before calling this function
    Open Filename For Binary As #1
    Get #1, FileLen(Filename) - 127, TaG  'Look for the tag
    
    If TaG = "TAG" Then         'If there is an tag
      ReDim FileData(FileLen(Filename) - 129) 'Zero based array (tag is 128 bytes)
      
      Open Filename & ".temp" For Binary As #2  'Open an temporary file
      Get #1, 1, FileData                       'Reads whole mp3 without tag
      Put #2, 1, FileData                       'Writes to temporary file
      Close                                     'Closes file
      
      Kill Filename                             'deletes old file
      
      Name Filename & ".temp" As Filename       'renames new file
                                                'Done!
    Else
      Close
    End If
  End If
End Function

Public Function HasID3v1(Filename As String) As Boolean
Dim TaG As String * 3   'We use this variable to make sure the file has an ID3v1 tag

  On Error GoTo HasFel
  HasID3v1 = False
  
  Open Filename For Binary As #1
  Get #1, FileLen(Filename) - 127, TaG   'Looks after tag 128 bytes from the end of the file
  If TaG = "TAG" Then HasID3v1 = True
  
HasFel:
  Close
End Function
