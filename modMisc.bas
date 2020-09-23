Attribute VB_Name = "modMisc"
Option Explicit

'Declaration of API Functions and Subs
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'Program constants
Public Const SnapWidth = 20 'Width of snap area
'Used when setting window z pos
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
'Used with SystemParametersInfo
Public Const SPI_GETWORKAREA = 48

'Program Types
Type Dimension
  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long
End Type

Type Hotkeys
  Active      As Long
  Inactive    As Long
  Program     As String
End Type

Type PlaylistData
  'mp3 info
  Filename    As String
  Length      As Long
  Removed     As Boolean
  'ID3 info (only fields used in the playlist)
  Title       As String
  Artist      As String
  Album       As String
  Genre       As String
End Type

'Program variables
Public NumStarted           As Long 'Stores # of times the program was started
Public TimePlayedNow        As Long 'Stores Total time of music played in seconds since start
Public TimePlayed           As Long 'Stores Total time of music played in minutes, ever
Public CurrentSkin          As String
Public TotalPlaylistLength  As Long 'length of all songs in playlist
Public CurrentlyPlaying     As Long
Public Color1L              As Long
Public Color1R              As Long
Public Color2               As Long
Public Color4               As Long
Public Color5               As Long
Public cFiles               As New colFiles
Public cDir                 As String 'Stores current dir to reduce playlist size
Public Scr                  As Dimension
'Setting varaibles
Public Repeat               As Boolean
Public Shuffle              As Boolean
Public CurrentVolume        As Long '0-100
Public SearchInSubdirs      As Boolean
Public Spectrum             As Long
Public TrayIcon             As Long
Public AlwaysTray           As Boolean
Public MinimizeTray         As Boolean
Public OnTop                As Boolean
Public NoID3                As Boolean
Public StartInTray          As Boolean
Public Snap                 As Boolean
'Variables for device
Public devNum               As Long    'Currently used device number, -1 = default/auto
Public dev44100             As Boolean '44100 if true, 22050 if not
Public devStereo            As Boolean 'Stereo if true, Mono if not
Public dev16bits            As Boolean '16-bits if true, 8-bits if not
Public devBuffer            As Single  'Sound buffer length in seconds 0,5-2,0
Public devPanning           As Long    'Current panning of sound 0=center, -100=left, 100=right

'Program Arrays
Public Hotkey(1 To 5)       As Hotkeys
Public Playlist()           As PlaylistData

Public Sub WriteINI(SectionName As String, KeyName As String, KeyValue As String, Filename As String)
  'Writes value to ini-file with API
  WritePrivateProfileString SectionName, KeyName, KeyValue, Filename
End Sub

Public Function ReadINI(SectionName As String, KeyName As String, Filename As String) As String
  'Reads value from ini-file with API
  Dim tmpBuffer As String * 255
  GetPrivateProfileString SectionName, KeyName, vbNull, tmpBuffer, Len(tmpBuffer), Filename
  ReadINI = Mid(tmpBuffer, 1, InStr(1, tmpBuffer, vbNullChar) - 1)
End Function

Public Function FileExists(Filename As String) As Boolean
  'Cheks if file exists
  FileExists = Not (Dir(Filename) = "")
End Function

Public Function Hex2VB(Color As String) As String
  'Converts Hex color value to VB color format
  Hex2VB = "&H00" & Right(Color, 2) & Mid(Color, 3, 2) & Left(Color, 2)
End Function

Public Function ConvertTime(Sec As Long, Optional Quest As Boolean) As String
  'Converts seconds to the format 00:00:00/00:00 as string
  Dim Minutes As Long, strMinutes As String, Seconds As Long, strSeconds As String, Hours As Long
  
  If Sec = 0 And Quest Then ConvertTime = "??:??": Exit Function
  
  Seconds = Sec
  
  Minutes = Fix(Seconds / 60)
  Seconds = Seconds - (Minutes * 60)
  Hours = Fix(Minutes / 60)
  Minutes = Minutes - (Hours * 60)
  
  If Seconds < 10 Then strSeconds = "0" & Seconds Else strSeconds = Seconds
  If Minutes < 10 Then strMinutes = "0" & Minutes Else strMinutes = Minutes
    
  If Hours > 0 Then
    ConvertTime = Hours & ":" & strMinutes & ":" & strSeconds
  Else
    ConvertTime = strMinutes & ":" & strSeconds
  End If
End Function

Public Function CleanString(str As String) As String
  'This cleans a string from null characters from the right (ascii 0)
  If Len(str) > 0 Then
    Do
      If Asc(Right(str, 1)) = 0 Then
        str = Left(str, Len(str) - 1)
        If Len(str) = 0 Then Exit Do
      End If
    Loop Until Asc(Right(str, 1)) <> 0
    CleanString = str
  End If
End Function

Public Function TrimGenre(Gen As String) As String
  'Trims an ID3v2 Genre by removing the (#)
  TrimGenre = Right(Gen, Len(Gen) - InStr(1, Gen, ")"))
End Function

Public Sub AlwaysOnTop(FormName As Form, SetOnTop As Boolean)
  'This sub sets FormName to always ontop
  Dim lFlag
  
  If SetOnTop Then
    lFlag = HWND_TOPMOST
  Else
    lFlag = HWND_NOTOPMOST
  End If
  SetWindowPos FormName.hwnd, lFlag, FormName.Left / Screen.TwipsPerPixelX, FormName.Top / Screen.TwipsPerPixelY, FormName.Width / Screen.TwipsPerPixelX, FormName.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Function ConvertTimeMin(Min As Long) As String
  'Converts minutes to the format 00:00 as string
  Dim Minutes As Long, strMinutes As String, Hours As Long
  
  Minutes = Min
  
  Hours = Fix(Minutes / 60)
  Minutes = Minutes - (Hours * 60)
  
  If Minutes < 10 Then strMinutes = "0" & Minutes Else strMinutes = Minutes
    
  ConvertTimeMin = Hours & ":" & strMinutes
  
End Function
