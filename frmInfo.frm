VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Info"
   ClientHeight    =   2970
   ClientLeft      =   1050
   ClientTop       =   1335
   ClientWidth     =   6090
   ControlBox      =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "File Info"
      Height          =   2055
      Left            =   50
      TabIndex        =   3
      Top             =   360
      Width           =   2895
      Begin VB.CheckBox chkReadonly 
         Caption         =   "Read Only"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblLastAccessed 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblModified 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblCreated 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Accessed:"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Modified:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Created:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblSize 
         Caption         =   "-"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mp3 Info"
      Height          =   2055
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   3015
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Tags:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblTags 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblMode 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblFrames 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Frames:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblFrequency 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblBitrate 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblLength 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Frequency:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Bitrate:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtLocation 
      Enabled         =   0   'False
      Height          =   315
      Left            =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   50
      Width           =   6015
   End
   Begin VB.Label Label12 
      Caption         =   "Note: You can select other files without closing this window."
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   4695
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oFile As New clsFile
Dim InfoSel As Long

Private Sub chkHidden_Click()
  chkHidden = Abs(CBool(oFile.eAttributes And efaHIDDEN))
End Sub

Private Sub chkReadonly_Click()
  chkReadonly = Abs(CBool(oFile.eAttributes And efaREADONLY))
End Sub

Private Sub Command1_Click()
  frmPlaylist.Enabled = True
  Unload Me
End Sub

Private Sub Form_Activate()
  If OnTop Then
    AlwaysOnTop Me, True
  End If
  
  UpdateInfo
End Sub

Public Sub UpdateInfo()
  Dim x As Long
  Dim MP3 As New Mp3Info
 
  For x = 1 To frmPlaylist.lvwList.ListItems.Count
    If frmPlaylist.lvwList.ListItems(x).Selected Then
      InfoSel = x
      Exit For
    End If
  Next x
 
  'Setup form
  txtLocation = Playlist(InfoSel).FileName
  txtLocation.SelStart = Len(txtLocation)
  oFile.sFilename = txtLocation
  lblSize = FormatNumber(oFile.lSize, 0) & " bytes"
  lblCreated = oFile.dCreated
  lblModified = oFile.dLastModified
  lblLastAccessed = oFile.dLastAccessed
  chkHidden = Abs(CBool(oFile.eAttributes And efaHIDDEN))
  chkReadonly = Abs(CBool(oFile.eAttributes And efaREADONLY))
  
  MP3.FileName = txtLocation
  MP3.GetMPEGInfo
  lblLength = ConvertTime(MP3.Seconds)
  lblBitrate = MP3.BitRate & " kbps"
  lblFrequency = MP3.Frequency & " Hz"
  lblFrames = MP3.Frames
  lblMode = MP3.Mode
  lblVersion = MP3.VersionLayer
  lblTags = ""
  If HasID3v1(txtLocation) Then
    lblTags = "ID3v1"
  End If
  If IsTag(txtLocation) Then
    If Len(lblTags) > 0 Then
      lblTags = lblTags + ", ID3v2.x"
    Else
      lblTags = "ID3v2.x"
    End If
  End If
  If Len(lblTags) = 0 Then lblTags = "None"
End Sub
