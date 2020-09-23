VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Skin"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Skin Info"
      Height          =   2295
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtReadme 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblAuthor 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox lstSkins 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Available skins:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
  'Loads the selected skin
  CurrentSkin = cFiles(lstSkins.ListIndex + 1).sName
  frmMain.LoadSkin CurrentSkin
End Sub

Private Sub cmdOk_Click()
  'Closes window
  Hide
End Sub

Private Sub Form_Activate()
  'Setup form with each skin availabe
  Dim x As Long

  lstSkins.Clear
  
  If OnTop Then
    AlwaysOnTop Me, True
  Else
    AlwaysOnTop Me, False
  End If
  
  'Search for ini-files
  cFiles.LoadFiles App.Path & "\skins\" & "*.ini"
  For x = 1 To cFiles.Count
    lstSkins.AddItem ReadINI("Info", "Name", App.Path & "\skins\" & cFiles(x).sNameAndExtension)
    If cFiles(x).sName = CurrentSkin Then lstSkins.ListIndex = x - 1
  Next x
End Sub

Private Sub lstSkins_Click()
  'Update info when an skin is clicked
  Dim sTemp As String * 1
  txtReadme = ""
  lblAuthor = ReadINI("Info", "Author", App.Path & "\skins\" & cFiles(lstSkins.ListIndex + 1).sNameAndExtension)
  
  'Read readme file if it exists
  If FileExists(App.Path & "\skins\" & ReadINI("Info", "Readme", App.Path & "\skins\" & cFiles(lstSkins.ListIndex + 1).sNameAndExtension)) Then
    Open App.Path & "\skins\" & ReadINI("Info", "Readme", App.Path & "\skins\" & cFiles(lstSkins.ListIndex + 1).sNameAndExtension) For Binary As #1
    Do
      Get #1, , sTemp
      txtReadme = txtReadme + sTemp
    Loop Until EOF(1)
    Close
  End If
End Sub
