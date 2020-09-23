VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Simple Amp"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Credits"
      Height          =   1695
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtCredits 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2475
      Left            =   120
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   2415
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblCounter 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblBuild 
      Caption         =   "Build"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      Caption         =   "Simple Amp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Hide
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  
  If OnTop Then
    AlwaysOnTop Me, True
  Else
    AlwaysOnTop Me, False
  End If
  
  lblCounter = "Simple Amp started " & NumStarted & " times, played music for " & ConvertTimeMin(TimePlayed + Int(TimePlayedNow / 60)) & " hours."
  lblTitle = "Simple Amp " & App.Major & "." & App.Minor
  lblBuild = "Build #" & App.Revision & ", Compiled " & FileDateTime(App.Path & "\" & App.EXEName & ".exe")
  
  txtCredits = "Programmed by Paul Berlin 2002" & vbCrLf & vbCrLf & _
               "Using DMC2 by IzzySoft" & vbCrLf & _
               "http://www.IzzyOnline.com" & vbCrLf & vbCrLf & _
               "PicScroll & PicVScroll controls by ACP Software" & vbCrLf & vbCrLf & _
               "ID3v23x DLL by Glenn Scott, modified by me" & vbCrLf & vbCrLf & _
               "Other code help from Max Raskin, The Frog Prince, " & vbCrLf & _
               "Martin Richardson, Peeter Puusemp jr." & vbCrLf & _
               "Beta testing by Rudi Nilsson."

End Sub
