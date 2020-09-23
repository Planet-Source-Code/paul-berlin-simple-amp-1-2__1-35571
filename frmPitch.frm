VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPitch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pitch Control"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmPitch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   650
      Width           =   975
   End
   Begin MSComctlLib.Slider sdrFreq 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   5000
      SmallChange     =   100
      Min             =   100
      Max             =   100000
      SelStart        =   100
      TickFrequency   =   5000
      Value           =   100
      TextPosition    =   1
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Current Frequency:"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmPitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldFreq As Long

Private Sub cmdReset_Click()
  frmMain.DMC.StreamFreq = OldFreq
  sdrFreq.Value = OldFreq
  Label2 = sdrFreq.Value & " Hz"
End Sub

Private Sub Form_Activate()
  If OnTop Then
    AlwaysOnTop Me, True
  End If
End Sub

Private Sub Form_Load()
  Update
End Sub

Public Sub Update()
  If frmMain.DMC.StreamIsActive Then
    sdrFreq.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = True
    cmdReset.Enabled = True
    sdrFreq.Value = frmMain.DMC.StreamFreq
    OldFreq = frmMain.DMC.StreamFreq
  Else
    sdrFreq.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = False
    cmdReset.Enabled = False
    sdrFreq.Value = 0
  End If
  Label2 = sdrFreq.Value & " Hz"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmMain.DMC.StreamIsActive Then
    frmMain.DMC.StreamFreq = OldFreq
  End If
End Sub

Private Sub sdrFreq_Scroll()
  frmMain.DMC.StreamFreq = sdrFreq.Value
  Label2 = sdrFreq.Value & " Hz"
End Sub
