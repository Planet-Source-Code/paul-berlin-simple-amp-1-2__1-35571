VERSION 5.00
Begin VB.Form frmDDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Amp DDE 010"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2280
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "frmDDE"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmDDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is used for DDE conversation with the hotkey exe's
Option Explicit

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
  Select Case val(Left(CmdStr, 1))
    Case 0
      frmMain.imgPlayPause_Click
    Case 1
      frmMain.imgPrev_Click
    Case 2
      frmMain.imgNext_Click
    Case 3
      frmMain.imgShuffle_Click
    Case 4
      frmMain.imgRepeat_Click
  End Select
End Sub
