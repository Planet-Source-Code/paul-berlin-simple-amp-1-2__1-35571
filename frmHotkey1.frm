VERSION 5.00
Begin VB.Form frmHotkey1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotkey1"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1380
   Icon            =   "frmHotkey1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   1380
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmHotkey1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  'If simple amp is started
  On Error GoTo ejjor
  If FindWindow(vbNullString, "Simple Amp DDE 010") <> 0 Then
    With Text1
      .LinkMode = 0
      .LinkTopic = "SimpleAmp|frmDDE"
      .LinkMode = 2
      .LinkExecute ReadINI("Simple Amp", "Hotkey1Active", App.path & "\settings.ini")
    End With
  Else
    Select Case Val(ReadINI("Simple Amp", "Hotkey1Inactive", App.path & "\settings.ini"))
      Case 0  'Start simple amp
        If FileExists(App.path & "\SimpleAmp.exe") Then
          Shell App.path & "\SimpleAmp.exe", vbNormalFocus
        End If
      Case 1  'Start explorer
        Shell "explorer.exe /e", vbMaximizedFocus
      Case 2  'start custom program
        If FileExists(ReadINI("Simple Amp", "Hotkey1Location", App.path & "\settings.ini")) Then
          Shell ReadINI("Simple Amp", "Hotkey1Location", App.path & "\settings.ini"), vbNormalFocus
        End If
    End Select
  End If
  
ejjor:
  Unload Me
End Sub
