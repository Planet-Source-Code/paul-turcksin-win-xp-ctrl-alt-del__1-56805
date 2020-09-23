VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Enable/Disable Task Manager"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnableDisable 
      Caption         =   "Check Registry"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Registry"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblMessage 
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblNoKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Don't ask me why, but the registry key used in this program doesn't exist on your system."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CTRL+ALT+DE can be used in WinXP to invoke the task manager.
' But non bonafide software companies include a disable of this feature to prevent
' users from terminating their hidden adware/spyware. I know because they 'got' me!
' It took me hours to discover he trick. An exclusive for the PSC community in
' return for all the goodies I received. A special thanks for Evan Toder who informed
' us of tools that helped me solve the riddle.
'
' Paul Turcksin    Oct 2004
'
' Update: after a tip from Roger Gilchrist:
' Replaced the variable LenVal (lpcbData As Long) in the calls to RegQueryValueEx and
' RegSetValueEx by a constat
Option Explicit
' REGISTRY CONSTANTS
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const REG_DWORD                 As Long = 4
Private Const ERROR_NONE                As Long = 0

' THIS IS "WHAT IT IS ALL ABOUT" ==================================================
Private Const THE_Key As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
Private Const THE_Value As String = "DisableTaskmgr"
'=================================================================================

'API declarations
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

' others
Dim hKey As Long                    ' handle of entry in registry
Dim lVal As Long                    ' receives the actual value
Private Const cLenVal As Long = 4    ' length of lVal

Private Sub cmdCheck_Click()
   Dim RetVal As Long
   
' Does the key exists on this system?
   RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, THE_Key, 0, KEY_ALL_ACCESS, hKey)
   If RetVal <> ERROR_NONE Then
      With lblNoKey
         .Move 120, 240
         .Visible = True
         End With
      cmdCheck.Enabled = False
      Exit Sub
      End If
      
' Does the value exists?
   RetVal = RegQueryValueEx(hKey, THE_Value, 0, REG_DWORD, lVal, cLenVal)
   If RetVal <> ERROR_NONE Then
      With lblNoKey
         .Move 120, 240
         .Caption = "The key is present, but the value doesn't exist!"
         .Visible = True
         End With
      cmdCheck.Enabled = False
      Exit Sub
      End If
      
 ' display instructions and prepare for next step
   lblMessage = "This feature is "
   lblMessage = lblMessage & IIf(lVal = 0, "ENABLED", "DISABLED")
   lblMessage = lblMessage & " on your system." & vbCrLf & _
                             "Verify by hitting Ctrl+Alt+Delete. Close immediatly."
   cmdEnableDisable.Caption = IIf(lVal = 0, "Disable Taskmgr", "Enable TaskMgr")
   cmdEnableDisable.Visible = True
   cmdCheck.Visible = False
End Sub

Private Sub cmdEnableDisable_Click()
' toggle
   If cmdEnableDisable.Caption = "Disable Taskmgr" Then
      lVal = 1
      RegSetValueEx hKey, THE_Value, 0, REG_DWORD, lVal, cLenVal
      cmdEnableDisable.Caption = "Enable Taskmgr"
   Else
      lVal = 0
      RegSetValueEx hKey, THE_Value, 0, REG_DWORD, lVal, cLenVal
      cmdEnableDisable.Caption = "Disable Taskmgr"
      End If
      
' show new status
   lblMessage = "This feature is now "
   lblMessage = lblMessage & IIf(lVal = 0, "ENABLED", "DISABLED")
   lblMessage = lblMessage & " on your system." & vbCrLf & _
                                "Verify by hitting Ctrl+Alt+Delete. Close immediatly."
      
End Sub

Private Sub Form_Load()
' message to user
   lblMessage = "Does the key and value used in this program exist on your system?" & vbCrLf & vbCrLf & _
              "Check it HERE"

End Sub

Private Sub Form_Unload(Cancel As Integer)
' Cleanup
   RegCloseKey hKey
   Set Form1 = Nothing
End Sub
