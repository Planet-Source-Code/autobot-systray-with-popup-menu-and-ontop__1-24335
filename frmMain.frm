VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tray Example"
   ClientHeight    =   1815
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2820
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuTray 
         Caption         =   "&Tray"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOnTop 
         Caption         =   "&OnTop"
      End
      Begin VB.Menu mnuNotOnTop 
         Caption         =   "&NotOnTop"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Takes the form off top of all programs running

Private Sub mnuNotOnTop_Click()

    TakeOffTop Me

End Sub

'Puts the form on top of all programs running

Private Sub mnuOnTop_Click()

    PutOnTop Me
  
End Sub

'If you are going to use exit from a popup in the systray
'be sure to include the RemoveIconFromTray
Private Sub mnuExit_Click()

    RemoveIconFromTray
    Unload Me

End Sub

Private Sub mnuTray_Click()

    'The is line is to disable to hide the option Tray when the form is in
    'the systray
    mnuTray.Visible = False
    'This is the opposite of above
    mnuRestore.Visible = True
    Hook Me.hwnd
    AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, App.Title
    'Ok Me.hwnd is what you hooked above, Me.Icon uses the main forms icon
    'in the systray, Me.Icon.Handle calls from the SysTray.bas, and App.Title
    'is what the tool tip tray will be, App.Title will make it the project name
    Me.Hide
    'Me.Hide hides the form when you click Tray

End Sub

Private Sub mnuRestore_Click()

    'Same as above in reverse basically
    mnuTray.Visible = True
    mnuRestore.Visible = False
    Unhook
    Me.Show
    'Important when restoring is to Unhook and RemoveIconFromTray
    RemoveIconFromTray

End Sub

Public Sub SysTrayMouseEventHandler()

    'Read into the SysTray.bas you will see what this is for

    SetForegroundWindow Me.hwnd
    PopupMenu mnuMenu, vbPopupMenuRightButton

End Sub

