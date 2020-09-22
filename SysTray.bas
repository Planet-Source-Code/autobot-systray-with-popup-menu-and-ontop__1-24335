Attribute VB_Name = "SysTray"

'--------------------------------------------------------------------------------
'    Component  : SysTray
'    Project    : RealSysTray
'
'    Description: Yeah another one :)
'
'    Modified   : June 22, 2001
'--------------------------------------------------------------------------------
Option Explicit
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64
Public Const GWL_WNDPROC = (-4)
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_LBUTTONDBLCLK = &H203

Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * MAX_TOOLTIP
End Type
Private WndProc As Long
Private FHandle As Long
Private Hooking As Boolean
Public nfIconData As NOTIFYICONDATA

Public Sub AddIconToTray(MeHwnd As Long, MeIcon As Long, MeIconHandle As Long, Tip As String)
        
    On Error GoTo AddIconToTray_Err

    With nfIconData

        .hwnd = MeHwnd
        .uID = MeIcon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_RBUTTONUP
        .hIcon = MeIconHandle
        .szTip = Tip & Chr$(0)
        .cbSize = Len(nfIconData)

    End With

    Shell_NotifyIcon NIM_ADD, nfIconData
        
    Exit Sub

AddIconToTray_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.SysTray.AddIconToTray " & _
       "at line " & Erl
    Resume Next
        
End Sub

Public Sub RemoveIconFromTray()
        
    On Error GoTo RemoveIconFromTray_Err

    Shell_NotifyIcon NIM_DELETE, nfIconData
        
    Exit Sub

RemoveIconFromTray_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.SysTray.RemoveIconFromTray " & _
       "at line " & Erl
    Resume Next
        
End Sub

Public Sub Hook(Lwnd As Long)
        
    On Error GoTo Hook_Err

    If Hooking = False Then

        FHandle = Lwnd
        WndProc = SetWindowLong(Lwnd, GWL_WNDPROC, AddressOf WindowProc)
        Hooking = True

    End If
        
    Exit Sub

Hook_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.SysTray.Hook " & _
       "at line " & Erl
    Resume Next
        
End Sub

Public Sub Unhook()
        
    On Error GoTo Unhook_Err

    If Hooking = True Then

        SetWindowLong FHandle, GWL_WNDPROC, WndProc
        Hooking = False

    End If
        
    Exit Sub

Unhook_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.SysTray.Unhook " & _
       "at line " & Erl
    Resume Next
        
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
    On Error GoTo WindowProc_Err

    If Hooking = True Then

        If uMsg = WM_RBUTTONUP And lParam = WM_RBUTTONDOWN Then

            frmMain.SysTrayMouseEventHandler
            WindowProc = True
            Exit Function

        End If

        WindowProc = CallWindowProc(WndProc, hw, uMsg, wParam, lParam)

    End If
        
    Exit Function

WindowProc_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.SysTray.WindowProc " & _
       "at line " & Erl
    Resume Next
        
End Function

