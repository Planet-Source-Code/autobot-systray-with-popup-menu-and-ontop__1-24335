Attribute VB_Name = "OnTop"
'--------------------------------------------------------------------------------
'    Component  : OnTop
'    Project    : RealSysTray
'
'    Description: Updated the code to include putting your form on top
'                 of all others.
'
'    Modified   : June 23, 2001
'--------------------------------------------------------------------------------
Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal wFlags As Long) As Long
   
'Takes form off top of all applications

Sub TakeOffTop(F As Form)
        
    On Error GoTo TakeOffTop_Err

    
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
    Exit Sub

TakeOffTop_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.OnTop.TakeOffTop " & _
       "at line " & Erl
    Resume Next
        
End Sub

'Puts form on top off all applications running

Sub PutOnTop(F As Form)
        
    On Error GoTo PutOnTop_Err

    
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
    Exit Sub

PutOnTop_Err:
    MsgBox Err.Description & vbCrLf & _
       "in RealSysTray.OnTop.PutOnTop " & _
       "at line " & Erl
    Resume Next
        
End Sub

