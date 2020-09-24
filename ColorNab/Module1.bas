Attribute VB_Name = "Module1"
Option Explicit

Public Const WH_KEYBOARD = 2
Public Const VK_SHIFT = &H10

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const WM_SYSCOMMAND As Long = &H112
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const VK_LBUTTON As Long = &H1
Public Const HTBOTTOM As Long = 15
Public Const HTTOP As Long = 12
Public Const HTTOPLEFT As Long = 13
Public Const HTTOPRIGHT As Long = 14
Public Const HTBOTTOMLEFT As Long = 16
Public Const HTBOTTOMRIGHT As Long = 17
Public Const HTLEFT As Long = 10
Public Const HTRIGHT As Long = 11
Public Const CURSBUFF = 110

Type POINTAPI
    X As Long
    Y As Long
End Type

Public hHook As Long
Public Function KeyboardProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'if idHook is less than zero, no further processing is required
    If idHook < 0 Then
        'call the next hook
        KeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
    Else
        'check if SHIFT is pressed
        If (GetKeyState(VK_SHIFT) And &HF0000000) Then
             Form1.GetPixelColor
        End If
        'call the next hook
        KeyboardProc = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
    End If
End Function

Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    SendKeys "{ENTER}"
    KillTimer Form1.hwnd, 1001
End Sub
