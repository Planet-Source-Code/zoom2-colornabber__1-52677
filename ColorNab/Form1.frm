VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ®ColorNab®"
   ClientHeight    =   3495
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   1170
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Autosize"
      ForeColor       =   &H00FF0000&
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   1485
      Width           =   2220
      Begin VB.OptionButton optSize 
         Caption         =   "72x72"
         Height          =   195
         Index           =   5
         Left            =   1260
         TabIndex        =   10
         Top             =   630
         Width           =   780
      End
      Begin VB.OptionButton optSize 
         Caption         =   "48x48"
         Height          =   195
         Index           =   4
         Left            =   1260
         TabIndex        =   9
         Top             =   427
         Width           =   780
      End
      Begin VB.OptionButton optSize 
         Caption         =   "32x32"
         Height          =   195
         Index           =   3
         Left            =   1260
         TabIndex        =   8
         Top             =   225
         Width           =   780
      End
      Begin VB.OptionButton optSize 
         Caption         =   "16x16"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   630
         Width           =   780
      End
      Begin VB.OptionButton optSize 
         Caption         =   "4x4"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   427
         Width           =   780
      End
      Begin VB.OptionButton optSize 
         Caption         =   "1x1"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   45
      TabIndex        =   1
      Top             =   -90
      Width           =   2130
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   870
         Left            =   45
         ScaleHeight     =   840
         ScaleWidth      =   1785
         TabIndex        =   2
         ToolTipText     =   "dbl click to copy actual color (size of box) to clipboard"
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2700
      Top             =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   2340
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   -90
      TabIndex        =   0
      ToolTipText     =   "dbl click to copy the RGB value to clipboard"
      Top             =   3150
      Width           =   2310
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "     &Capture Color"
      Begin VB.Menu mnuFromDialog 
         Caption         =   "                    From Dialog"
      End
      Begin VB.Menu mnuCopyToClip 
         Caption         =   "&Copy color block(color box) top clipboard"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mPT As POINTAPI

Public bCaptureColor As Boolean
Private lCLR&, lfHwnd&, lDC&, ScreenX&, ScreenY&, lClrTemp&


 

Private Sub Form_Load()
Dim hBit&
  On Error Resume Next
  ScreenX = Screen.TwipsPerPixelX
  ScreenY = Screen.TwipsPerPixelY
  SetWindowPos hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
  Call msgboxInform
  Call GetPixelColor
  Call ShowPicBoxSize
  Picture1.Cls
  hHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, App.ThreadID)
End Sub
 

Private Sub Label2_DblClick()
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Label2
    Call TempMsgBox
End Sub

 

Private Sub mnuFromDialog_Click()
On Error GoTo ERR:
   With cmDlg
       .CancelError = True
       .ShowColor
        Picture1.Cls
        Picture1.BackColor = .Color
        lClrTemp = .Color
        Label2 = RgbParse(Picture1.hdc, 0, 0)
   End With
   Exit Sub
ERR:
End Sub



Private Sub mnuHelp_Click()
  MsgBox "See " & Chr(34) & " cnHelp.txt" & Chr(34) & " which is in this folder"
End Sub


Private Sub msgboxInform()
  MsgBox "Press and hold the SHIFT key to capture the pixel color under the mouse" & vbCrLf & _
         "When you have captured the color you desire, release the SHIFT key." & vbCrLf & vbCrLf & _
         "Dbl clicking the label at the bottom of the form that displays the" & vbCrLf & _
         "RGB() color value of the captured color will copy the RGB() value" & vbCrLf & _
         "to the system clipboard." & vbCrLf & vbCrLf & _
         "Dbl clicking the color box itself will capture the actual color block" & vbCrLf & _
         "to the clipboard and may be pasted to, your graphics program, for instance.", vbExclamation, "®ColorNab®"
End Sub
'fast size pic box to known icon sizes
Private Sub optSize_Click(index As Integer)
On Error Resume Next
  Dim num%
  With Picture1
     .Cls
     Select Case index 'we have to add 2 to make up for the 2 pixels that
                       'the pic boxes border consumes
        Case Is = 0:   num = 3
        Case Is = 1:   num = 6
        Case Is = 2:   num = 18
        Case Is = 3:   num = 34
        Case Is = 4:   num = 50
        Case Is = 5:   num = 74
     End Select
      
      .Width = (num * ScreenX):   .Height = (num * ScreenY)
      .BackColor = lClrTemp
      Call ShowPicBoxSize 'update label
  End With

End Sub


'next two subs copy the actual contents of pic box to clipboard
Private Sub mnuCopyToClip_Click()
On Error GoTo ERR:
  Clipboard.Clear
  Clipboard.SetData Picture1.Image
  Call TempMsgBox
  Exit Sub
ERR:
End Sub

Private Sub Picture1_DblClick()
On Error GoTo ERR:
  Clipboard.Clear
  Clipboard.SetData Picture1.Image
  Call TempMsgBox
  Exit Sub
ERR:
End Sub

Private Sub TempMsgBox()
    SetTimer hwnd, 1001, 1000, AddressOf TimerProc
    MsgBox "Clipboard success !!", , "®ColorNab®"
End Sub
'allows user to resize picture box at runntime
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim nParam As Long

 With Picture1
  'resizing from left to right
    If X > 0 And X < 150 Then
            nParam = HTLEFT
    ElseIf X > (.Width - CURSBUFF) And (X < .Width) Then
            nParam = HTRIGHT
     'resize top to bottom
    ElseIf Y > 0 And Y < CURSBUFF Then
            nParam = HTTOP
    ElseIf Y > (.Height - CURSBUFF) And (Y < .Height) Then
            nParam = HTBOTTOM
    End If

    If nParam Then
        .Cls
        Call ReleaseCapture
        Call SendMessage(.hwnd, WM_NCLBUTTONDOWN, nParam, 0)
        .BackColor = lClrTemp
    End If
    
    Call ShowPicBoxSize
 End With
End Sub

 
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim NewPointer As MousePointerConstants

With Picture1
  'creates E/W pointer at left/right edge of control
    If X > 0 And X < 100 Then
        NewPointer = vbSizeWE
    ElseIf X > (.Width - CURSBUFF) And (X < .Width) Then
        NewPointer = vbSizeWE
    'createsN/S pointer at top/bottom  edge of control
    ElseIf Y > 0 And (Y < CURSBUFF) Then
        NewPointer = vbSizeNS
    ElseIf Y > (.Height - CURSBUFF) And Y < .Height Then
        NewPointer = vbSizeNS
    Else 'move the picture box
        If Button = 1 Then
             NewPointer = vbSizeAll
             ReleaseCapture
             SendMessage .hwnd, WM_SYSCOMMAND, &HF012&, 0&
         Else
             NewPointer = vbDefault
         End If
    End If

    If NewPointer <> .MousePointer Then
        .MousePointer = NewPointer
    End If
 End With
End Sub


Public Sub GetPixelColor()
 On Error Resume Next
  'color under mouse captured as long as shift is being pressed
   GetCursorPos mPT
   lDC = GetWindowDC(0)
   lCLR = GetPixel(lDC, mPT.X, mPT.Y)
   Picture1.BackColor = lCLR
   lClrTemp = lCLR
   Label2.ForeColor = lCLR
   Label2 = RgbParse(lDC, mPT.X, mPT.Y)
   ReleaseDC 0, lDC
End Sub

'stop hooking keyboard events
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
   UnhookWindowsHookEx hHook
End Sub



'display the current picboxes with and height in pixels and twips
Private Sub ShowPicBoxSize()
On Error Resume Next
  With Picture1
    Label1.Caption = "PIXEL WIDTH: " & ((.Width / ScreenX) - 2) & vbCrLf & _
                     "PIXEL HEIGHT: " & ((.Height / ScreenY) - 2) & vbCrLf & _
                     "TWIP WIDTH: " & (.Width - (2 * ScreenX)) & vbCrLf & _
                     "TWIP HEIGHT: " & (.Height - (2 * ScreenY))
  End With
End Sub

'this function converts the long value of the getpixel and converts to rgb
Private Function RgbParse(hdc As Long, X&, Y&) As String
On Error Resume Next
    Dim ColorMe&, rgbRed&, rgbGreen&, rgbBlue&

    If X = 0 And Y = 0 Then
        ColorMe = Picture1.BackColor
    Else
        ColorMe = GetPixel(hdc, X, Y)
    End If
    
    rgbRed = Abs(ColorMe Mod &H100)
    ColorMe = Abs(ColorMe \ &H100)
    rgbGreen = Abs(ColorMe Mod &H100)
    ColorMe = Abs(ColorMe \ &H100)
    rgbBlue = Abs(ColorMe Mod &H100)
    ColorMe = RGB(rgbRed, rgbGreen, rgbBlue)
    RgbParse = "RGB(" & rgbRed & ", " & rgbGreen & ", " & rgbBlue & ")"
End Function

