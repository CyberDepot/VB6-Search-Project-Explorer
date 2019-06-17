Attribute VB_Name = "Mod_Main"
Option Explicit

Public mObjDoc As UD_Search
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public mWindow As Window

'Droup Down ComboBox
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_SHOWDROPDOWN = &H14F

'Save Setting
Public Show_Disign As Boolean
Public Show_Code As Boolean
Public Close_Disign As Boolean
Public Close_Code As Boolean

'Send Mail
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL         As Long = 1
Private Const SE_NO_ERROR           As Long = 33 'Values below 33 are error returns

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Enum ShowFormTypes
    [f_Ontop] = 0
    [f_Normal] = 1
End Enum


'Monitor Keyboard

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Public KeyboardHandle As Long
Private hWndActiveCodePane As Long
Public Const WH_KEYBOARD = 2

' Virtual Keys, Standard Set
Const VK_LBUTTON = &H1
Const VK_RBUTTON = &H2
Const VK_CANCEL = &H3
Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Const VK_BACK = &H8
Const VK_TAB = &H9

Const VK_CLEAR = &HC
Const VK_RETURN = &HD

Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12
Const VK_PAUSE = &H13
Const VK_CAPITAL = &H14

Const VK_ESCAPE = &H1B

Const VK_SPACE = &H20
Const VK_PRIOR = &H21
Const VK_NEXT = &H22
Const VK_END = &H23
Const VK_HOME = &H24
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28
Const VK_SELECT = &H29
Const VK_PRINT = &H2A
Const VK_EXECUTE = &H2B
Const VK_SNAPSHOT = &H2C
Const VK_INSERT = &H2D
Const VK_DELETE = &H2E
Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Const VK_NUMPAD0 = &H60
Const VK_NUMPAD1 = &H61
Const VK_NUMPAD2 = &H62
Const VK_NUMPAD3 = &H63
Const VK_NUMPAD4 = &H64
Const VK_NUMPAD5 = &H65
Const VK_NUMPAD6 = &H66
Const VK_NUMPAD7 = &H67
Const VK_NUMPAD8 = &H68
Const VK_NUMPAD9 = &H69
Const VK_MULTIPLY = &H6A
Const VK_ADD = &H6B
Const VK_SEPARATOR = &H6C
Const VK_SUBTRACT = &H6D
Const VK_DECIMAL = &H6E
Const VK_DIVIDE = &H6F
Const VK_F1 = &H70
Const VK_F2 = &H71
Const VK_F3 = &H72
Const VK_F4 = &H73
Const VK_F5 = &H74
Const VK_F6 = &H75
Const VK_F7 = &H76
Const VK_F8 = &H77
Const VK_F9 = &H78
Const VK_F10 = &H79
Const VK_F11 = &H7A
Const VK_F12 = &H7B
Const VK_F13 = &H7C
Const VK_F14 = &H7D
Const VK_F15 = &H7E
Const VK_F16 = &H7F
Const VK_F17 = &H80
Const VK_F18 = &H81
Const VK_F19 = &H82
Const VK_F20 = &H83
Const VK_F21 = &H84
Const VK_F22 = &H85
Const VK_F23 = &H86
Const VK_F24 = &H87

Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91

'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /
Const VK_LSHIFT = &HA0
Const VK_RSHIFT = &HA1
Const VK_LCONTROL = &HA2
Const VK_RCONTROL = &HA3
Const VK_LMENU = &HA4
Const VK_RMENU = &HA5

Const VK_ATTN = &HF6
Const VK_CRSEL = &HF7
Const VK_EXSEL = &HF8
Const VK_EREOF = &HF9
Const VK_PLAY = &HFA
Const VK_ZOOM = &HFB
Const VK_NONAME = &HFC
Const VK_PA1 = &HFD
Const VK_OEM_CLEAR = &HFE

Public Sub SetOntop(ByRef MyFrm As Form, Optional FormMode As ShowFormTypes = f_Ontop)
    Select Case FormMode
    
        Case Is = f_Ontop
            'To set form to be "on top" of other forms use this code:
            Call SetWindowPos(MyFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
        Case Is = f_Normal
            'To return the form to normal, this is the code you need:
            Call SetWindowPos(MyFrm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            
    End Select
End Sub

Public Sub SendMeMail(FromhWnd As Long, Subject As String)
    If ShellExecute(FromhWnd, vbNullString, "mailto:UMGEDV@AOL.COM?subject=" & Subject & " &body=Hi Ulli,    .......    Best regards from ", vbNullString, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        Beep
        MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
    End If
End Sub

Public Sub HookKeyboard()
    hWndActiveCodePane = FindWindowEx(VBInstance.MainWindow.hwnd, 0, "MDIClient", vbNullString) 'find topmost (active) child window of class "MDIClient" in VB's main MDI window
    
    If hWndActiveCodePane Then 'found one - should be a code pane window
        KeyboardHandle = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, App.ThreadID)
    End If
End Sub

Public Function KeyboardProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        'if idHook is less than zero, no further processing is required
        If idHook < 0 Then
            'call the next hook
            KeyboardProc = CallNextHookEx(KeyboardHandle, idHook, wParam, ByVal lParam)
        Else
        
        Static MyTimer As Long
        
        'check if Left ALT-S is pressed
        If (GetKeyState(VK_LMENU) And &HF0000000) And wParam = Asc("S") Then
            'if mwindow not set
            If mWindow Is Nothing Then Exit Function
            
            'ÏÇÑ ÇÎÊáÇá äÔå IDE íå ˜äÊÑá ÒãÇäí ÐÇÔÊã ÊÇ Èå ÎÇØÑ ÚæÖ ÔÏä Ý˜æÓ
            If MyTimer = 0 Then MyTimer = Timer - 3
            If Timer - MyTimer < 0.3 Then Exit Function
            MyTimer = Timer
            
            'if windows is hidden first show it
            If mWindow.Visible = False Then mWindow.Visible = True
            
            'set focus to my add-in
            mWindow.SetFocus
        End If
        
        'check if Left ALT-C is pressed
        If (GetKeyState(VK_LMENU) And &HF0000000) And wParam = Asc("Z") Then
            'ÏÇÑ ÇÎÊáÇá äÔå IDE íå ˜äÊÑá ÒãÇäí ÐÇÔÊã ÊÇ Èå ÎÇØÑ ÚæÖ ÔÏä Ý˜æÓ
            If MyTimer = 0 Then MyTimer = Timer - 3
            If Timer - MyTimer < 0.5 Then Exit Function
            MyTimer = Timer
            
            Dim Comp As VBComponent
            Dim Ctl As VBControl
            Dim Frm As VBForm
            
            For Each Comp In VBInstance.ActiveVBProject.VBComponents
                    If Close_Code = True Then Comp.CodeModule.CodePane.Window.Visible = False
                    If Close_Disign = True Then Comp.DesignerWindow.Visible = False
            Next
        End If
        
        'check if Left ALT-X is pressed
        If (GetKeyState(VK_LMENU) And &HF0000000) And wParam = Asc("X") Then
            'ÏÇÑ ÇÎÊáÇá äÔå IDE íå ˜äÊÑá ÒãÇäí ÐÇÔÊã ÊÇ Èå ÎÇØÑ ÚæÖ ÔÏä Ý˜æÓ
            If MyTimer = 0 Then MyTimer = Timer - 3
            If Timer - MyTimer < 0.5 Then Exit Function
            MyTimer = Timer
            
            frmAddIn.TimerImmadiate.Enabled = True
        End If
        
        'call the next hook
        KeyboardProc = CallNextHookEx(KeyboardHandle, idHook, wParam, ByVal lParam)
    End If
End Function

Private Function Hooked() As Boolean
    Hooked = KeyboardHandle <> 0
End Function

Public Sub UnhookKeyboard()
    If (Hooked) Then Call UnhookWindowsHookEx(KeyboardHandle)
End Sub

