Attribute VB_Name = "Key_Events"
Public Const VK_LBUTTON = &H1 'Left mouse button
Public Const VK_RBUTTON = &H2 'Right mouse button
Public Const VK_CANCEL = &H3 'Control-break processing
Public Const VK_MBUTTON = &H4 'Middle mouse button (three-button mouse)
Public Const VK_BACK = &H8 'backspace key
Public Const VK_TAB = &H9 'tab key
Public Const VK_CLEAR = &HC 'clear key
Public Const VK_RETURN = &HD 'enter key
Public Const VK_SHIFT = &H10 'shift key
Public Const VK_CONTROL = &H11 'ctrl key
Public Const VK_MENU = &H12 'alt key
Public Const VK_PAUSE = &H13 'pause key
Public Const VK_CAPITAL = &H14 'caps lock key
Public Const VK_ESCAPE = &H1B 'esc key
Public Const VK_SPACE = &H20 'spacebar
Public Const VK_PRIOR = &H21 'page up key
Public Const VK_NEXT = &H22 'page down key
Public Const VK_END = &H23 'end key
Public Const VK_HOME = &H24 'home key
Public Const VK_LEFT = &H25 'left arrow key
Public Const VK_UP = &H26 'up arrow key
Public Const VK_RIGHT = &H27 'right arrow key
Public Const VK_DOWN = &H28 'down arrow key
Public Const VK_SELECT = &H29 'select key
Public Const VK_EXECUTE = &H2B 'execute key
Public Const VK_SNAPSHOT = &H2C 'print screen key
Public Const VK_INSERT = &H2D 'ins key
Public Const VK_DELETE = &H2E 'del key
Public Const VK_HELP = &H2F 'help key
Public Const VK_0 = &H30 '0 key
Public Const VK_1 = &H31 '1 key
Public Const VK_2 = &H32 '2 key
Public Const VK_3 = &H33 '3 key
Public Const VK_4 = &H34 '4 key
Public Const VK_5 = &H35 '5 key
Public Const VK_6 = &H36 '6 key
Public Const VK_7 = &H37 '7 key
Public Const VK_8 = &H38 '8 key
Public Const VK_9 = &H39 '9 key
Public Const VK_A = &H41 'a key
Public Const VK_B = &H42 'b key
Public Const VK_C = &H43 'c key
Public Const VK_D = &H44 'd key
Public Const VK_E = &H45 'e key
Public Const VK_F = &H46 'f key
Public Const VK_G = &H47 'g key
Public Const VK_H = &H48 'h key
Public Const VK_I = &H49 'i key
Public Const VK_J = &H4A 'j key
Public Const VK_K = &H4B 'k key
Public Const VK_L = &H4C 'l key
Public Const VK_M = &H4D 'm key
Public Const VK_N = &H4E 'n key
Public Const VK_O = &H4F 'o key
Public Const VK_P = &H50 'p key
Public Const VK_Q = &H51 'q key
Public Const VK_R = &H52 'r key
Public Const VK_S = &H53 's key
Public Const VK_T = &H54 't key
Public Const VK_U = &H55 'u key
Public Const VK_V = &H56 'v key
Public Const VK_W = &H57 'w key
Public Const VK_X = &H58 'x key
Public Const VK_Y = &H59 'y key
Public Const VK_Z = &H5A 'z key
Public Const VK_LWIN = &H5B 'Left Windows key (Microsoft Natural Keyboard)
Public Const VK_RWIN = &H5C 'Right Windows key (Microsoft Natural Keyboard)
Public Const VK_APPS = &H5D 'Applications key (Microsoft Natural Keyboard)
Public Const VK_NUMPAD0 = &H60 'Numeric keypad 0 key
Public Const VK_NUMPAD1 = &H61 'Numeric keypad 1 key
Public Const VK_NUMPAD2 = &H62 'Numeric keypad 2 key
Public Const VK_NUMPAD3 = &H63 'Numeric keypad 3 key
Public Const VK_NUMPAD4 = &H64 'Numeric keypad 4 key
Public Const VK_NUMPAD5 = &H65 'Numeric keypad 5 key
Public Const VK_NUMPAD6 = &H66 'Numeric keypad 6 key
Public Const VK_NUMPAD7 = &H67 'Numeric keypad 7 key
Public Const VK_NUMPAD8 = &H68 'Numeric keypad 8 key
Public Const VK_NUMPAD9 = &H69 'Numeric keypad 9 key
Public Const VK_MULTIPLY = &H6A 'Multiply key
Public Const VK_ADD = &H6B 'Add key
Public Const VK_SEPARATOR = &H6C 'Separator key
Public Const VK_SUBTRACT = &H6D 'Subtract key
Public Const VK_DECIMAL = &H6E 'Decimal key
Public Const VK_DIVIDE = &H6F 'Divide key
Public Const VK_F1 = &H70 'f1 key
Public Const VK_F2 = &H71 'f2 key
Public Const VK_F3 = &H72 'f3 key
Public Const VK_F4 = &H73 'f4 key
Public Const VK_F5 = &H74 'f5 key
Public Const VK_F6 = &H75 'f6 key
Public Const VK_F7 = &H76 'f7 key
Public Const VK_F8 = &H77 'f8 key
Public Const VK_F9 = &H78 'f9 key
Public Const VK_F10 = &H79 'f10 key
Public Const VK_F11 = &H7A 'f11 key
Public Const VK_F12 = &H7B 'f12 key
Public Const VK_F13 = &H7C 'f13 key
Public Const VK_F14 = &H7D 'f14 key
Public Const VK_F15 = &H7E 'f15 key
Public Const VK_F16 = &H7F 'f16 key
Public Const VK_F17 = &H80 'f17 key
Public Const VK_F18 = &H81 'f18 key
Public Const VK_F19 = &H82 'f19 key
Public Const VK_F20 = &H83 'f20 key
Public Const VK_F21 = &H84 'f21 key
Public Const VK_F22 = &H85 'f22 key
Public Const VK_F23 = &H86 'f23 key
Public Const VK_F24 = &H87 'f24 key
Public Const VK_NUMLOCK = &H90 'num lock key
Public Const VK_SCROLL = &H91 'scroll lock key

Private Const KEYEVENTF_KEYUP = &H2
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Function Control_C()
    Call keybd_event(VK_CONTROL, 0, 0&, 0&)
    Call keybd_event(VK_C, 0, 0&, 0&)
    Call keybd_event(VK_CONTROL, 0, 2, 0&)
    DoEvents
End Function

Public Function Control_V()
    Call keybd_event(VK_CONTROL, 0, 0&, 0&)
    Call keybd_event(VK_V, 0, 0&, 0&)
    Call keybd_event(VK_CONTROL, 0, 2, 0&)
    DoEvents
End Function

Public Function Enter()
    Call keybd_event(VK_RETURN, 0, 0&, 0&)
End Function
