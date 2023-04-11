VERSION 5.00
Begin VB.UserControl MorphListBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "ucMorphListBox.ctx":0000
End
Attribute VB_Name = "MorphListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* MorphListBox v1.20 - VB6 ownerdrawn listbox control replacement.      *
'* Author: Matthew R. Usner, Sept 2005 for www.planet-source-code.com.   *
'* Copyright ©2006 - 2007, Matthew R. Usner.  All rights reserved.       *
'* The latest version of this control can be found at:                   *
'* www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=63527&lngWId=1       *
'*************************************************************************
'* Last Updates: 23 Nov 2006 - Added 'MorphBorder' 3D gradient border,   *
'* tweaked display (focus rect, selection bar) for XP, added speed comp- *
'* arison demo form, fixed gradient calculation and display bugs for the *
'* scrollbar that only became visible when I upgraded from Windows 98 to *
'* WinXP.  Added .ClearOrSelect public method to allow user to clear or  *
'* set selected status of supplied range of listitems (or entire list).  *
'* Made theme property lists more comprehensive.  Organized properties.  *
'*************************************************************************
'* A graphical replacement and major functional enhancement for the VB   *
'* listbox control.  Standard listbox behavior is consistently emulated, *
'* with the exception of a couple intricacies of list item selection in  *
'* .MultiSelect Extended I thought unnecessary.  Loads lists much faster *
'* than standard listbox, depending on system. Control features an integ-*
'* rated graphical scrollbar.  Background can be a linear / circular     *
'* gradient or bitmap.  Border can be a 3D gradient.  Small background   *
'* bitmaps can be tiled or stretched.  Icons or small bitmaps can be ass-*
'* igned and displayed next to desired list items.  >32767 listitems,    *
'* limited only by system resources.  8 color schemes can be selected via*
'* .Theme property, and you can add your own themes as desired.  Unicode *
'* character display supported.  Drag-and-drop capability is included.   *
'* Sort order can be maintained numerically for lists of numbers, though *
'* numbers are still stored as strings.  Any portion of list can be dis- *
'* played via .DisplayFrom method.  A .MouseOverIndex method allows det- *
'* ermination of the list item (index or item) the mouse cursor is hov-  *
'* ering over.  .RightToLeft property allows for easy use of the control *
'* for users whose written language is right to left.  Ranges of items   *
'* can be selected or deselected using the .ClearOrSelect method.        *
'*************************************************************************
'* <<<<<<< FEEDBACK ALWAYS WELCOME... VOTES ALWAYS APPRECIATED! >>>>>>>  *
'*************************************************************************
'* Miscellaneous Usage Notes:                                            *
'*   1) Due to size of control (>380 procedures), this is best used as a *
'*      compiled .OCX.                                                   *
'*   2) IMPORTANT!!!  When filling the MorphListBox with a large number  *
'*      of list items, set the .RedrawFlag property to False prior to    *
'*      the loop that fills the list.  Afterwards, set .RedrawFlag to    *
'*      True.  This is a HUGE timesaver because list won't redraw after  *
'*      adding each item.  Use this same technique when performing any   *
'*      loop operations on a large number of items.  I can't stress this *
'*      enough!                                                          *
'*   3) I have noticed listitem focus rectangle and checkmark display    *
'*      location differences when using Win98 and XP. The differences    *
'*      are only 1 pixel but can be distracting.  For your own projects, *
'*      adjust according to OS in the appropriate display routines.      *
'*   4) Since this uses subclassing, use Unload Me, not End, in projects.*
'*      Do NOT use the 'End' button in the IDE.                          *
'*************************************************************************
'* Legal:  Redistribution of this code, whole or in part, as source code *
'* or in binary form, alone or as part of a larger distribution or prod- *
'* uct, is forbidden for any commercial or for-profit use without the    *
'* author's explicit written permission.                                 *
'*                                                                       *
'* Non-commercial redistribution of this code, as source code or in      *
'* binary form, with or without modification, is permitted provided that *
'* the following conditions are met:                                     *
'*                                                                       *
'* Redistributions of source code must include this list of conditions,  *
'* and the following acknowledgment:                                     *
'*                                                                       *
'* This VB6 usercontrol was developed by Matthew R. Usner.               *
'* Source code, written in Visual Basic 6.0, is freely available for     *
'* non-commercial, non-profit use.                                       *
'*                                                                       *
'* Redistributions in binary form, as part of a larger project, must     *
'* include the above acknowledgment in the end-user documentation.       *
'* Alternatively, the above acknowledgment may appear in the software    *
'* itself, if and where such third-party acknowledgments normally appear.*
'*************************************************************************
'* Credits and Thanks:                                                   *
'* Carles P.V., for the gradient, bitmap tiling, and corner rounding     *
'* routines.                                                             *
'* LaVolpe, for the gradient border segment generation code.             *
'* Paul Caton, for the self-subclassing usercontrol code.                *
'* Phillipe Lord, for his array handling routines.  His original module  *
'* can be found at www.pscode.com/vb, txtCodeId=24546.                   *
'* Richard Mewett, for the Unicode support routines.                     *
'* Paul Turcksin, for spending hours checking this before I submitted.   *
'* Jeff Mayes, for the .SortAsNumeric idea.                              *
'* Redbird77, for fixing a glitch with the background gradient draw and  *
'* reorganizing and optimizing the DisplayListBoxItem routine.           *
'* xpert, for reporting a minor design mode graphics update bug.         *
'*************************************************************************

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2        ' for generating a mousedown event to replace double click.
Private m_hBrush As Long                        ' pattern brush for bitmap tiling.

'==================================================================================================
' Subclasser declarations.
' windows messages to be intercepted by subclassing.
Private Const WM_MOUSEMOVE            As Long = &H200
Private Const WM_MOUSELEAVE           As Long = &H2A3
Private Const WM_SETFOCUS             As Long = &H7
Private Const WM_KILLFOCUS            As Long = &H8
Private Const WM_MOUSEWHEEL           As Long = &H20A

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize                             As Long
   dwFlags                            As TRACKMOUSEEVENT_FLAGS
   hwndTrack                          As Long
   dwHoverTime                        As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private bInCtrl                       As Boolean          ' flag that indicates if mouse is in control.

Private Enum eMsgWhen
   MSG_AFTER = 1                                          'Message calls back after the original (previous) WndProc
   MSG_BEFORE = 2                                         'Message calls back before the original (previous) WndProc
   MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE         'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES            As Long = -1        'All messages added or deleted
Private Const GMEM_FIXED              As Long = 0         'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC             As Long = -4        'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04                As Long = 88        'Table B (before) address patch offset
Private Const PATCH_05                As Long = 93        'Table B (before) entry count patch offset
Private Const PATCH_08                As Long = 132       'Table A (after) address patch offset
Private Const PATCH_09                As Long = 137       'Table A (after) entry count patch offset

Private Type tSubData                                     'Subclass data type
   hwnd                               As Long             'Handle of the window being subclassed
   nAddrSub                           As Long             'The address of our new WndProc (allocated memory).
   nAddrOrig                          As Long             'The address of the pre-existing WndProc
   nMsgCntA                           As Long             'Msg after table entry count
   nMsgCntB                           As Long             'Msg before table entry count
   aMsgTblA()                         As Long             'Msg after table array
   aMsgTblB()                         As Long             'Msg Before table array
End Type
Private sc_aSubData()                 As tSubData         'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'==================================================================================================

'  declares for Unicode support.
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
   dwOSVersionInfoSize                As Long
   dwMajorVersion                     As Long
   dwMinorVersion                     As Long
   dwBuildNumber                      As Long
   dwPlatformId                       As Long
   szCSDVersion                       As String * 128     '  Maintenance string for PSS usage
End Type
Private mWindowsNT                    As Boolean
Private Const DT_SINGLELINE           As Long = &H20      ' strip cr/lf from string before draw.
Private Const DT_NOPREFIX             As Long = &H800     ' ignore access key ampersand.
Private Const DT_LEFT                 As Long = &H0       ' draw from left edge of rectangle.
Private Const DT_RIGHT                As Long = &H2

' declares for gradient painting and bitmap tiling.
Private Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type

Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0                  ' also used in gradient generation.
Private Const OBJ_BITMAP     As Long = 7                  ' used to determine if picture is a bitmap.

'  used to define various graphics areas and listbox component locations.
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

' declares for Ralph Eastwood's 'American Flag' radix sort.
'----------------------------------------------------------------
Private Const MAX_CHARACTER As Long = 65536
Private m_atStack() As STRING_STACK
Private m_saStringHeader1 As SAFEARRAYHEADER
Private m_aiString1() As Integer
Private m_saStringHeader2 As SAFEARRAYHEADER
Private m_aiString2() As Integer
Private m_saStringPtrArrayHeader As SAFEARRAYHEADER
Private m_alStringPtrArray() As Long
Private m_alPile(0 To MAX_CHARACTER - 1) As Long
Private m_alCount(0 To MAX_CHARACTER - 1) As Long
Private Const AF_THRESHOLD As Long = 10
Private Type STRING_STACK
    lpStringArray As Long
    lStringCount As Long
    lStringIndex As Long
End Type
Private Type SAFEARRAYHEADER
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    cElements   As Long
    lLBound     As Long
End Type
'----------------------------------------------------------------

Private Const AF_SIZE As Long = 1024

' enum tied to the .MultiSelect property.
Public Enum SelectionOptions
   [None] = 0
   [Simple] = 1
   [Extended] = 2
End Enum

' enum tied to the .Style property.
Public Enum ListItemOptions
   [Standard] = 0
   [CheckBox] = 1
End Enum

' enum tied to .Theme property.
Public Enum ThemeOptions
   [None] = 0
   [Cyan Eyed] = 1
   [Gunmetal Grey] = 2
   [Blue Moon] = 3
   [Red Rum] = 4
   [Green With Envy] = 5
   [Purple People Eater] = 6
   [Golden Goose] = 7
   [Penny Wise] = 8
End Enum

'  enum tied to the .DoubleClickBehavior property.
Public Enum DblClickBehaviorOptions
   [Double Click] = 0
   [Two Single Clicks] = 1
End Enum

'  enum tied to .PictureMode property.
Public Enum MLB_PictureModeOptions
   [Normal] = 0
   [Stretch] = 1
   [Tiled] = 2
End Enum

'  enum tied to .CheckStyle property.
Public Enum CheckStyleOptions
   [Arrow] = 0
   [Tick] = 1
   [X] = 2
End Enum

'  list display declares.
'  type used to hold indices of the first and last displayable list items, based on current listbox height.
Private Type DisplayRangeType
   FirstListItem As Long                                  ' the index of the first displayed list item.
   LastListItem As Long                                   ' the index of the last displayed list item.
End Type
Private DisplayRange               As DisplayRangeType    ' indices of first and last displayed listbox items.

Private ListItemHeight             As Long                ' height of text in current font.
Private Const Y_CLEARANCE          As Long = 5            ' pixel offset to start and stop displaying list items
Private YCoords(0 To 99)           As Long                ' y coordinate of each displayable list item.
Private MaxDisplayItems            As Long                ' max displayable based on height, font, borderwidth.
Private TextClearance              As Long                ' # pixels from left edge to start drawing text.
Private Const MIN_FONT_HEIGHT      As Long = 13           ' helps adjust item spacing when using very small fonts.
Private ChangingPicture            As Boolean             ' if true, re-blits new picture to virtual DC.
Private Const NO_IMAGE             As Long = -1           ' no image assigned to a particular listitem.
Private Const EQUAL_TO_TEXTHEIGHT  As Long = 0            ' item image will be height/width of text height.

'  'active' variables - these are the values actually used in displaying the control.
'  The Enabled or Disabled color property sets are transferred into these variables.
Private m_ActiveArrowDownColor     As OLE_COLOR
Private m_ActiveArrowUpColor       As OLE_COLOR
Private m_ActiveBackColor1         As OLE_COLOR
Private m_ActiveBackColor2         As OLE_COLOR
Private m_ActiveBorderColor1       As OLE_COLOR
Private m_ActiveBorderColor2       As OLE_COLOR
Private m_ActiveButtonColor1       As OLE_COLOR
Private m_ActiveButtonColor2       As OLE_COLOR
Private m_ActiveCheckboxArrowColor As OLE_COLOR
Private m_ActiveCheckBoxColor      As OLE_COLOR
Private m_ActiveFocusRectColor     As OLE_COLOR
Private m_ActivePicture            As StdPicture
Private m_ActivePictureMode        As MLB_PictureModeOptions
Private m_ActiveSelColor1          As OLE_COLOR
Private m_ActiveSelColor2          As OLE_COLOR
Private m_ActiveSelTextColor       As OLE_COLOR
Private m_ActiveTextColor          As OLE_COLOR
Private m_ActiveThumbBorderColor   As OLE_COLOR
Private m_ActiveThumbColor1        As OLE_COLOR
Private m_ActiveThumbColor2        As OLE_COLOR
Private m_ActiveTrackBarColor1     As OLE_COLOR
Private m_ActiveTrackBarColor2     As OLE_COLOR

'  property variables.
Private m_ArrowDownColor        As OLE_COLOR                 ' scroll arrow color when button is down.
Private m_ArrowUpColor          As OLE_COLOR                 ' scroll arrow color when button is up.
Private m_AutoRedraw            As Boolean                   ' usercontrol .AutoRedraw property.
Private m_BackAngle             As Single                    ' background gradient display angle
Private m_BackColor1            As OLE_COLOR                 ' the first gradient color of the background.
Private m_BackColor2            As OLE_COLOR                 ' the second gradient color of the background.
Private m_BackMiddleOut         As Boolean                   ' flag for background middle-out gradient.
Private m_BorderColor1          As OLE_COLOR                 ' first color of gradient border.
Private m_BorderColor2          As OLE_COLOR                 ' first color of gradient border.
Private m_BorderMiddleOut       As Boolean                   ' border gradient middle-out display status.
Private m_BorderWidth           As Long                      ' width, in pixels, of border.
Private m_ButtonColor1          As OLE_COLOR                 ' first scrollbar button gradient color.
Private m_ButtonColor2          As OLE_COLOR                 ' second scrollbar button gradient color.
Private m_CheckBoxArrowColor    As OLE_COLOR                 ' selection checkbox arrow color.
Private m_CheckBoxColor         As OLE_COLOR                 ' checkbox border color.
Private m_CheckStyle            As CheckStyleOptions         ' arrow and checkmark style options.
Private m_CircularGradient      As Boolean                   ' background gradient circular? flag.
Private m_CurveBottomLeft       As Long                      ' the curvature of the bottom left corner.
Private m_CurveBottomRight      As Long                      ' the curvature of the bottom right corner.
Private m_CurveTopLeft          As Long                      ' the curvature of the top left corner.
Private m_CurveTopRight         As Long                      ' the curvature of the top right corner.
Private m_DblClickBehavior      As DblClickBehaviorOptions   ' double click or two rapid single clicks?
Private m_DisArrowDownColor     As OLE_COLOR
Private m_DisArrowUpColor       As OLE_COLOR
Private m_DisBackColor1         As OLE_COLOR
Private m_DisBackColor2         As OLE_COLOR
Private m_DisBorderColor1       As OLE_COLOR
Private m_DisBorderColor2       As OLE_COLOR
Private m_DisButtonColor1       As OLE_COLOR
Private m_DisButtonColor2       As OLE_COLOR
Private m_DisCheckboxArrowColor As OLE_COLOR
Private m_DisCheckboxColor      As OLE_COLOR
Private m_DisFocusRectColor     As OLE_COLOR
Private m_DisPicture            As Picture
Private m_DisPictureMode        As MLB_PictureModeOptions
Private m_DisSelColor1          As OLE_COLOR
Private m_DisSelColor2          As OLE_COLOR
Private m_DisSelTextColor       As OLE_COLOR
Private m_DisTextColor          As OLE_COLOR
Private m_DisThumbBorderColor   As OLE_COLOR
Private m_DisThumbColor1        As OLE_COLOR
Private m_DisThumbColor2        As OLE_COLOR
Private m_DisTrackbarColor1     As OLE_COLOR
Private m_DisTrackbarColor2     As OLE_COLOR
Private m_DragEnabled           As Boolean                   ' boolean that allows drag and drop.
Private m_Enabled               As Boolean                   ' enabled/disabled flag.
Private m_FocusBorderColor1     As OLE_COLOR                 ' first border color when control has focus.
Private m_FocusBorderColor2     As OLE_COLOR                 ' second border color when control has focus.
Private m_FocusRectColor        As OLE_COLOR                 ' custom focus rectangle color.
Private m_ItemImageSize         As Long                      ' listitem icon size (0=textheight, <>0=custom).
Private m_ListCount             As Long                      ' the number of items in the list.
Private m_ListIndex             As Long                      ' index of currently selected item; -1 if none selected.
Private m_ListFont              As Font                      ' the font to display listbox items with.
Private m_MultiSelect           As SelectionOptions          ' non/simple/extended list item selection.
Private m_NewIndex              As Long                      ' index of most recently added item.
Private m_Picture               As Picture                   ' the image to use in lieu of gradient background.
Private m_PictureMode           As MLB_PictureModeOptions    ' normal, stretched or tiled picture display.
Private m_RedrawFlag            As Boolean                   ' for redraw yes/no.
Private m_RightToLeft           As Boolean                   ' right-to-left display mode flag.
Private m_ScaleMode             As Integer                   ' usercontrol .ScaleMode property.
Private m_ScaleHeight           As Single                    ' usercontrol .ScaleHeight property.
Private m_ScaleWidth            As Single                    ' usercontrol .ScaleWidth property.
Private m_SelCount              As Long                      ' read-only selected item counter.
Private m_SelColor1             As OLE_COLOR                 ' first selection bar gradient color.
Private m_SelColor2             As OLE_COLOR                 ' second selection bar gradient color.
Private m_SelTextColor          As OLE_COLOR                 ' color to draw selected list item text.
Private m_SortAsNumeric         As Boolean                   ' is list sorted as string or numbers?
Private m_ShowItemImages        As Boolean                   ' will listitem icons be shown? flag.
Private m_ShowSelectRect        As Boolean                   ' show selected listitem focus rect flag.
Private m_Sorted                As Boolean                   ' if True, new items put in proper order.
Private m_Style                 As ListItemOptions           ' standard or checkbox listbox style.
Private m_Text                  As String                    ' text of currently selected listitem.
Private m_TextColor             As OLE_COLOR                 ' color to draw unselected list item text.
Private m_Theme                 As ThemeOptions              ' color scheme to use.
Private m_ThumbBorderColor      As OLE_COLOR                 ' border color for scroll thumb.
Private m_ThumbColor1           As OLE_COLOR                 ' first scrollbar thumb gradient color.
Private m_ThumbColor2           As OLE_COLOR                 ' second scrollbar thumb gradient color.
Private m_TopIndex              As Long                      ' index of topmost-displayed listitem.
Private m_TrackBarColor1        As OLE_COLOR                 ' first trackbar gradient color.
Private m_TrackBarColor2        As OLE_COLOR                 ' second trackbar gradient color.
Private m_TrackClickColor1      As OLE_COLOR                 ' track portion clicked first color.
Private m_TrackClickColor2      As OLE_COLOR                 ' track portion clicked second color.

'  default property constants.
Private Const m_def_ArrowDownColor = &H0                  ' black arrow down color.
Private Const m_def_ArrowUpColor = &HE0E0E0               ' light grey arrow up color.
Private Const m_def_BackAngle = 45                        ' horizontal background gradient.
Private Const m_def_BackColor1 = &H606060                 ' darker grey start color.
Private Const m_def_BackColor2 = &HE0E0E0                 ' lighter grey end color.
Private Const m_def_BackMiddleOut = True                  ' middle-out background gradient.
Private Const m_def_BorderColor1 = &H0                    ' black listbox border.
Private Const m_def_BorderColor2 = &H0                    ' black listbox border.
Private Const m_def_BorderMiddleOut = True                ' disabled border gradient middle-out status set.
Private Const m_def_BorderWidth = 1                       ' border width 1 pixel.
Private Const m_def_ButtonColor1 = &H0                    ' black start color.
Private Const m_def_ButtonColor2 = &HC0C0C0               ' grey end color.
Private Const m_def_CheckBoxArrowColor = &H0              ' black check arrow color.
Private Const m_def_CheckBoxColor = &H0                   ' black checkbox border color.
Private Const m_def_CheckStyle = 1                        ' checkmark check style default.
Private Const m_def_CircularGradient = False              ' linear disabled background gradient.
Private Const m_def_CurveBottomLeft = 0                   ' no bottom left curvature.
Private Const m_def_CurveBottomRight = 0                  ' no bottom right curvature.
Private Const m_def_CurveTopLeft = 0                      ' no top left curvature.
Private Const m_def_CurveTopRight = 0                     ' no top right curvature.
Private Const m_def_DblClickBehavior = 1                  ' default is two rapid single clicks.
Private Const m_def_DisArrowDownColor = &HC0C0C0          ' disabled arrow down color. (not used!)
Private Const m_def_DisArrowUpColor = &HC0C0C0            ' disabled arrow up color.
Private Const m_def_DisBackColor1 = &H808080              ' disabled background gradient color 1.
Private Const m_def_DisBackColor2 = &HC0C0C0              ' disabled background gradient color 2.
Private Const m_def_DisBorderColor1 = &H0
Private Const m_def_DisBorderColor2 = &H0
Private Const m_def_DisButtonColor1 = &H404040
Private Const m_def_DisButtonColor2 = &H808080
Private Const m_def_DisCheckboxArrowColor = &H0
Private Const m_def_DisCheckboxColor = &H0
Private Const m_def_DisFocusRectColor = &H808080
Private Const m_def_DisPictureMode = 0
Private Const m_def_DisSelColor1 = &H808080
Private Const m_def_DisSelColor2 = &HC0C0C0
Private Const m_def_DisSelTextColor = &H808080
Private Const m_def_DisTextColor = &H404040
Private Const m_def_DisThumbBorderColor = &H808080
Private Const m_def_DisThumbColor1 = &H404040
Private Const m_def_DisThumbColor2 = &H808080
Private Const m_def_DisTrackbarColor1 = &H808080
Private Const m_def_DisTrackbarColor2 = &HC0C0C0
Private Const m_def_DragEnabled = False                   ' no drag and drop is default.
Private Const m_def_Enabled = True                        ' enabled.
Private Const m_def_FocusBorderColor1 = 0
Private Const m_def_FocusBorderColor2 = 0
Private Const m_def_FocusRectColor = &H0                  ' black focus rectangle.
Private Const m_def_ItemImageSize = 0
Private Const m_def_ListIndex = -1                        ' no items selected.
Private Const m_def_MultiSelect = vbMultiSelectNone       ' one selection at a time.
Private Const m_def_NewIndex = -1                         ' indicating list is empty (no new added).
Private Const m_def_PictureMode = 0                       ' normal picture display (not stretched/tiled).
Private Const m_def_RedrawFlag = True                     ' internal redraw flag to True.
Private Const m_def_RightToLeft = False                   ' no right-to-left by default.
Private Const m_def_SortAsNumeric = False                 ' default sort as string, not numeric.
Private Const m_def_SelCount = 0                          ' read-only selected item counter.
Private Const m_def_SelColor1 = &HC0FFFF                  ' lighter amber selection bar first gradient color.
Private Const m_def_SelColor2 = &HC0FFFF                  ' lighter amber selection bar second gradient color.
Private Const m_def_SelTextColor = &H0                    ' black selected text color.
Private Const m_def_ShowItemImages = False
Private Const m_def_ShowSelectRect = True                 ' show selected listitem focus rectangle.
Private Const m_def_Sorted = True                         ' list sorting by default.
Private Const m_def_Style = vbListBoxStandard             ' no checkbox.
Private Const m_def_Text = ""                             ' no listitem selected by default.
Private Const m_def_TextColor = &H0                       ' black text color.
Private Const m_def_Theme = 2                             ' gunmetal grey default color scheme.
Private Const m_def_ThumbBorderColor = &HE0E0E0           ' light grey thumb border color.
Private Const m_def_ThumbColor1 = &H0                     ' black first thumb color.
Private Const m_def_ThumbColor2 = &H909090                ' medium grey second thumb colot.
Private Const m_def_TopIndex = 0
Private Const m_def_TrackBarColor1 = &H606060             ' darker grey start color.
Private Const m_def_TrackBarColor2 = &HE0E0E0             ' lighter grey end color.
Private Const m_def_TrackClickColor1 = &H0                ' default track portion clicked first color.
Private Const m_def_TrackClickColor2 = &HE0E0E0           ' default track portion clicked second color.

'  events.
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()
Public Event MouseEnter()
Public Event MouseLeave()

Private HasFocus As Boolean       ' master 'control has focus' flag.

'  gradient generation constants.
Private Const RGN_DIFF              As Long = 4
Private Const PI                    As Single = 3.14159265358979
Private Const TO_DEG                As Single = 180 / PI
Private Const TO_RAD                As Single = PI / 180
Private Const INT_ROT               As Long = 1000

'  gradient information for background.
Private BGuBIH                      As BITMAPINFOHEADER
Private BGlBits()                   As Long

'  gradient information for list item selection bar.
Private SeluBIH                     As BITMAPINFOHEADER
Private SellBits()                  As Long

'  gradient information for vertical trackbar.
Private VTrackuBIH                  As BITMAPINFOHEADER
Private VTracklBits()               As Long

'  gradient information for clicked portion of vertical trackbar.
Private vClickTrackuBIH             As BITMAPINFOHEADER
Private vClickTracklBits()          As Long

'  gradient information for scrollbar buttons.
Private TrackButtonuBIH             As BITMAPINFOHEADER
Private TrackButtonlBits()          As Long

'  gradient information for vertical scrollbar thumb.
Private vThumbuBIH                  As BITMAPINFOHEADER
Private vThumblBits()               As Long

'  ************** border segment information **************
'  gradient information for horizontal and vertical border segments.
Private SegV1uBIH                  As BITMAPINFOHEADER
Private SegV1lBits()               As Long
Private SegV2uBIH                  As BITMAPINFOHEADER
Private SegV2lBits()               As Long
Private SegH1uBIH                  As BITMAPINFOHEADER
Private SegH1lBits()               As Long
Private SegH2uBIH                  As BITMAPINFOHEADER
Private SegH2lBits()               As Long

' constants defining the four border segments.
Private Const TOP_SEGMENT         As Long = 0
Private Const RIGHT_SEGMENT       As Long = 1
Private Const BOTTOM_SEGMENT      As Long = 2
Private Const LEFT_SEGMENT        As Long = 3

' holds region pointers for border segments.
Private BorderSegment(0 To 3)     As Long

' declares for virtual horizontal border segment gradient bitmap.
Private VirtualDC_SegH            As Long                    ' handle of the created DC.
Private mMemoryBitmap_SegH        As Long                    ' handle of the created bitmap.
Private mOriginalBitmap_SegH      As Long                    ' used in destroying virtual DC.

' declares for virtual vertical border segment gradient bitmap.
Private VirtualDC_SegV            As Long                    ' handle of the created DC.
Private mMemoryBitmap_SegV        As Long                    ' handle of the created bitmap.
Private mOriginalBitmap_SegV      As Long                    ' used in destroying virtual DC.
'**********************************************************

Private Const ScrollBarButtonHeight As Long = 15          ' the height, in pixels, of scrollbar button.
Private Const ScrollBarButtonWidth  As Long = 15          ' the width, in pixels, of scrollbar button.
Private vScrollTrackHeight          As Long               ' the height, in pixels, of thumb scroll track.
Private Const vScrollMinThumbHeight As Long = 9           ' keep it an odd number so middle isn't between pixels.
Private VerticalScrollBarActive     As Boolean            ' indicates scrollbar is drawn.
Private vThumbHeight                As Long               ' the height in pixels of the vertical thumb.
Private RecalculateThumbHeight      As Boolean            ' prevents unnecessary thumb height recalulations.
Private ThumbYPos                   As Long               ' y coordinate of top of scrollbar thumb.

' structure for containing exact scrollbar component location info for mouseover.
Private Type ScrollBarLocationType
   UpButtonLocation                 As RECT
   DownButtonLocation               As RECT
   ScrollTrackLocation              As RECT
   ScrollThumbLocation              As RECT
End Type
Private vScrollBarLocation As ScrollBarLocationType

' keyboard, mouse, and list item tracking variables.
Private ItemWithFocus                    As Long          ' the item in the list that has the "virtual focus".
Private CtrlKeyDown                      As Boolean       ' global "control key being pressed" flag.
Private ShiftKeyDown                     As Boolean       ' global "shift key is being pressed" flag.
Private LastSelectedItem                 As Long          ' item last clicked or otherwise selected.
Private RightClickFlag                   As Boolean       ' doubleclick bypass (DblClick detects right click too)
Private MouseAction                      As Long          ' set to value of one of below constants.
Private Const MOUSE_NOACTION             As Long = 0      ' mouse button is not down.
Private Const MOUSE_DOWNED_IN_LIST       As Long = 1      ' mouse downed in text portion of listbox.
Private Const MOUSE_DOWNED_IN_UPPERTRACK As Long = 2      ' mouse downed in trackbar above thumb.
Private Const MOUSE_DOWNED_IN_LOWERTRACK As Long = 3      ' mouse downed in trackbar below thumb.
Private Const MOUSE_DOWNED_IN_DOWNBUTTON As Long = 4      ' mouse downed in down scrollbar button.
Private Const MOUSE_DOWNED_IN_UPBUTTON   As Long = 5      ' mouse downed in up scrollbar button.
Private Const MOUSE_DOWNED_IN_THUMB      As Long = 6      ' mouse downed in scrollbar thumb.

Private FirstExtendedSelection      As Long               ' in Extended mode, the original list item clicked on.
Private LastExtendedSelection       As Long               ' in Extended mode, the last list item clicked on.
Private ItemMouseIsIn               As Long               ' to prevent redraws when mouse moves in same list item.
Private ShiftDownStartItem          As Long               ' item with focus when shift key is pressed.

' flags indicating how list items should be drawn.
Private Const DrawAsSelected        As Boolean = True     ' draw list item with selection bar gradient.
Private Const DrawAsUnselected      As Boolean = False    ' draw list item without selection bar gradient.
Private Const FocusRectangleYes     As Boolean = True     ' draw item with focus rectangle.
Private Const FocusRectangleNo      As Boolean = False    ' draw item without focus rectangle.
Private Const KeepSelectionAsIs     As Boolean = True     ' draw item, keeping item's selection status.
Private Const KeepSelectionNo       As Boolean = False    ' draw item, don't keep item's selection status.

' the arrays for the properties .List, .ItemData and .Selected.
Private ListArray()                 As String             ' tied to the .List property.
Private ItemDataArray()             As Long               ' tied to the .ItemData property.
Private SelectedArray()             As Boolean            ' tied to the .Selected property.
Private ImageIndexArray()           As Long               ' tied to the .ImageIndex property.

Private Images()                    As StdPicture         ' array that holds the images that are displayed by listitems.
Private ImageCount                  As Long               ' total number of stored listitem images.
Private PicX                        As Long               ' x coordinate of listitem image.

Private lBarWid                     As Long
Private SelBarOffset                As Long

' for keeping track of where the mouse is at any given time.
Private Const OVER_BORDER           As Long = 0           ' not used at this time.
Private Const OVER_LIST             As Long = 1           ' mouse cursor is over list portion of control.
Private Const OVER_UPBUTTON         As Long = 2           ' mouse cursor is over vertical scrollbar up button.
Private Const OVER_DOWNBUTTON       As Long = 3           ' mouse cursor is over vertical scrollbar down button.
Private Const OVER_VTRACKBAR        As Long = 4           ' mouse cursor is over vertical scrollbar trackbar.
Private Const OVER_VTHUMB           As Long = 5           ' mouse cursor is over vertical scrollbar thumb.
Private MouseLocation               As Long               ' set to one of the above constants.
Private MouseOverCheckBox           As Boolean            ' set to True if mouse is in checkbox display area.

Private DragFlag                    As Boolean            ' master 'is drag enabled?' flag.

' for scrollbar button display.
Private Const UPBUTTON              As Long = 1           ' display the vertical scrollbar up button.
Private Const DOWNBUTTON            As Long = 2           ' display the vertical scrollbar up button.

' for space bar selection of list items.
Private Const SPACEBAR              As Long = 32          ' space bar is chr(32).

'  the pixel range for the center of the vertical scrollbar
'  thumb as it goes up and down the scroller track.
Private Type vThumbRangeType
   Top                              As Long               ' top pixel position of middle of thumb.
   Bottom                           As Long               ' bottom pixel position of middle of thumb.
End Type
Private vThumbRange                 As vThumbRangeType

' thumb scroll tracking variables and constants.
Private MouseX                      As Single             ' global mouse X position variable.
Private MouseY                      As Single             ' global mouse Y position variable.
Private MouseDownYPos               As Single             ' mouse y position when mouse is clicked down.
Private Const SCROLL_TICKCOUNT      As Long = 50          ' scroll time delay interval.
Private Const INITIAL_SCROLL_DELAY  As Long = 400         ' delay before scrolling commences.
Private DraggingVThumb              As Boolean            ' flag indicating mouse is down on thumb..
Private ThumbScrolling              As Boolean            ' flag indicating thumb scrolling is now in progress.
Private MousePosInVThumb            As Single             ' distance in pixels from top of vertical thumb.
Private ScrollFlag                  As Boolean            ' mouse down on list, then moved above/below list.
Private Const SCROLL_LISTDOWN       As Long = 1           ' display range increment for scrolling list down.
Private Const SCROLL_LISTUP         As Long = -1          ' display range increment for scrolling list up.

' declares for virtual listbox background bitmap.
Private VirtualBackgroundDC         As Long               ' DC handle of the created Device Context
Private mMemoryBitmap               As Long               ' Handle of the created bitmap
Private mOrginalBitmap              As Long               ' Used in Destroy

Private ListIsSorted                As Boolean            ' flag to indicate listitems are in ascending sorted order.

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

'*************************************************************************
'* processes the intercepted windows messages.  This subclass handler    *
'* MUST be the first Public routine in this file.  This includes         *
'* public properties also.  Other subclass routines at bottom of code.   *
'*************************************************************************

   Select Case uMsg

      Case WM_MOUSEWHEEL
         If HasFocus Then ' may want to take focus check out - regular vb textbox wheels w/o focus.
             Select Case wParam
             Case Is > False
                ProcessUpButton
             Case Else
                ProcessDownButton
             End Select
         End If

      Case WM_MOUSEMOVE
'        detect when mouse has entered the control.
         If m_Enabled And Not bInCtrl Then
            bInCtrl = True
            Call TrackMouseLeave(lng_hWnd)
            RaiseEvent MouseEnter
         End If

'     detect when mouse has left the control.
      Case WM_MOUSELEAVE
         bInCtrl = False
         RaiseEvent MouseLeave

'     detect when control has gained the focus.
      Case WM_SETFOCUS
         If m_Enabled Then
            HasFocus = True
            m_ActiveBorderColor1 = m_FocusBorderColor1
            m_ActiveBorderColor2 = m_FocusBorderColor2
            InitBorder
            CreateBorder
            If m_Style = [Standard] Then
               DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
            Else
               DisplayListBoxItem ItemWithFocus, DrawAsSelected, FocusRectangleYes
            End If
            UserControl.Refresh
         End If

'     detect when control has lost the focus.
      Case WM_KILLFOCUS
         If m_Enabled Then
            HasFocus = False
            m_ActiveBorderColor1 = m_BorderColor1
            m_ActiveBorderColor2 = m_BorderColor2
            InitBorder
            CreateBorder
            MouseAction = MOUSE_NOACTION
            ShiftKeyDown = False
            CtrlKeyDown = False
'           since listbox has lost the focus, display focused listbox item without the focus rectangle.
            If m_Style = [Standard] Then
               DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
            Else
               DisplayListBoxItem ItemWithFocus, DrawAsSelected, FocusRectangleNo
            End If
            UserControl.Refresh
         End If

   End Select

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_Initialize()

'*************************************************************************
'* the first event in the control's life cycle.                          *
'*************************************************************************

   Dim OS As OSVERSIONINFO

   ReDim ListArray(0)       ' the array tied to the .List property.
   ReDim ItemDataArray(0)   ' the array tied to the .ItemData property.
   ReDim SelectedArray(0)   ' the array tied to the .Selected property.
   ReDim ImageIndexArray(0) ' the array tied to the .ImageIndex property.

'  initialize property and internal variables.
   m_ListCount = 0
   LastSelectedItem = -1
   m_ListIndex = -1
   ItemWithFocus = 0
   ItemMouseIsIn = -1
   m_hBrush = 0             ' bitmap tiling pattern brush.
   RecalculateThumbHeight = True

'  get the operating system version for text drawing purposes.
   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)
   mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_Click()

'*************************************************************************
'* return a click event.  A VB listbox only returns a click event when   *
'* the mouse cursor is over the populated portion of list area; this be- *
'* havior is emulated here by utilizing the MouseOverIndex function.     *
'*************************************************************************

   If m_Enabled And Not RightClickFlag And MouseOverIndex(MouseY) <> -1 Then
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_DblClick()

'*************************************************************************
'* process a double-click event as normal, or as two rapid single click  *
'* events.  Why offer a choice?  You know how a standard VB button       *
'* interprets double-clicks as two rapid button presses?  I like that    *
'* quick responsiveness and incorporate it into controls that I don't    *
'* need doubleclick functionality for.  However, in a listbox I can see  *
'* where double-clicking a list item might be a desirable feature for    *
'* some, so I provided both options via the DblClickBehavior property.   *
'* Even if normal doubleclicks are permitted, they are still treated as  *
'* two rapid single-clicks if mouse is in vertical scrollbar area or the *
'* listbox is in CheckBox mode (to help emulate vb listbox).             *
'*************************************************************************

   If (m_DblClickBehavior = [Two Single Clicks]) Or (Not IsInList(MouseX, MouseY)) Or _
      (m_Style = [CheckBox] And IsInList(MouseX, MouseY)) Then
      If m_Enabled And Not RightClickFlag Then
'        originally I just sent control to the UserControl_MouseDown routine.  But I found that when
'        double-clicking (keeping second click held down), then drag-scrolling the list, the list would
'        not scroll.  This is because MouseMove would not fire when mouse cursor left the list in
'        this scenario. So I generate the actual mousedown event via this API to solve the problem.
         mouse_event MOUSEEVENTF_LEFTDOWN, MouseX, MouseY, 0, 0
      End If
   Else
      RaiseEvent DblClick
   End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'************************************************************************
'* processes a clicked item or vertical scrollbar component.            *
'************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   m_DragEnabled = False

   If Button = vbRightButton Then
      RightClickFlag = True
      RaiseEvent MouseDown(Button, Shift, X, Y)
      Exit Sub
   End If

'  added 21 May 2006.  Helps in combining MorphListBox with a standard VB popup menu.
'  Thanks to Paul Bahlawan for happening upon a bug that this line of code fixes.
   MouseY = Y

   RightClickFlag = False

   Select Case MouseLocation

      Case OVER_UPBUTTON
'        process click of vertical scrollbar up button.
         MouseAction = MOUSE_DOWNED_IN_UPBUTTON
         ProcessUpButton

      Case OVER_DOWNBUTTON
'        process click of vertical scrollbar down button.
         MouseAction = MOUSE_DOWNED_IN_DOWNBUTTON
         ProcessDownButton

      Case OVER_VTRACKBAR
'        process a page up or page down based on mouse-click of vertical scrollbar in relation to thumb.
         If MouseCursorIsAboveThumb(Y) Then
            MouseAction = MOUSE_DOWNED_IN_UPPERTRACK
            ProcessPageUp
         Else
            MouseAction = MOUSE_DOWNED_IN_LOWERTRACK
            ProcessPageDown
         End If

      Case OVER_VTHUMB
'        initiate scolling of list via dragging of vertical scrollbar thumb.
         MouseAction = MOUSE_DOWNED_IN_THUMB
         DraggingVThumb = True
         MouseDownYPos = Y
         ProcessVThumbScroll

      Case OVER_LIST
'        make sure mouse pointer is in populated area of listbox before continuing.
         If MouseOverIndex(Y) <> -1 Then
            m_DragEnabled = DragFlag ' set to original user-selected property state.
'           process selection or deselection of a list item.
            MouseAction = MOUSE_DOWNED_IN_LIST
            If m_Style = [CheckBox] Then
               ProcessMouseDown_CheckBoxMode
            Else
               Select Case m_MultiSelect
                  Case vbMultiSelectNone
                     ProcessMouseDown_MultiSelectNone
                  Case vbMultiSelectSimple
                     ProcessMouseDown_MultiSelectSimple
                  Case vbMultiSelectExtended
                     ProcessMouseDown_MultiSelectExtended
               End Select
            End If
         End If

   End Select

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* processes mouse up event.                                             *
'*************************************************************************

   If m_Enabled Then

      If Button = vbLeftButton Then

         Select Case MouseAction
            Case MOUSE_DOWNED_IN_UPBUTTON
'              if mouse was down on one of the scrollbar buttons, redisplay that button in its 'up' colors.
               MouseAction = MOUSE_NOACTION
               DisplayTrackBarButton UPBUTTON
            Case MOUSE_DOWNED_IN_DOWNBUTTON
               MouseAction = MOUSE_NOACTION
               DisplayTrackBarButton DOWNBUTTON
            Case MOUSE_DOWNED_IN_UPPERTRACK, MOUSE_DOWNED_IN_LOWERTRACK
'              if mouse was down on the scrollbar track, redisplay it so clicked portion is normal color.
               MouseAction = MOUSE_NOACTION
               DisplayVerticalScrollBar
            Case Else
               MouseAction = MOUSE_NOACTION
               m_DragEnabled = False
         End Select

         UserControl.Refresh
         DraggingVThumb = False
         RaiseEvent MouseUp(Button, Shift, X, Y)

      Else

'        raise the mouseup event for right mouse button.
         RaiseEvent MouseUp(Button, Shift, X, Y)

      End If

   End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* processes mouse movement based on .MultiSelect property.              *
'*************************************************************************

   Dim DidIt As Boolean    ' flag that lets routine know an action was performed.

'  set the global cursor coordinate variables.
   MouseX = X
   MouseY = Y

'  determine which component of the listbox the mouse is over (list, scrollbar, thumb, buttons, track).
   DetermineMouseLocation X, Y

'  check for and process possible drag scrolling.
   ProcessMouseDragScrolling DidIt
   If DidIt Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  check for and process possible dragging of scrollbar out of range.
   ProcessMouseDragThumbOutOfRange DidIt
   If DidIt Then
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  if user was scrolling with the thumb (started in MouseDown), make
'  sure scrolling can continue even if mouse has moved off the thumb.
   If MouseAction <> MOUSE_NOACTION And DraggingVThumb And Not ThumbScrolling Then
      ProcessVThumbScroll
      RaiseEvent MouseMove(Button, Shift, X, Y)
      Exit Sub
   End If

'  if we're dragging mouse over list portion of control, process
'  it.  Ignore this if drag and drop operation is enabled.
   If MouseAction = MOUSE_DOWNED_IN_LIST And MouseLocation = OVER_LIST And Not m_DragEnabled Then
      ProcessMouseMoveItemSelection
   End If

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* allows user to navigate listbox via keyboard.                         *
'*************************************************************************

   If m_Enabled Then

'     determine Shift and Ctrl key status.
      ShiftKeyDown = (Shift And vbShiftMask) > 0
      If ShiftKeyDown Then
         ShiftDownStartItem = ItemWithFocus
      End If

      CtrlKeyDown = (Shift And vbCtrlMask) > 0

'     process the appropriate key.
      Select Case KeyCode
         Case vbKeyPageDown
            ProcessPageDownKey
         Case vbKeyPageUp
            ProcessPageUpKey
         Case vbKeyEnd
            ProcessEndKey
         Case vbKeyHome
            ProcessHomeKey
         Case vbKeyUp, vbKeyLeft
            ProcessUpArrowKey
         Case vbKeyDown, vbKeyRight
            ProcessDownArrowKey
         Case SPACEBAR
            ProcessSpaceBar
      End Select

      RaiseEvent KeyDown(KeyCode, Shift)

   End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   
'*************************************************************************
'* processes keypress event.                                             *
'*************************************************************************

   If m_Enabled Then
      RaiseEvent KeyPress(KeyAscii)
   End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* processes key up event.                                               *
'*************************************************************************

   If m_Enabled Then
'     determine shift and Ctrl key status.
      ShiftKeyDown = (Shift And vbShiftMask) > 0
      CtrlKeyDown = (Shift And vbCtrlMask) > 0
      RaiseEvent KeyUp(KeyCode, Shift)
   End If

End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* currently only used in design mode.                                   *
'*************************************************************************

'  when a new listbox is drawn onto the form in design mode, the event sequence is
'  Initialize-Show instead of Initialize-ReadProperties-Show.  This means we have to
'  calculate the gradient info from the defaults, as opposed to reading properties.
   CalculateGradients
   RedrawControl
   RaiseEvent Resize

End Sub

Private Sub UserControl_Terminate()

'*************************************************************************
'* restores memory used by listbox and stops subclassing.                *
'*************************************************************************

   On Error GoTo Catch

'  deallocate property arrays.
   Erase ListArray
   Erase ItemDataArray
   Erase SelectedArray
   Erase ImageIndexArray

'  destroy the virtual DC's used in background storage.
   DestroyVirtualDC VirtualBackgroundDC, mMemoryBitmap, mOrginalBitmap

'  destroy bitmap tiling pattern.
   DestroyPattern

'  destroy border segments.
   DestroyVirtualDC VirtualDC_SegH, mMemoryBitmap_SegH, mOriginalBitmap_SegH
   DestroyVirtualDC VirtualDC_SegV, mMemoryBitmap_SegV, mOriginalBitmap_SegV

'  halt subclassing.
   Call Subclass_StopAll

Catch:

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<< Mouse Processing Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ProcessMouseDown_CheckBoxMode()

'*************************************************************************
'* processes a mouse down in the list in Checkbox Style mode.            *
'*************************************************************************

   If LastSelectedItem = -1 Then
      LastSelectedItem = MouseOverIndex(MouseY)
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
   End If

   If Not MouseOverCheckBox Then

      If MouseOverIndex(MouseY) = LastSelectedItem Then
'        if the mouse is over the list (not a checkbox), and we are clicking on the
'        item that is already the focused item, then reverse its selection status and exit.
         SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
         If SelectedArray(LastSelectedItem) Then
            m_SelCount = m_SelCount + 1
         Else
            m_SelCount = m_SelCount - 1
         End If
         DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
         UserControl.Refresh
         Exit Sub
      Else
'        if the mouse is over the list (not a checkbox), and we are NOT clicking on the
'        item that is already the focused item, then set the focus and selection gradient
'        to the new item and exit.  No selected status changes are made.
         DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
         LastSelectedItem = MouseOverIndex(MouseY)
         ItemWithFocus = LastSelectedItem
         m_ListIndex = ItemWithFocus
         ItemMouseIsIn = LastSelectedItem
'        redisplay the newly selected item.
         DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
         UserControl.Refresh
         Exit Sub
      End If

   Else

'     if the mouse is clicked in a list item's checkbox, that item's selection status
'     is immediately reversed, and the selection gradient moves to that list item.

'     display the previous list item without selection bar.
      DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      LastSelectedItem = MouseOverIndex(MouseY)
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      ItemMouseIsIn = LastSelectedItem
      SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
      If SelectedArray(LastSelectedItem) Then
         m_SelCount = m_SelCount + 1
      Else
         m_SelCount = m_SelCount - 1
      End If

'     redisplay the newly selected item.
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      UserControl.Refresh

   End If

End Sub

Private Sub ProcessMouseDown_MultiSelectNone()

'*************************************************************************
'* processes a mouse down in the list in MultiSelect None mode.          *
'*************************************************************************

'  repaint previously selected list item without selection gradient / focus rectangle.
   ClearPreviousSelection KeepSelectionNo

'  make the newly clicked item the last selected item.
   LastSelectedItem = MouseOverIndex(MouseY)
   ProcessSelectedItem

End Sub

Private Sub ProcessMouseDown_MultiSelectSimple()

'*************************************************************************
'* processes a mouse down in the list in MultiSelect Simple mode.        *
'*************************************************************************

'  repaint previously selected list item, keeping selection status but not focus rectangle.
   ClearPreviousSelection KeepSelectionAsIs

   LastSelectedItem = MouseOverIndex(MouseY) ' even if we're deselecting with the mouse click?
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  reverse the selection status of the list item that was clicked.
   If SelectedArray(LastSelectedItem) Then
      SelectedArray(LastSelectedItem) = False
      DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleYes
      m_SelCount = m_SelCount - 1
   Else
      SelectedArray(LastSelectedItem) = True
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      m_SelCount = m_SelCount + 1
   End If

   UserControl.Refresh

End Sub

Private Sub ProcessMouseDown_MultiSelectExtended()

'*************************************************************************
'* processes a mouse down in the list in MultiSelect None mode.          *
'*************************************************************************

'  make the newly clicked item the last selected item.
   LastSelectedItem = MouseOverIndex(MouseY)

'  this item must also have the focus rectangle.
   ItemWithFocus = LastSelectedItem

'  in Extended mode, the .ListIndex property is always the item with the focus.
   m_ListIndex = ItemWithFocus

'  if the Ctrl or Shift keys are not down, set the entire Selected array to False.
   If (Not CtrlKeyDown) And (Not ShiftKeyDown) Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

'  in MultiSelect Extended mode, a mouse down just selects one item unless shift is pressed.
   FirstExtendedSelection = LastSelectedItem
   LastExtendedSelection = LastSelectedItem

'  make sure item is selected.  If Ctrl key is pressed, flip the selection status.
   If (Not CtrlKeyDown) Then
      SelectedArray(LastSelectedItem) = True
   Else
      SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
   End If

'  if Shift is down, all items from the item that had the focus when
'  shift was pressed to the new item that has the focus are selected.
   If ShiftKeyDown Then
      SetSelectedArrayRange ShiftDownStartItem, ItemWithFocus, True
      CalculateSelCount
   End If

'  display whole list instead of just item - other displayed items' selection status may have changed.
   DisplayList

'  used in MouseMove to detect whether mouse is still in the same list item as it is now.
   ItemMouseIsIn = LastSelectedItem

'  since mouse is being clicked down, SelCount is always 1 if Ctrl or Shift keys not pressed.
   If (Not CtrlKeyDown) And (Not ShiftKeyDown) Then
      m_SelCount = 1
   Else
      If Not ShiftKeyDown Then
         If SelectedArray(LastSelectedItem) And Not ShiftKeyDown Then
            m_SelCount = m_SelCount + 1
         Else
            m_SelCount = m_SelCount - 1
         End If
      End If
   End If

End Sub

Private Sub ProcessDownButton()

'*************************************************************************
'* shifts displayed list items down on scrollbar down arrow button click.*
'*************************************************************************

'  only do this if last list item is not already displayed.
   If Not InDisplayedItemRange(m_ListCount - 1) Then
      DisplayRange.FirstListItem = DisplayRange.FirstListItem + 1
      DisplayRange.LastListItem = DisplayRange.LastListItem + 1
      DisplayList
'     check for and process possible continuous scroll (i.e. mouse button held down).
      ProcessContinuousScroll SCROLL_LISTDOWN
   End If

End Sub

Private Sub ProcessUpButton()

'*************************************************************************
'* shifts displayed list items up on scrollbar up arrow button click.    *
'*************************************************************************

'  only do this if first list item is not already displayed.
   If Not InDisplayedItemRange(0) Then
      DisplayRange.FirstListItem = DisplayRange.FirstListItem - 1
      DisplayRange.LastListItem = DisplayRange.LastListItem - 1
      DisplayList
'     check for and process possible continuous scroll (i.e. mouse arrow button held down).
      ProcessContinuousScroll SCROLL_LISTUP
   End If

End Sub

Private Sub ProcessPageUp()

'*************************************************************************
'* shifts displayed list items up one page when mouse is clicked above   *
'* vertical scroll thumb.                                                *
'*************************************************************************

'  only perform a page up if first page is not already displayed.
   If Not InDisplayedItemRange(0) Then

'     adjust the displayed item range.
      DisplayRange.FirstListItem = DisplayRange.FirstListItem - MaxDisplayItems + 1
      If DisplayRange.FirstListItem < 0 Then
         DisplayRange.FirstListItem = 0
      End If
      If DisplayRange.FirstListItem + MaxDisplayItems - 1 > m_ListCount Then
         DisplayRange.LastListItem = m_ListCount - 1
      Else
         DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
      End If

      DisplayList

'     check for and process possible continuous scroll (i.e. mouse button held down).
      ProcessContinuousScroll -MaxDisplayItems

   End If

End Sub

Private Sub ProcessPageDown()

'*************************************************************************
'* shifts displayed list items down one page when mouse is clicked below *
'* vertical scroll thumb.                                                *
'*************************************************************************

'  only perform a page down if last page is not already being displayed.
   If Not InDisplayedItemRange(m_ListCount - 1) Then

'     adjust the displayed item range.
      If DisplayRange.LastListItem + MaxDisplayItems - 1 <= m_ListCount - 1 Then
         DisplayRange.FirstListItem = DisplayRange.LastListItem
      Else
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
      End If
      DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems
      If DisplayRange.LastListItem > m_ListCount - 1 Then
         DisplayRange.LastListItem = m_ListCount - 1
      End If

      DisplayList

'     check for and process possible continuous scroll (i.e. mouse button held down).
      ProcessContinuousScroll MaxDisplayItems

   End If

End Sub

Private Sub ClearPreviousSelection(SelectionSaveIndicator As Boolean)

'*************************************************************************
'* clears the listbox of last selected item's gradient (if the Style is  *
'* set to None or Extended) and erases the focus rectangle in whatever   *
'* list item possesses it.  Only happens if list item(s) are displayed.  *
'*************************************************************************

'  repaints the selected item without the focus rectangle (saving gradient selection
'  highlight), or redisplays as unselected without the focus rectangle
'  (depends on SelectionSaveStatus parameter).
   If LastSelectedItem <> -1 And InDisplayedItemRange(LastSelectedItem) Then
      Select Case SelectionSaveIndicator
        Case KeepSelectionAsIs
           DisplayListBoxItem LastSelectedItem, SelectedArray(LastSelectedItem), FocusRectangleNo
        Case KeepSelectionNo
           DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      End Select
   End If

'  make sure the item with the focus rectangle is 'de-rectangled'.
'  Selected or deselected appearance of item is unchanged.
   If ItemWithFocus <> LastSelectedItem Then
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
   End If

End Sub

Private Function InDisplayedItemRange(Index As Long) As Boolean

'*************************************************************************
'* returns whether a given list item index is in the displayed range.    *
'*************************************************************************

   If Index >= DisplayRange.FirstListItem And Index <= DisplayRange.LastListItem Then
      InDisplayedItemRange = True
   End If

End Function

Private Sub ProcessMouseDragScrolling(DidIt As Boolean)

'*************************************************************************
'* this allows user to scroll through the list by clicking within the    *
'* list area and then dragging the mouse above or below the listbox,     *
'* like a regular vb listbox.                                            *
'*************************************************************************

   If ScrollFlag And MouseY >= 0 And MouseY <= ScaleHeight Then
      ScrollFlag = False
      DidIt = False
   ElseIf MouseAction <> MOUSE_NOACTION And Not DraggingVThumb And Not ScrollFlag Then
      If MouseY > ScaleHeight - m_BorderWidth Then    ' start drag scrolling when border is reached.
         ScrollFlag = True
         ProcessContinuousScroll SCROLL_LISTDOWN
         ScrollFlag = False
         DidIt = True
      ElseIf MouseY < m_BorderWidth - 1 Then          ' start drag scrolling when border is reached.
         ScrollFlag = True
         ProcessContinuousScroll SCROLL_LISTUP
         ScrollFlag = False
         DidIt = True
      End If
   End If

End Sub

Private Sub ProcessMouseMoveItemSelection()

'*************************************************************************
'* controls selection of items by mouse drag in all listbox states.      *
'*************************************************************************

'  first, make sure the mouse hasn't just been moved within the same list
'  item it was in the last time the mouse was moved.  This avoids unnecessary
'  processing and graphics redraws.
   If ItemMouseIsIn = MouseOverIndex(MouseY) Then
      Exit Sub
   Else
      ItemMouseIsIn = MouseOverIndex(MouseY)
   End If

   If m_Style = [CheckBox] Then
      ProcessMouseMoveItemSelection_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessMouseMoveItemSelection_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessMouseMoveItemSelection_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessMouseMoveItemSelection_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessMouseMoveItemSelection_CheckBoxMode()

'*************************************************************************
'* processes selecting items by mouse drag in CheckBox mode.             *
'*************************************************************************

'  display the previously selected item as unselected, with no focus rectangle.
   DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo

'  the newly moved-over item is now the last item with the selection bar.
   LastSelectedItem = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If LastSelectedItem = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      LastSelectedItem = m_ListCount - 1
   End If

   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  display the item with the selection bar.
   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessMouseMoveItemSelection_MultiSelectNone()

'*************************************************************************
'* processes selecting items by mouse drag in MultiSelect None mode.     *
'*************************************************************************

'  display the previously selected item as unselected, with no focus rectangle.
   DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo

'  the newly moved-over item is now the last item selected.
   LastSelectedItem = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If LastSelectedItem = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      LastSelectedItem = m_ListCount - 1
   End If

   ProcessSelectedItem

End Sub

Private Sub ProcessMouseMoveItemSelection_MultiSelectSimple()

'*************************************************************************
'* processes selecting items by mouse drag in MultiSelect Simple mode.   *
'*************************************************************************

'  get rid of the focus rectangle in the previously focused item, keeping selection status.
   DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo

'  set the moused-over item as the new item with focus.
   ItemWithFocus = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If ItemWithFocus = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      ItemWithFocus = m_ListCount - 1
   End If

'  in MultiSelect Simple mode, the .ListIndex property is always the item with the focus.
   m_ListIndex = ItemWithFocus

'  paint the focus rectangle over the moused-over item, keeping selection status as-is.
   DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessMouseMoveItemSelection_MultiSelectExtended()

'*************************************************************************
'* processes selecting items by mouse drag in MultiSelect Extended mode. *
'* when mouse is moved in MultiSelect Extended mode then the item's      *
'* selection status is reversed (unless it's the originally selected     *
'* item, in which case it stays selected.)                               *
'*************************************************************************

'  get rid of the focus rectangle in the previously focused item, keeping selection status.
   DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
   UserControl.Refresh

'  if Shift or Ctrl not pressed, clear all selected items so that redisplay of list is handled correctly.
   If (Not ShiftKeyDown) And (Not CtrlKeyDown) Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

'  set the moused-over item as the new item with focus.
   LastSelectedItem = MouseOverIndex(MouseY)

'  if mouse is moved really fast past the end of a less than 1 page list the last item is
'  not always highlighted because the mousemove didn't fire.  This helps correct that.
   If LastSelectedItem = -1 Or (m_ListCount < MaxDisplayItems And MouseY > ScaleHeight) Then
      LastSelectedItem = m_ListCount - 1
   End If

'  focus rectangle and .ListIndex property are always set to last selected item in this mode.
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  set selection status of all items from first to last to True.
   LastExtendedSelection = ItemWithFocus

'  set the selected range.
   SetSelectedArrayRange FirstExtendedSelection, LastExtendedSelection, True

'  determine the number of selected items.
   If (Not ShiftKeyDown) And (Not CtrlKeyDown) Then
      m_SelCount = Abs(LastExtendedSelection - FirstExtendedSelection) + 1
   Else
      CalculateSelCount
   End If

   DisplayList

End Sub

Private Sub DetermineMouseLocation(X As Single, Y As Single)

'*************************************************************************
'* sets the MouseLocation variable based on which listbox component the  *
'* mouse cursor is at the time of the call to this routine.              *
'*************************************************************************

   If IsInList(X, Y) Then
      MouseLocation = OVER_LIST
   ElseIf IsInVerticalThumb(X, Y) Then
      MouseLocation = OVER_VTHUMB
   ElseIf IsInVerticalTrackbar(X, Y) Then
      MouseLocation = OVER_VTRACKBAR
   ElseIf IsInUpButton(X, Y) Then
      MouseLocation = OVER_UPBUTTON
   ElseIf IsInDownButton(X, Y) Then
      MouseLocation = OVER_DOWNBUTTON
   Else
      MouseLocation = OVER_BORDER
   End If

'  check to see if mouse is over a checkbox if in CheckBox mode.  Account for .RightToLeft.
   If m_Style = [CheckBox] And MouseLocation = OVER_LIST Then
      If Not m_RightToLeft Then
         If X >= m_BorderWidth + 3 And X <= m_BorderWidth + 18 Then
            MouseOverCheckBox = True
         Else
            MouseOverCheckBox = False
         End If
      Else
         If X >= ScaleWidth - m_BorderWidth - 17 And X <= ScaleWidth - m_BorderWidth - 4 Then
            MouseOverCheckBox = True
         Else
            MouseOverCheckBox = False
         End If
      End If
   End If

End Sub

Private Function IsInList(XPos As Single, YPos As Single) As Boolean

'*************************************************************************
'* returns True if mouse cursor is in list display portion of control.   *
'*************************************************************************

   Dim ListBorder As Long   ' right or left edge of list area, depending on .RightToLeft orientation.

'  account for .RightToLeft and also possible scrollbar being displayed.
   If Not m_RightToLeft Then ' determine rightmost part of list area.
      If VerticalScrollBarActive Then
         ListBorder = ScaleWidth - m_BorderWidth - ScrollBarButtonWidth - 1
      Else
         ListBorder = ScaleWidth - m_BorderWidth - 1
      End If
      If XPos >= m_BorderWidth And _
         XPos <= ListBorder And _
         YPos >= m_BorderWidth And _
         YPos <= ScaleHeight - m_BorderWidth - 1 Then
            IsInList = True
      End If
   Else ' determine leftmost part of list area.
      If VerticalScrollBarActive Then
         ListBorder = m_BorderWidth + ScrollBarButtonWidth
      Else
         ListBorder = m_BorderWidth
      End If
      If XPos >= ListBorder And _
         XPos <= ScaleWidth - m_BorderWidth - 1 And _
         YPos >= m_BorderWidth And _
         YPos <= ScaleHeight - m_BorderWidth - 1 Then
            IsInList = True
      End If
   End If


End Function

Private Function IsInVerticalScrollbar(XPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in any part of the vertical scrollbar.*
'*************************************************************************

   If VerticalScrollBarActive Then
      If XPos >= vScrollBarLocation.ScrollTrackLocation.Left And _
         XPos <= vScrollBarLocation.ScrollTrackLocation.Right Then
            IsInVerticalScrollbar = True
      End If
   End If

End Function

Private Function IsInVerticalThumb(XPos As Single, YPos As Single) As Boolean

'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar thumb.          *
'*************************************************************************

   If VerticalScrollBarActive Then
      If YPos >= vScrollBarLocation.ScrollThumbLocation.Top And _
         YPos <= vScrollBarLocation.ScrollThumbLocation.Bottom And _
         XPos >= vScrollBarLocation.ScrollThumbLocation.Left And _
         XPos <= vScrollBarLocation.ScrollThumbLocation.Right Then
            IsInVerticalThumb = True
      End If
   End If

End Function

Private Function IsInVerticalTrackbar(XPos As Single, YPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar trackbar.       *
'*************************************************************************

   If IsInVerticalScrollbar(XPos) Then
      If YPos >= vScrollBarLocation.ScrollTrackLocation.Top And _
         YPos <= vScrollBarLocation.ScrollTrackLocation.Bottom And _
         Not IsInVerticalThumb(XPos, YPos) Then
            IsInVerticalTrackbar = True
      End If
   End If

End Function

Private Function IsInUpButton(XPos As Single, YPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar up button.      *
'*************************************************************************

   If IsInVerticalScrollbar(XPos) Then
      If YPos >= vScrollBarLocation.UpButtonLocation.Top And _
         YPos <= vScrollBarLocation.UpButtonLocation.Bottom Then
            IsInUpButton = True
      End If
   End If

End Function

Private Function IsInDownButton(XPos As Single, YPos As Single) As Boolean
      
'*************************************************************************
'* returns True if mouse cursor is in vertical scrollbar down button.    *
'*************************************************************************

   If IsInVerticalScrollbar(XPos) Then
      If YPos >= vScrollBarLocation.DownButtonLocation.Top And _
         YPos <= vScrollBarLocation.DownButtonLocation.Bottom Then
            IsInDownButton = True
      End If
   End If

End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Keyboard Processing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ProcessPageUpKey()

'*************************************************************************
'* processes page up key for all listbox states.                         *
'* page down ALWAYS starts from the item with the focus rectangle, even  *
'* if that item is not currently in the display range.  After the page   *
'* up the item that formerly had focus rect will be last displayed entry *
'* (unless said entry was less than "MaxDisplayItems" entries below top  *
'* item in list)                                                         *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessPageUpKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessPageUpKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessPageUpKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessPageUpKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessPageUpKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, PgDn moves selection bar down one page.  Selection  *
'* status of list items is unchanged.                                    *
'*************************************************************************

   CalculatePageUpDisplayRange

   LastSelectedItem = DisplayRange.FirstListItem
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus
   DisplayList

   End Sub

Private Sub ProcessPageUpKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, PgUp moves selection bar up one page.       *
'*************************************************************************

   CalculatePageUpDisplayRange

   LastSelectedItem = DisplayRange.FirstListItem
   ProcessSelectedItem
   DisplayList

End Sub

Private Sub ProcessPageUpKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, PgUp moves focus rectangle up one page.   *
'*************************************************************************

   CalculatePageUpDisplayRange

   ItemWithFocus = DisplayRange.FirstListItem
   m_ListIndex = ItemWithFocus
   DisplayList

End Sub

Private Sub ProcessPageUpKey_MultiSelectExtended()

'*************************************************************************
'* in MultiSelect Extended mode, PgUp acts just like MultiSelect None    *
'* mode if the shift key is not pressed.  If shift is down, all items    *
'* from the item that had the focus when shift was pressed to the new    *
'* item that has the focus are selected.                                 *
'*************************************************************************

   If Not ShiftKeyDown Then
'     reinitialize the selected array to all False.
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

   CalculatePageUpDisplayRange

   LastSelectedItem = DisplayRange.FirstListItem
   ProcessSelected_MultiSelectExtended False ' don't calculate m_SelCount

'  if shift is down, all items from the item that had the focus when
'  shift was pressed to the new item that has the focus are selected.
   If ShiftKeyDown Then
      SetSelectedArrayRange ShiftDownStartItem, ItemWithFocus, True
      CalculateSelCount
   Else
      m_SelCount = 1
   End If

   DisplayList

End Sub

Private Sub CalculatePageUpDisplayRange()

'*************************************************************************
'* determines the first and last list items to display on PgUp keypress. *
'*************************************************************************

   If ItemWithFocus - MaxDisplayItems + 1 >= 0 Then
      DisplayRange.LastListItem = ItemWithFocus
      DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
   Else
      DisplayRange.FirstListItem = 0
      If m_ListCount >= MaxDisplayItems Then
         DisplayRange.LastListItem = MaxDisplayItems - 1
      Else
         DisplayRange.LastListItem = m_ListCount - 1
      End If
   End If

End Sub

Private Sub ProcessPageDownKey()

'*************************************************************************
'* processes page down key for all listbox states.                       *
'* page down ALWAYS starts from the item with the focus rectangle, even  *
'* if that item is not currently in the display range.  After the page   *
'* down the item that formerly had focus rect will be first displayed    *
'* entry (unless said entry was below first item on last page).          *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Or m_ListCount = 0 Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessPageDownKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessPageDownKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessPageDownKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessPageDownKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessPageDownKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, PgDn moves selection bar down one page.  Selection  *
'* status of list items is unchanged.                                    *
'*************************************************************************

   CalculatePageDownDisplayRange

   LastSelectedItem = DisplayRange.LastListItem
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus
   DisplayList

End Sub

Private Sub ProcessPageDownKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, PgDn moves selection bar down one page.     *
'*************************************************************************

   CalculatePageDownDisplayRange

   LastSelectedItem = DisplayRange.LastListItem
   ProcessSelectedItem
   DisplayList

End Sub

Private Sub ProcessPageDownKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, PgDn moves focus rectangle down one page. *
'*************************************************************************

   CalculatePageDownDisplayRange

   ItemWithFocus = DisplayRange.LastListItem
   m_ListIndex = ItemWithFocus
   DisplayList

End Sub

Private Sub ProcessPageDownKey_MultiSelectExtended()

'*************************************************************************
'* in MultiSelect Extended mode, PgDn acts just like MultiSelect None    *
'* mode if the shift key is not pressed.  If shift is down, all items    *
'* from the item that had the focus when shift was pressed to the new    *
'* item that has the focus are selected.                                 *
'*************************************************************************

   If Not ShiftKeyDown Then
'     reinitialize the selected array to all False.
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

   CalculatePageDownDisplayRange

   LastSelectedItem = DisplayRange.LastListItem
   ProcessSelected_MultiSelectExtended False ' don't calculate m_SelCount

'  if shift is down, all items from the item that had the focus when
'  shift was pressed to the new item that has the focus are selected.
   If ShiftKeyDown Then
      SetSelectedArrayRange ShiftDownStartItem, ItemWithFocus, True
      CalculateSelCount
   Else
      m_SelCount = 1
   End If

   DisplayList

End Sub

Private Sub CalculatePageDownDisplayRange()

'*************************************************************************
'* determines the first and last list items to display on PgDn keypress. *
'*************************************************************************

   If ItemWithFocus + MaxDisplayItems - 1 < m_ListCount Then
      DisplayRange.FirstListItem = ItemWithFocus
      DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
   Else
      DisplayRange.LastListItem = m_ListCount - 1
      If m_ListCount >= MaxDisplayItems Then
         DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
      Else
         DisplayRange.FirstListItem = 0
      End If
   End If

End Sub

Private Sub ProcessEndKey()

'*************************************************************************
'* processes end key for all listbox states.                             *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessEndKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessEndKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessEndKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessEndKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessEndKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, selection bar is moved to bottom of list.  The sel- *
'* ected status of listbox items is unchanged.                           *
'*************************************************************************

   If LastSelectedItem <> m_ListCount - 1 Then
      LastSelectedItem = m_ListCount - 1
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      DetermineLastPageDisplayRange
      DisplayList
   End If

End Sub

Private Sub ProcessEndKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, selection bar is moved to end of list.      *
'*************************************************************************

   If LastSelectedItem <> m_ListCount - 1 Then
      ClearPreviousSelection KeepSelectionNo
      DetermineLastPageDisplayRange
      LastSelectedItem = m_ListCount - 1
      ProcessSelectedItem
      DisplayList
   End If

End Sub

Private Sub ProcessEndKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, focus rectangle is moved to end of list.  *
'*************************************************************************

   If ItemWithFocus <> m_ListCount - 1 Then
'     repaint previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
      DetermineLastPageDisplayRange
      LastSelectedItem = m_ListCount - 1
      ItemWithFocus = m_ListCount - 1
      m_ListIndex = ItemWithFocus
      DisplayList
   End If

End Sub

Private Sub ProcessEndKey_MultiSelectExtended()

'*************************************************************************
'* processes End/Shift-End key in MultiSelect Extended mode.           *
'*************************************************************************

'  make sure any previously selected items are de-selected if shift key is not being pressed.
   If Not ShiftKeyDown Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
      m_SelCount = 1
   Else
'     if the shift key is down, all items from ItemWithFocus to the bottom of the list are selected.
      SetSelectedArrayRange ItemWithFocus, m_ListCount - 1, True
'     for special cases like Extended mode Shift-Home/Shift-End we need to brute-force it.
      CalculateSelCount
   End If

   If LastSelectedItem <> m_ListCount - 1 Then
      ClearPreviousSelection KeepSelectionNo
   End If

   DetermineLastPageDisplayRange
   LastSelectedItem = m_ListCount - 1
   ProcessSelected_MultiSelectExtended False
   DisplayList

End Sub

Private Sub DetermineLastPageDisplayRange()

'*************************************************************************
'* calculates range of items to display at bottom of list.               *
'*************************************************************************

   DisplayRange.LastListItem = m_ListCount - 1
   If m_ListCount >= MaxDisplayItems Then
      DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
   Else
      DisplayRange.FirstListItem = 0
   End If

End Sub

Private Sub ProcessHomeKey()

'*************************************************************************
'* processes home key for all listbox states.                            *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessHomeKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessHomeKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessHomeKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessHomeKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessHomeKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, selection bar is moved to top of list.  The selec-  *
'* ted status of listbox items is unchanged.                             *
'*************************************************************************

   If LastSelectedItem <> 0 Then
      LastSelectedItem = 0
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      DetermineFirstPageDisplayRange
      DisplayList
   End If

End Sub

Private Sub ProcessHomeKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, selection bar is moved to top of list.      *
'*************************************************************************

   If LastSelectedItem <> 0 Then
      ClearPreviousSelection KeepSelectionNo
      DetermineFirstPageDisplayRange
      LastSelectedItem = 0
      ProcessSelectedItem
      DisplayList
   End If

End Sub

Private Sub ProcessHomeKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, focus rectangle is moved to top of list.  *
'*************************************************************************

   If ItemWithFocus <> 0 Then
'     repaint previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
      DetermineFirstPageDisplayRange
      ItemWithFocus = 0
      m_ListIndex = ItemWithFocus
      DisplayList
   End If

End Sub

Private Sub ProcessHomeKey_MultiSelectExtended()

'*************************************************************************
'* processes Home/Shift-Home key in MultiSelect Extended mode.           *
'*************************************************************************

'  make sure any previously selected items are de-selected if shift key is not being pressed.
   If Not ShiftKeyDown Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
      m_SelCount = 1
   Else
'     if the shift key is down, all items from ItemWithFocus to the top of the list are selected.
      SetSelectedArrayRange 0, ItemWithFocus, True
'     for special cases like Extended mode Shift-Home/Shift-End we need to brute-force it.
      CalculateSelCount
   End If

   If LastSelectedItem <> 0 Then
      ClearPreviousSelection KeepSelectionNo
   End If

   LastSelectedItem = 0
   ProcessSelected_MultiSelectExtended False

   DetermineFirstPageDisplayRange
   DisplayList

End Sub

Private Sub DetermineFirstPageDisplayRange()

'*************************************************************************
'* calculates range of items to display at top of list.                  *
'*************************************************************************

   DisplayRange.FirstListItem = 0
   If m_ListCount < MaxDisplayItems Then
      DisplayRange.LastListItem = m_ListCount - 1
   Else
      DisplayRange.LastListItem = MaxDisplayItems - 1
   End If

End Sub

Private Sub ProcessSpaceBar()

'*************************************************************************
'* processes list item selection via spacebar for all listbox states.    *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessSpaceBar_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessSpaceBar_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessSpaceBar_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessSpaceBar_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessSpaceBar_CheckBoxMode()

'*************************************************************************
'* toggles selection status of focused item in CheckBox mode.            *
'*************************************************************************

   SelectedArray(LastSelectedItem) = Not SelectedArray(LastSelectedItem)
   If SelectedArray(LastSelectedItem) Then
      m_SelCount = m_SelCount + 1
   Else
      m_SelCount = m_SelCount - 1
   End If

'  redisplay the item to reflect current selection status.
   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessSpaceBar_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, spacebar selects item, with no toggle.      *
'*************************************************************************

   LastSelectedItem = ItemWithFocus
   ProcessSelectedItem

End Sub

Private Sub ProcessSpaceBar_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, the space bar toggles selection of the    *
'* item with focus rectangle.                                            *
'*************************************************************************

   If SelectedArray(ItemWithFocus) Then
      SelectedArray(ItemWithFocus) = False
      m_SelCount = m_SelCount - 1
      m_ListIndex = ItemWithFocus
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh
   Else
      LastSelectedItem = ItemWithFocus
      ProcessSelectedItem
   End If

End Sub

Private Sub ProcessSpaceBar_MultiSelectExtended()

'*************************************************************************
'* in MultiSelect Extended mode, space bar selects item contained in the *
'* focus rectangle, deselecting all other selected items if the Shift    *
'* key is not being pressed.                                             *
'*************************************************************************

   Dim NumPreviouslySelected As Long

   NumPreviouslySelected = m_SelCount

'  make sure any previously selected items are de-selected if shift key is not being pressed.
   If Not ShiftKeyDown Then
      SetSelectedArrayRange 0, m_ListCount - 1, False
      m_SelCount = 0
   End If

   LastSelectedItem = ItemWithFocus
   ProcessSelectedItem
   AdjustDisplayRange

   If NumPreviouslySelected > 1 Then
      DisplayList ' instead of just the list item; other items may have to be visually deselected.
   Else
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessUpArrowKey()

'*************************************************************************
'* processes up arrow key for all listbox states.                        *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessUpArrowKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessUpArrowKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessUpArrowKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessUpArrowKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessUpArrowKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, up and down arrow keys only move selection gradient *
'* bar.  The selection status of the item is unchanged.                  *
'*************************************************************************

   If LastSelectedItem > 0 Then
      If LastSelectedItem <> DisplayRange.FirstListItem Then ' prevents flicker
         DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      End If
      LastSelectedItem = LastSelectedItem - 1
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      AdjustDisplayRange
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessUpArrowKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, up arrow key selects the item with the      *
'* focus rectangle if it's not already selected.  Otherwise, it moves    *
'* the selection bar up one list item.                                   *
'*************************************************************************

   If m_SelCount = 0 Then
      LastSelectedItem = ItemWithFocus
      ProcessSelectedItem
      AdjustDisplayRange
   Else
      If LastSelectedItem > 0 Then
'        repaint previously selected list item without selection gradient / focus rectangle.
'        to prevent flicker, don't repaint if first item in display is focused.
         If ItemWithFocus <> DisplayRange.FirstListItem Then
            ClearPreviousSelection KeepSelectionNo
         End If
         LastSelectedItem = LastSelectedItem - 1
         ProcessSelectedItem
         AdjustDisplayRange
      End If
   End If

End Sub

Private Sub ProcessUpArrowKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, the arrow keys move just the focus rect-  *
'* angle.  Selection status of each affected list item is unchanged.     *
'*************************************************************************

   If ItemWithFocus > 0 Then

'     'defocus' previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo

      ItemWithFocus = ItemWithFocus - 1
      m_ListIndex = ItemWithFocus

      AdjustDisplayRange
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh

   End If

End Sub

Private Sub ProcessUpArrowKey_MultiSelectExtended()

'*************************************************************************
'* controls up arrow processing in MultiSelect Extended mode.            *
'*************************************************************************

   Dim NumPreviouslySelected As Long

   NumPreviouslySelected = m_SelCount

   If Not ShiftKeyDown Then
'     make sure any previously selected items are de-selected if shift key is not being pressed.
      SetSelectedArrayRange 0, m_ListCount - 1, False
   End If

'  repaint previously focused list item without focus rectangle, maintaining selection status.
'  to prevent flicker, don't repaint if item with focus is first in displayed range.
   If ItemWithFocus <> DisplayRange.FirstListItem Then
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
   End If

   If LastSelectedItem > 0 Then
      LastSelectedItem = LastSelectedItem - 1
   End If

   ProcessSelected_MultiSelectExtended False

   If ShiftKeyDown Then
      CalculateSelCount
   Else
      m_SelCount = 1
   End If

   AdjustDisplayRange
   If NumPreviouslySelected > 1 Then
      DisplayList ' instead of just the list item; other items may have to be visually deselected
   Else
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessDownArrowKey()

'*************************************************************************
'* processes down arrow key for all listbox states.                      *
'*************************************************************************

'  if the mouse button is down, exit.
   If MouseAction <> MOUSE_NOACTION Then
      Exit Sub
   End If

   If m_Style = [CheckBox] Then
      ProcessDownArrowKey_CheckBoxMode
   Else
      Select Case m_MultiSelect
         Case vbMultiSelectNone
            ProcessDownArrowKey_MultiSelectNone
         Case vbMultiSelectSimple
            ProcessDownArrowKey_MultiSelectSimple
         Case vbMultiSelectExtended
            ProcessDownArrowKey_MultiSelectExtended
      End Select
   End If

End Sub

Private Sub ProcessDownArrowKey_CheckBoxMode()

'*************************************************************************
'* in CheckBox mode, up and down arrow keys only move selection gradient *
'* bar.  The selection status of the item is unchanged.                  *
'*************************************************************************

   If LastSelectedItem < m_ListCount - 1 Then
      If LastSelectedItem <> DisplayRange.LastListItem Then ' prevents flicker
         DisplayListBoxItem LastSelectedItem, DrawAsUnselected, FocusRectangleNo
      End If
      LastSelectedItem = LastSelectedItem + 1
      ItemWithFocus = LastSelectedItem
      m_ListIndex = ItemWithFocus
      AdjustDisplayRange
      DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
      UserControl.Refresh
   End If

End Sub

Private Sub ProcessDownArrowKey_MultiSelectNone()

'*************************************************************************
'* in MultiSelect None mode, down arrow key selects the item with the    *
'* focus rectangle if it's not already selected.  Otherwise, it moves    *
'* the selection bar down one list item.                                 *
'*************************************************************************

   If m_SelCount = 0 Then
      LastSelectedItem = ItemWithFocus
      ProcessSelectedItem
      AdjustDisplayRange
   Else
      If LastSelectedItem < m_ListCount - 1 Then
'        repaint previously selected list item without selection gradient / focus rectangle.
'        to prevent flicker, don't repaint if last item in display is focused.
         If ItemWithFocus <> DisplayRange.LastListItem Then
            ClearPreviousSelection KeepSelectionNo
         End If
         LastSelectedItem = LastSelectedItem + 1
         ProcessSelectedItem
         AdjustDisplayRange
      End If
   End If

End Sub

Private Sub ProcessDownArrowKey_MultiSelectSimple()

'*************************************************************************
'* in MultiSelect Simple mode, the arrow keys move just the focus        *
'* rectangle.  Selection status of each affected list item is unchanged. *
'*************************************************************************

   If ItemWithFocus < m_ListCount - 1 Then

'     repaint previously focused list item, maintaining selection status.
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo

      ItemWithFocus = ItemWithFocus + 1
      m_ListIndex = ItemWithFocus

      AdjustDisplayRange
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
      UserControl.Refresh

   End If

End Sub

Private Sub ProcessDownArrowKey_MultiSelectExtended()

'*************************************************************************
'* controls down arrow processing in MultiSelect Extended mode.          *
'*************************************************************************

   Dim NumPreviouslySelected As Long

   NumPreviouslySelected = m_SelCount

   If LastSelectedItem < m_ListCount - 1 Then

'     make sure any previously selected items are de-selected if shift key is not being pressed.
      If Not ShiftKeyDown Then
         SetSelectedArrayRange 0, m_ListCount - 1, False
      End If

'     repaint previously selected list item without selection gradient / focus rectangle.
'     to prevent flicker, don't repaint if item with focus is last in displayed range.
      If ItemWithFocus <> DisplayRange.LastListItem Then
         DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleNo
      End If

      LastSelectedItem = LastSelectedItem + 1
      ProcessSelected_MultiSelectExtended False
      
      If ShiftKeyDown Then
'        other items may selected throughout the list; do it the hard way.
         CalculateSelCount
      Else
         m_SelCount = 1
      End If

      AdjustDisplayRange
      If NumPreviouslySelected > 1 Then
         DisplayList
      Else
         DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
         UserControl.Refresh
      End If

   Else

      If Not ShiftKeyDown Then
         SetSelectedArrayRange 0, m_ListCount - 1, False
         m_SelCount = 1
         ProcessSelected_MultiSelectExtended False
         AdjustDisplayRange
         DisplayList
      End If

   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'******************* bitmap tiling routines by Carles P.V.
' adapted from Carles' class titled "DIB Brush - Easy Image Tiling Using FillRect"
' at Planet Source Code, txtCodeId=40585.

Private Function SetPattern(Picture As StdPicture) As Boolean

'*************************************************************************
'* creates the brush pattern for tiling into the listbox.  By Carles P.V.*
'*************************************************************************

   Dim tBI       As BITMAP
   Dim tBIH      As BITMAPINFOHEADER
   Dim Buff()    As Byte 'Packed DIB

   Dim lhDC      As Long
   Dim lhOldBmp  As Long

   If (GetObjectType(Picture) = OBJ_BITMAP) Then

'     -- Get image info
      GetObject Picture, Len(tBI), tBI

'     -- Prepare DIB header and redim. Buff() array
      With tBIH
         .biSize = Len(tBIH) '40
         .biPlanes = 1
         .biBitCount = 24
         .biWidth = tBI.bmWidth
         .biHeight = tBI.bmHeight
         .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
      End With
      ReDim Buff(1 To Len(tBIH) + tBIH.biSizeImage) '[Header + Bits]

'     -- Create DIB brush
      lhDC = CreateCompatibleDC(0)
      If (lhDC <> 0) Then
         lhOldBmp = SelectObject(lhDC, Picture)

'        -- Build packed DIB:
'        - Merge Header
         CopyMemory Buff(1), tBIH, Len(tBIH)
'        - Get and merge DIB Bits
         GetDIBits lhDC, Picture, 0, tBI.bmHeight, Buff(Len(tBIH) + 1), tBIH, DIB_RGB_COLORS

         SelectObject lhDC, lhOldBmp
         DeleteDC lhDC

'        -- Create brush from packed DIB
         DestroyPattern
         m_hBrush = CreateDIBPatternBrushPt(Buff(1), DIB_RGB_COLORS)
      End If

   End If

   SetPattern = (m_hBrush <> 0)

End Function

Private Sub Tile(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)

'*************************************************************************
'* performs the tiling of the bitmap on the control.  By Carles P.V.     *
'*************************************************************************

   Dim TileRect As RECT
   Dim PtOrg    As POINTAPI

   If (m_hBrush <> 0) Then
      SetRect TileRect, X1, Y1, X2, Y2
      SetBrushOrgEx hdc, X1, Y1, PtOrg
'     -- Tile image
      FillRect hdc, TileRect, m_hBrush
   End If

End Sub

Private Sub DestroyPattern()
   
'*************************************************************************
'* destroys the pattern brush used to tile the bitmap.  By Carles P.V.   *
'*************************************************************************
   
   If (m_hBrush <> 0) Then
      DeleteObject m_hBrush
      m_hBrush = 0
   End If

End Sub

'******************* end of bitmap tiling routines by Carles P.V.

Private Sub CalculatePicXCoordinate()

'*************************************************************************
'* calculates the leftmost x coordinate for displaying listitem icons.   *
'*************************************************************************

   If Not m_RightToLeft Then
      If m_Style = [Standard] Then
         PicX = m_BorderWidth + 2
      Else
         PicX = m_BorderWidth + 20
      End If
   Else
      If m_Style = [Standard] Then
         If m_ItemImageSize = 0 Then
            PicX = ScaleWidth - m_BorderWidth - ListItemHeight - 1
         Else
            PicX = ScaleWidth - m_BorderWidth - m_ItemImageSize - 1
         End If
      Else
         If m_ItemImageSize = 0 Then
            PicX = ScaleWidth - m_BorderWidth - ListItemHeight - 19
         Else
            PicX = ScaleWidth - m_BorderWidth - m_ItemImageSize - 19
         End If
      End If
   End If

End Sub

Private Sub InitListBoxDisplayCharacteristics()

'*************************************************************************
'* initializes gradients, listitem height, and display coordinates.      *
'*************************************************************************

   Dim i As Long

'  get the height range characters in the current font.
   ListItemHeight = TextHeight("^j")

'  determine x coordinate of listitem icons based on checkbox mode and .RightToLeft.
   CalculatePicXCoordinate

'  calculate selection bar offset from left or right side of control.
'  Account for listitem images and checkboxes possibly being displayed.
   SelBarOffset = 20 * -(m_Style = [CheckBox]) + 1
   If m_ShowItemImages Then
      If m_ItemImageSize = 0 Then
         SelBarOffset = SelBarOffset + ListItemHeight
      Else
         SelBarOffset = SelBarOffset + m_ItemImageSize
      End If
   End If

'  initialize text draw coordinates and boundaries.
   InitTextDisplayCharacteristics

'  create a virtual bitmap that will hold the background gradient or picture.  Portions of
'  this virtual bitmap are blitted to the control background to restore the background
'  gradient/picture when list items are changed.  Saves time over repainting whole control
'  when we're just doing things like adding a listbox item or changing selection gradient.
   CreateVirtualDC hdc, VirtualBackgroundDC, mMemoryBitmap, mOrginalBitmap, ScaleWidth, ScaleHeight

'  initialize segments and graphics for border.
   InitBorder

'  calculate the various gradients that may be used in the control.
   CalculateGradients

'  place either the picture or gradient background onto the virtual DC.
   If IsPictureThere(m_ActivePicture) Then
      DisplayPicture
      CreateBorder
'     transfer the picture (with border) to the virtual DC bitmap.
      i = BitBlt(VirtualBackgroundDC, 0, 0, ScaleWidth, ScaleHeight, hdc, 0, 0, vbSrcCopy)
   Else
'     paint the gradient onto the virtual DC bitmap.
      Call StretchDIBits(VirtualBackgroundDC, _
                         0, 0, _
                         ScaleWidth, _
                         ScaleHeight, _
                         0, 0, _
                         ScaleWidth, _
                         ScaleHeight, _
                         BGlBits(0), _
                         BGuBIH, _
                         DIB_RGB_COLORS, _
                         vbSrcCopy)
'     transfer the gradient in the virtual bitmap to the usercontrol.
      i = BitBlt(hdc, 0, 0, ScaleWidth, ScaleHeight, VirtualBackgroundDC, 0, 0, vbSrcCopy)
      CreateBorder
   End If

End Sub

Private Sub InitBorder()

'*************************************************************************
'* create all segments and graphics for the control's border.            *
'*************************************************************************

'  create the horizontal border segment virtual DC.
   CreateVirtualDC hdc, VirtualDC_SegH, _
                   mMemoryBitmap_SegH, mOriginalBitmap_SegH, _
                   ScaleWidth + 1, m_BorderWidth

'  create the vertical border segment virtual DC.
   CreateVirtualDC hdc, VirtualDC_SegV, _
                   mMemoryBitmap_SegV, mOriginalBitmap_SegV, _
                   m_BorderWidth, ScaleHeight

'  calculate the primary horizontal segment gradient.
   CalculateGradient ScaleWidth, BorderWidth + 1, TranslateColor(m_ActiveBorderColor1), TranslateColor(m_ActiveBorderColor2), _
                     90, m_BorderMiddleOut, SegH1uBIH, SegH1lBits()

'  if gradients are not middle-out, calculate the secondary horizontal segment gradient.
   If Not m_BorderMiddleOut Then
      CalculateGradient ScaleWidth, BorderWidth + 1, TranslateColor(m_ActiveBorderColor2), TranslateColor(m_ActiveBorderColor1), _
                        90, m_BorderMiddleOut, SegH2uBIH, SegH2lBits()
   End If

'  calculate the primary vertical segment gradient.
   CalculateGradient BorderWidth + 1, ScaleHeight, TranslateColor(m_ActiveBorderColor1), TranslateColor(m_ActiveBorderColor2), _
                     180, m_BorderMiddleOut, SegV1uBIH, SegV1lBits()

'  if gradients are not middle-out, calculate the secondary vertical segment gradient.
   If Not m_BorderMiddleOut Then
      CalculateGradient BorderWidth + 1, ScaleHeight, TranslateColor(m_ActiveBorderColor2), TranslateColor(m_ActiveBorderColor1), _
                        180, m_BorderMiddleOut, SegV2uBIH, SegV2lBits()
   End If

'  create the four border segments.
   CreateBorderSegments

End Sub

Private Sub CreateBorderSegments()

'*************************************************************************
'* creates the vertical and horizontal trapezoidal border segments.      *
'*************************************************************************

   DeleteBorderSegmentObjects    ' make sure the segments don't already exist.

   BorderSegment(TOP_SEGMENT) = CreateDiagRectRegion(ScaleWidth, m_BorderWidth, 1, 1)
   BorderSegment(BOTTOM_SEGMENT) = CreateDiagRectRegion(ScaleWidth, m_BorderWidth, -1, -1)
   BorderSegment(RIGHT_SEGMENT) = CreateDiagRectRegion(m_BorderWidth, ScaleHeight, -1, -1)
   BorderSegment(LEFT_SEGMENT) = CreateDiagRectRegion(m_BorderWidth, ScaleHeight, 1, 1)

End Sub

Private Sub DeleteBorderSegmentObjects()

'*************************************************************************
'* destroys the border segment objects if they exist, to save memory.    *
'*************************************************************************

   If BorderSegment(TOP_SEGMENT) Then
      DeleteObject BorderSegment(TOP_SEGMENT)
   End If

   If BorderSegment(RIGHT_SEGMENT) Then
      DeleteObject BorderSegment(RIGHT_SEGMENT)
   End If

   If BorderSegment(BOTTOM_SEGMENT) Then
      DeleteObject BorderSegment(BOTTOM_SEGMENT)
   End If

   If BorderSegment(LEFT_SEGMENT) Then
      DeleteObject BorderSegment(LEFT_SEGMENT)
   End If

End Sub

Private Function CreateDiagRectRegion(ByVal cx As Long, ByVal cy As Long, SideAStyle As Integer, SideBStyle As Integer) As Long

'**************************************************************************
'* Author: LaVolpe                                                        *
'* the cx & cy parameters are the respective width & height of the region *
'* the passed values may be modified which coder can use for other purp-  *
'* oses like drawing borders or calculating the client/clipping region.   *
'* SideAStyle is -1, 0 or 1 depending on horizontal/vertical shape,       *
'*            reflects the left or top side of the region                 *
'*            -1 draws left/top edge like /                               *
'*            0 draws left/top edge like  |                               *
'*            1 draws left/top edge like  \                               *
'* SideBStyle is -1, 0 or 1 depending on horizontal/vertical shape,       *
'*            reflects the right or bottom side of the region             *
'*            -1 draws right/bottom edge like \                           *
'*            0 draws right/bottom edge like  |                           *
'*            1 draws right/bottom edge like  /                           *
'**************************************************************************

   Dim tpts(0 To 4) As POINTAPI    ' holds polygonal region vertices.

   If cx > cy Then ' horizontal

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cx < cy * 2 Then cy = cx \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).X = cy - 1
         tpts(1).X = -1
      ElseIf SideAStyle > 0 Then
         tpts(1).X = cy
      End If
      tpts(1).Y = cy

      tpts(2).X = cx + Abs(SideBStyle < 0)
      If SideBStyle > 0 Then tpts(2).X = tpts(2).X - cy
      tpts(2).Y = cy

      tpts(3).X = cx + Abs(SideBStyle < 0)
      If SideBStyle < 0 Then tpts(3).X = tpts(3).X - cy

   Else

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cy < cx * 2 Then cx = cy \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).Y = cx - 1
         tpts(3).Y = -1
      ElseIf SideAStyle > 0 Then
         tpts(3).Y = cx - 1
         tpts(0).Y = -1
      End If

      tpts(1).Y = cy
      If SideBStyle < 0 Then tpts(1).Y = tpts(1).Y - cx
      tpts(2).X = cx

      tpts(2).Y = cy
      If SideBStyle > 0 Then tpts(2).Y = tpts(2).Y - cx
      tpts(3).X = cx

   End If

   tpts(4) = tpts(0)

   CreateDiagRectRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Sub InitTextDisplayCharacteristics()

'*************************************************************************
'* Calculate text display coordinates and boundaries.  I keep this a     *
'* separate routine for when properties (such as .BorderWidth) that      *
'* affect text display are changed programmatically.  I can then quickly *
'* change text boundaries.                                               *
'*************************************************************************

   Dim AvailableDisplayHeight As Long    ' height, in pixels of displayable listbox area.
   Dim i As Long                         ' loop variable.

'  determine the number of items that can be displayed given listbox height, list
'  item height in the current font, and display style (normal or checkbox).
'  Also account for listitem images possibly being displayed.
   AvailableDisplayHeight = ScaleHeight - (Y_CLEARANCE * 2) - (m_BorderWidth * 2)
   If m_Style = [Standard] Then
      If m_ShowItemImages Then
         If ListItemHeight >= m_ItemImageSize Then
            MaxDisplayItems = Int(AvailableDisplayHeight / ListItemHeight)
         Else
            MaxDisplayItems = Int(AvailableDisplayHeight / m_ItemImageSize)
         End If
      Else
         MaxDisplayItems = Int(AvailableDisplayHeight / ListItemHeight)
      End If
   Else
      If m_ShowItemImages Then
         If ListItemHeight >= m_ItemImageSize Then
            If ListItemHeight < MIN_FONT_HEIGHT Then
               MaxDisplayItems = Int(AvailableDisplayHeight / ((ListItemHeight + 2) + (MIN_FONT_HEIGHT - ListItemHeight)))
            Else
               MaxDisplayItems = Int(AvailableDisplayHeight / (ListItemHeight + 2))
            End If
         Else
            MaxDisplayItems = Int(AvailableDisplayHeight / (m_ItemImageSize + 2))
         End If
      Else
         If ListItemHeight < MIN_FONT_HEIGHT Then
            MaxDisplayItems = Int(AvailableDisplayHeight / ((ListItemHeight + 2) + (MIN_FONT_HEIGHT - ListItemHeight)))
         Else
            MaxDisplayItems = Int(AvailableDisplayHeight / (ListItemHeight + 2))
         End If
      End If

   End If

'  determine the number of pixels of clearance from the border to start drawing text,
'  given the display mode (checkbox or normal) and listitem image display mode.
   If m_Style = [Standard] Then
      TextClearance = m_BorderWidth + 3
   Else
      TextClearance = m_BorderWidth + 23
   End If
   If m_ShowItemImages Then
      If m_ItemImageSize = 0 Then
         TextClearance = TextClearance + ListItemHeight + 1
      Else
         TextClearance = TextClearance + m_ItemImageSize + 1
      End If
   End If

'  initialize the y coordinate array.  Make the spacing between
'  list items a little wider if Checkbox display mode is active.
   YCoords(0) = m_BorderWidth + Y_CLEARANCE
   For i = 1 To MaxDisplayItems - 1
      If m_Style = [Standard] Then
         If m_ShowItemImages Then
            If ListItemHeight >= m_ItemImageSize Then
               YCoords(i) = YCoords(i - 1) + ListItemHeight
            Else
               YCoords(i) = YCoords(i - 1) + m_ItemImageSize
            End If
         Else
            YCoords(i) = YCoords(i - 1) + ListItemHeight
         End If
      Else
'        helps keeps checkboxes from getting squished when displaying in very small fonts.
         If m_ShowItemImages Then
            If ListItemHeight >= m_ItemImageSize Then
               YCoords(i) = YCoords(i - 1) + ListItemHeight + 2
            Else
               YCoords(i) = YCoords(i - 1) + m_ItemImageSize + 2
            End If
         Else
            If ListItemHeight < MIN_FONT_HEIGHT Then
               YCoords(i) = YCoords(i - 1) + ListItemHeight + 2 + (MIN_FONT_HEIGHT - ListItemHeight)
            Else
               YCoords(i) = YCoords(i - 1) + ListItemHeight + 2
            End If
         End If
      End If
   Next i

End Sub

Private Sub DrawRectangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)

'*************************************************************************
'* draws the checkbox, thumb border, and focus rectangles.               *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim hBrush As Long        ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long        ' the outer boundary of the border region.
   Dim hRgn2  As Long        ' the inner boundary of the border region.

'  create the outer region.
   hRgn1 = CreateRoundRectRgn(X1, Y1, X2, Y2, 0, 0)
'  create the inner region.
   hRgn2 = CreateRoundRectRgn(X1 + 1, Y1 + 1, X2 - 1, Y2 - 1, 0, 0)
   
'  combine the regions into one border region.
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create and apply the color brush.
   hBrush = CreateSolidBrush(TranslateColor(lColor))
   FillRgn hdc, hRgn2, hBrush

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub

Private Sub CalculateGradients()

'*************************************************************************
'* define all gradients (background, scroll track/button, selection bar).*
'* I split them up into different procedures so that when a particular   *
'* gradient property is changed, all gradients don't get regenerated.    *
'*************************************************************************

   CalculateBackGroundGradient
   CalculateHighlightBarGradient
   CalculateVerticalTrackbarGradients
   CalculateScrollbarButtonGradient
   CalculateScrollbarThumbGradient

End Sub

Private Sub CalculateBackGroundGradient()

'*************************************************************************
'* calculate the gradient for the background.  Even if a picture is used *
'* instead of a gradient, this allows control user to switch back and    *
'* forth between those two options in design or runtime modes.           *
'*************************************************************************

   CalculateGradient ScaleWidth, ScaleHeight, TranslateColor(m_ActiveBackColor1), TranslateColor(m_ActiveBackColor2), m_BackAngle, m_BackMiddleOut, BGuBIH, BGlBits(), m_CircularGradient

End Sub

Private Sub CalculateHighlightBarGradient()

'*************************************************************************
'*  calculate the gradient for the selected item highlight bar.          *
'*************************************************************************

   CalculateGradient ScaleWidth - (m_BorderWidth * 2), ListItemHeight, TranslateColor(m_ActiveSelColor1), TranslateColor(m_ActiveSelColor2), 90, True, SeluBIH, SellBits()

End Sub

Private Sub CalculateScrollbarThumbGradient()

'*************************************************************************
'* calculate the vertical scrollbar thumb gradient.                      *
'*************************************************************************

'  sized to scrollbar height at first, sized on the fly by StretchDIBits when drawing thumb.
   CalculateGradient ScrollBarButtonWidth, vScrollTrackHeight, TranslateColor(m_ActiveThumbColor1), TranslateColor(m_ActiveThumbColor2), 180, True, vThumbuBIH, vThumblBits()

End Sub

Private Sub CalculateVerticalTrackbarGradients()

'*************************************************************************
'* master routine for generating clicked/unclicked trackbar gradients.   *
'*************************************************************************

'  determine the height of the scrollbar track (area between up and down buttons).
   vScrollTrackHeight = CalculateScrollTrackHeight

   CalculateVerticalTrackbarGradientUnclicked
   CalculateVerticalTrackbarGradientClicked

End Sub

Private Sub CalculateVerticalTrackbarGradientUnclicked()

'*************************************************************************
'* calculate the gradient for the vertical scrollbar trackbar when the   *
'* mouse is not down on the trackbar.                                    *
'*************************************************************************

   CalculateGradient ScrollBarButtonWidth, vScrollTrackHeight, TranslateColor(m_ActiveTrackBarColor1), TranslateColor(m_ActiveTrackBarColor2), 180, True, VTrackuBIH, VTracklBits()

End Sub

Private Sub CalculateVerticalTrackbarGradientClicked()

'*************************************************************************
'* calculate the gradient for the vertical scrollbar trackbar when the   *
'* mouse is down on the trackbar.                                        *
'*************************************************************************

   CalculateGradient ScrollBarButtonWidth, vScrollTrackHeight, TranslateColor(m_TrackClickColor1), TranslateColor(m_TrackClickColor2), 180, True, vClickTrackuBIH, vClickTracklBits()

End Sub

Private Sub CalculateScrollbarButtonGradient()

'*************************************************************************
'*  calculate the gradient for the vertical scrollbar buttons.           *
'*************************************************************************

   CalculateGradient ScrollBarButtonWidth, ScrollBarButtonHeight, TranslateColor(m_ActiveButtonColor1), TranslateColor(m_ActiveButtonColor2), 180, True, TrackButtonuBIH, TrackButtonlBits()

End Sub

Private Sub RedrawControl()

'*************************************************************************
'* master routine for painting of MorphListBox control.                  *
'*************************************************************************

'  if the .RedrawFlag property is false, then don't redraw.  This property is
'  set to False by the programmer before large operations on the listbox (for
'  example, adding or removing a thousand items) and set back to True after the
'  operations are complete.  This saves unnecessary and time-consuming redraws.
   If m_RedrawFlag Then
      SetBackGround
      CreateBorder
      DisplayList
   End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub SetBackGround()

'*************************************************************************
'* displays control's background gradient or picture in initial draw.    *
'*************************************************************************

   If IsPictureThere(m_ActivePicture) Then
'     if the .Picture property has been defined, it takes precedence over gradient.
      DisplayPicture
   Else
'     paint the gradient onto the actual usercontrol DC.  Most subsequent repaints are handled
'     by blitting the appropriate gradient portions from the virtual bitmap's DC to the usercontrol.
'     Thanks to RedBird77 for tweaking this to work correctly with wide borders!
      Call StretchDIBits(hdc, m_BorderWidth, m_BorderWidth, _
                         ScaleWidth - (m_BorderWidth * 2), _
                         ScaleHeight - (m_BorderWidth * 2), _
                         m_BorderWidth, m_BorderWidth, _
                         ScaleWidth - (m_BorderWidth * 2), _
                         ScaleHeight - (m_BorderWidth * 2), _
                         BGlBits(0), BGuBIH, DIB_RGB_COLORS, vbSrcCopy)
   End If

End Sub

Private Sub DisplayPicture()

'*************************************************************************
'* if the .Picture property is defined, paints the picture onto the      *
'* control.  If picture tiling is indicated, that is performed.          *
'*************************************************************************

   Select Case m_ActivePictureMode
      Case [Normal]
         Set UserControl.Picture = m_ActivePicture
      Case [Tiled]
         SetPattern m_ActivePicture
         Tile hdc, m_BorderWidth, m_BorderWidth, ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth
      Case [Stretch]
         StretchPicture
   End Select

End Sub

Private Sub StretchPicture()

'*************************************************************************
'* stretch bitmap to fit listbox background.  Thanks to LaVolpe for the  *
'* suggestion and AllAPI.net / VBCity.com for the learning to do it.     *
'*************************************************************************

   Dim TempBitmap As BITMAP       ' bitmap structure that temporarily holds picture.
   Dim CreateDC As Long           ' used in creating temporary bitmap structure virtual DC.
   Dim TempBitmapDC As Long       ' virtual DC of temporary bitmap structure.
   Dim TempBitmapOld As Long      ' used in destroying temporary bitmap structure virtual DC.
   Dim r As Long                  ' result long for StretchBlt call.

'  create a temporary bitmap and DC to place the picture in.
   GetObjectAPI m_ActivePicture.handle, Len(TempBitmap), TempBitmap
   CreateDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   TempBitmapDC = CreateCompatibleDC(CreateDC)
   TempBitmapOld = SelectObject(TempBitmapDC, m_ActivePicture.handle)

'  streeeeeeeetch it...
   r = StretchBlt(hdc, m_BorderWidth, m_BorderWidth, _
                  ScaleWidth - m_BorderWidth * 2, ScaleHeight - m_BorderWidth * 2, _
                  TempBitmapDC, _
                  0, 0, _
                  TempBitmap.bmWidth, _
                  TempBitmap.bmHeight, vbSrcCopy)

'  destroy temporary bitmap DC.
   SelectObject TempBitmapDC, TempBitmapOld
   DeleteDC TempBitmapDC
   DeleteDC CreateDC

End Sub

Private Function IsPictureThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture.  Thanks to Roger Gilchrist.        *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsPictureThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub CreateBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvatures.    *
'*************************************************************************

   Dim i       As Long   ' return variable for BitBlt.
   Dim hRgn1   As Long   ' the outer region of the border.
   Dim hRgn2   As Long   ' the inner region of the border.
   Dim hBrush  As Long   ' the solid-color brush used to paint the combined border regions.

'  if the borderwidth is greater than 1 pixel, use the gradient border.
   If m_BorderWidth > 1 Then
'     display each border segment.
      DisplaySegment hdc, ScaleWidth, ScaleHeight, m_BorderWidth, TOP_SEGMENT, 0, 0, m_BorderMiddleOut
      DisplaySegment hdc, ScaleWidth, ScaleHeight, m_BorderWidth, LEFT_SEGMENT, 0, 0, m_BorderMiddleOut
      DisplaySegment hdc, ScaleWidth, ScaleHeight, m_BorderWidth, RIGHT_SEGMENT, ScaleWidth - m_BorderWidth, 0, m_BorderMiddleOut
      DisplaySegment hdc, ScaleWidth, ScaleHeight, m_BorderWidth, BOTTOM_SEGMENT, -1, ScaleHeight - m_BorderWidth, m_BorderMiddleOut
      Exit Sub
   End If

'  if border width is 1 or 0, use the line border.  In this instance, border curvature can be used also.
'  create the outer region.
   hRgn1 = pvGetRoundedRgn(0, 0, _
                           ScaleWidth, _
                           ScaleHeight, _
                           m_CurveTopLeft, _
                           m_CurveTopRight, _
                           m_CurveBottomLeft, _
                           m_CurveBottomRight)
'  create the inner region.
   hRgn2 = pvGetRoundedRgn(m_BorderWidth, _
                           m_BorderWidth, _
                           ScaleWidth - m_BorderWidth, _
                           ScaleHeight - m_BorderWidth, _
                           m_CurveTopLeft, _
                           m_CurveTopRight, _
                           m_CurveBottomLeft, _
                           m_CurveBottomRight)

'  combine the outer and inner regions.
   CombineRgn hRgn2, hRgn1, hRgn2, RGN_DIFF

'  create the solid brush pattern used to color the combined regions.
   hBrush = CreateSolidBrush(TranslateColor(m_ActiveBorderColor1))

'  color the combined regions.
   FillRgn hdc, hRgn2, hBrush

'  set the container's visibility region.
   SetWindowRgn hwnd, hRgn1, True

'  delete created objects to restore memory.
   DeleteObject hBrush
   DeleteObject hRgn1
   DeleteObject hRgn2

'  if we are redrawing the control because of a change to the .Picture property,
'  this is the time to re-blit the new picture/border to the virtual DC. I do
'  it here because I blit the entire control surface, including border.
   If ChangingPicture Then
      i = BitBlt(VirtualBackgroundDC, 0, 0, ScaleWidth, ScaleHeight, hdc, 0, 0, vbSrcCopy)
      ChangingPicture = False
   End If

End Sub

Private Sub DisplaySegment(ByVal TargetDC As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long, _
                           ByVal BorderWidth As Long, ByVal SegmentNdx As Long, _
                           ByVal StartX As Long, ByVal StartY As Long, ByVal bMOut As Boolean)

'*************************************************************************
'* displays one border segment.  Border segment gradients are displayed  *
'* to virtual bitmaps on the fly so that correct gradient orientation    *
'* is maintained if the .MiddleOut property is set to False.             *
'*************************************************************************

'  position the border segment region in the correct location.
   MoveRegionToXY BorderSegment(SegmentNdx), StartX, StartY

   Select Case SegmentNdx

      Case LEFT_SEGMENT
         PaintVerticalGradient ScaleHeight, SegV1uBIH, SegV1lBits()
         BlitToRegion VirtualDC_SegV, hdc, m_BorderWidth, ScaleHeight, BorderSegment(SegmentNdx), StartX, StartY

      Case RIGHT_SEGMENT
         If m_BorderMiddleOut Then
            PaintVerticalGradient ScaleHeight, SegV1uBIH, SegV1lBits()
         Else
            PaintVerticalGradient ScaleHeight, SegV2uBIH, SegV2lBits()
         End If
         BlitToRegion VirtualDC_SegV, hdc, m_BorderWidth, ScaleHeight, BorderSegment(SegmentNdx), StartX, StartY

      Case TOP_SEGMENT
         PaintHorizontalGradient ScaleWidth, SegH1uBIH, SegH1lBits()
         BlitToRegion VirtualDC_SegH, hdc, ScaleWidth, m_BorderWidth, BorderSegment(SegmentNdx), StartX, StartY

      Case BOTTOM_SEGMENT
         If m_BorderMiddleOut Then
            PaintHorizontalGradient ScaleWidth, SegH1uBIH, SegH1lBits()
         Else
            PaintHorizontalGradient ScaleWidth, SegH2uBIH, SegH2lBits()
         End If
         BlitToRegion VirtualDC_SegH, hdc, ScaleWidth, m_BorderWidth, BorderSegment(SegmentNdx), StartX, StartY

   End Select

End Sub

Private Sub PaintHorizontalGradient(ByVal TargetWidth As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate horizontal gradient to horizontal virtual bitmap.  *
'*************************************************************************

   Call StretchDIBits(VirtualDC_SegH, _
                      0, 0, _
                      TargetWidth, m_BorderWidth, _
                      0, 1, _
                      TargetWidth, m_BorderWidth - 1, _
                      lBits(0), uBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

End Sub

Private Sub PaintVerticalGradient(ByVal TargetHeight As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate vertical gradient to vertical virtual bitmap.      *
'*************************************************************************

   Call StretchDIBits(VirtualDC_SegV, _
                      0, 0, _
                      m_BorderWidth, TargetHeight, _
                      1, 0, _
                      m_BorderWidth - 1, TargetHeight, _
                      lBits(0), uBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

End Sub

Private Sub MoveRegionToXY(ByVal Rgn As Long, ByVal X As Long, ByVal Y As Long)

'*************************************************************************
'* moves the supplied region to absolute X,Y coordinates.                *
'*************************************************************************

   Dim r As RECT    ' holds current X and Y coordinates of region.

'  get the current X,Y coordinates of the region.
   GetRgnBox Rgn, r

'  shift the region to 0,0 then to X,Y.
   OffsetRgn Rgn, -r.Left + X, -r.Top + Y

End Sub

Private Sub BlitToRegion(ByVal SourceDC As Long, DestDC As Long, lWidth As Long, lHeight As Long, Region As Long, ByVal XPos As Long, ByVal YPos As Long)

'*************************************************************************
'* blits the contents of a source DC to a non-rectangular region in a    *
'* destination DC.  A clipping region is selected in the destination DC, *
'* then the source DC is blitted to that location.  Technique is used in *
'* this control to blit to the trapezoid-shaped border regions.  Thanks  *
'* to LaVolpe for his help in tweaking this routine.                     *
'*************************************************************************

   Dim r              As Long    ' bitblt function call return.
   Dim ClippingRegion As Long    ' clipping region for bitblt.

'  move the region to the desired position.
   MoveRegionToXY Region, XPos, YPos

'  select a clipping region consisting of the segment parameter.
   ClippingRegion = SelectClipRgn(DestDC, Region)

'  blit the virtual bitmap to the control or form.  Since the clipping region has been
'  selected, only that region-shaped portion of the background will actually be drawn.
   r = BitBlt(DestDC, XPos, YPos, lWidth, lHeight, SourceDC, 0, 0, vbSrcCopy)

'  remove the clipping region constraint from the control.
   SelectClipRgn DestDC, ByVal 0&

'  reset the region coordinates to 0,0.
   MoveRegionToXY Region, 0, 0

End Sub

Private Function pvGetRoundedRgn(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
                                 ByVal TopLeftRadius As Long, ByVal TopRightRadius As Long, _
                                 ByVal BottomLeftRadius As Long, ByVal BottomRightRadius As Long) As Long

'*************************************************************************
'* allows each corner of the container to have its own curvature.        *
'* Code by the Amazing Carles P.V.  Thanks a million (as usual) Carles.  *
'*************************************************************************

   Dim hRgnMain As Long   ' the original "starting point" region.
   Dim hRgnTmp1 As Long   ' the first region that defines a corner's radius.
   Dim hRgnTmp2 As Long   ' the second region that defines a corner's radius.

'  bounding region.
   hRgnMain = CreateRectRgn(X1, Y1, X2, Y2)

'  top-left corner.
   hRgnTmp1 = CreateRectRgn(X1, Y1, X1 + TopLeftRadius, Y1 + TopLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(X1, Y1, X1 + 2 * TopLeftRadius, Y1 + 2 * TopLeftRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  top-right corner.
   hRgnTmp1 = CreateRectRgn(X2, Y1, X2 - TopRightRadius, Y1 + TopRightRadius)
   hRgnTmp2 = CreateEllipticRgn(X2 + 1, Y1, X2 + 1 - 2 * TopRightRadius, Y1 + 2 * TopRightRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  bottom-left corner.
   hRgnTmp1 = CreateRectRgn(X1, Y2, X1 + BottomLeftRadius, Y2 - BottomLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(X1, Y2 + 1, X1 + 2 * BottomLeftRadius, Y2 + 1 - 2 * BottomLeftRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  bottom-right corner.
   hRgnTmp1 = CreateRectRgn(X2, Y2, X2 - BottomRightRadius, Y2 - BottomRightRadius)
   hRgnTmp2 = CreateEllipticRgn(X2 + 1, Y2 + 1, X2 + 1 - 2 * BottomRightRadius, Y2 + 1 - 2 * BottomRightRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

   pvGetRoundedRgn = hRgnMain

End Function

Private Sub CombineRegions(ByVal Region1 As Long, ByVal Region2 As Long, ByVal MainRegion As Long)

'*************************************************************************
'* combines outer/inner rectangular regions for border painting.         *
'*************************************************************************

   Call CombineRgn(Region1, Region1, Region2, RGN_DIFF)
   Call CombineRgn(MainRegion, MainRegion, Region1, RGN_DIFF)
   Call DeleteObject(Region1)
   Call DeleteObject(Region2)

End Sub

Private Sub CalculateGradient(Width As Long, Height As Long, _
                             ByVal Color1 As Long, ByVal Color2 As Long, _
                             ByVal Angle As Single, ByVal bMOut As Boolean, _
                             ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long, _
                             Optional ByVal Circular As Boolean = False)

'*************************************************************************
'* Carles P.V.'s linear and circular gradient routines, modified by me   *
'* to allow both gradient types to be generated from one procedure.      *
'*************************************************************************

   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long

   Dim Offset1   As Long, Offset2 As Long
   Dim iPad      As Long, jPad    As Long

   Dim ia        As Long, iaa     As Long
   Dim ja        As Long, jaa     As Long

   Dim s()       As Long ' squares sequence
   Dim sc        As Long ' squares sequence counter (sequence index -> root)

   If (Width > 0 And Height > 0) Then

      If Circular = False Then

'        when angle is >= 91 and <= 270, the colors
'        invert in MiddleOut mode.  This corrects that.
         If bMOut And Angle >= 91 And Angle <= 270 Then
            g = Color1
            Color1 = Color2
            Color2 = g
         End If

'        -- Right-hand [+] (ox=0º)
         Angle = -Angle + 90

'        -- Normalize to [0º;360º]
         Angle = Angle Mod 360
         If (Angle < 0) Then
            Angle = 360 + Angle
         End If

'        -- Get quadrant (0 - 3)
         lQuad = Angle \ 90

'        -- Normalize to [0º;90º]
           Angle = Angle Mod 90

'        -- Calc. gradient length ('distance')
         If (lQuad Mod 2 = 0) Then
            AngleDiag = Atn(Width / Height) * TO_DEG
         Else
            AngleDiag = Atn(Height / Width) * TO_DEG
         End If
         AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
         Angle = Angle * TO_RAD
         g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'        -- Decompose colors
         If (lQuad > 1) Then
            lClr = Color1
            Color1 = Color2
            Color2 = lClr
         End If
         R1 = (Color1 And &HFF&)
         G1 = (Color1 And &HFF00&) \ 256
         b1 = (Color1 And &HFF0000) \ 65536
         R2 = (Color2 And &HFF&)
         G2 = (Color2 And &HFF00&) \ 256
         b2 = (Color2 And &HFF0000) \ 65536

'        -- Get color distances
         dR = R2 - R1
         dG = G2 - G1
         dB = b2 - b1

'        -- Size gradient-colors array
         ReDim lGrad(0 To g - 1)
         ReDim lGrad2(0 To g - 1)

'        -- Calculate gradient-colors
         iEnd = g - 1
         If (iEnd = 0) Then
'           -- Special case (1-pixel wide gradient)
            lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
         Else
            For i = 0 To iEnd
               lGrad2(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
            Next i
         End If

'        'if' block added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
         If bMOut Then
            k = 0
            For i = 0 To iEnd Step 2
               lGrad(k) = lGrad2(i)
               k = k + 1
            Next i
            For i = iEnd - 1 To 1 Step -2
               lGrad(k) = lGrad2(i)
               k = k + 1
            Next i
         Else
            For i = 0 To iEnd
               lGrad(i) = lGrad2(i)
            Next i
         End If

'        -- Size DIB array
         ReDim lBits(Width * Height - 1) As Long
         iEnd = Width - 1
         jEnd = Height - 1
         Scan = Width

'        -- Render gradient DIB
         Select Case lQuad

            Case 0, 2
               luSin = Sin(Angle) * INT_ROT
               luCos = Cos(Angle) * INT_ROT
               Offset = 0
               jIn = 0
               For j = 0 To jEnd
                  For i = 0 To iEnd
                     lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
                  Next i
                  jIn = jIn + luCos
                  Offset = Offset + Scan
               Next j

            Case 1, 3
               luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
               luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
               Offset = jEnd * Scan
               jIn = 0
               For j = 0 To jEnd
                  For i = 0 To iEnd
                     lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
                  Next i
                  jIn = jIn + luCos
                  Offset = Offset - Scan
               Next j

         End Select

'        -- Define DIB header
         With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height
         End With

      Else

         '-- Calc. gradient length ('diagonal')
         g = Sqr(Width * Width + Height * Height) \ 2

         '-- Decompose colors
         R1 = (Color1 And &HFF&)
         G1 = (Color1 And &HFF00&) \ 256
         b1 = (Color1 And &HFF0000) \ 65536
         R2 = (Color2 And &HFF&)
         G2 = (Color2 And &HFF00&) \ 256
         b2 = (Color2 And &HFF0000) \ 65536

         '-- Get color distances
         dR = R2 - R1
         dG = G2 - G1
         dB = b2 - b1

         '-- Size gradient-colors array
         ReDim lGrad(0 To g)

         '-- Build squares sequence LUT
         ReDim s(0 To g)
         For i = 1 To g
            s(i) = s(i - 1) + i + i - 1
         Next i

         '-- Calculate gradient-colors
         If (g = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
         Else
            For i = 0 To g
               lGrad(i) = b1 + (dB * i) \ g + 256 * (G1 + (dG * i) \ g) + 65536 * (R1 + (dR * i) \ g)
            Next i
         End If

         '-- Size DIB array
         ReDim lBits(Width * Height - 1) As Long

         '== Render gradient DIB

         '-- First "quadrant"...

         Scan = Width
         iPad = Width Mod 2
         jPad = Height Mod 2

         iEnd = Scan \ 2 + iPad - 1
         jEnd = Height \ 2 + jPad - 1
         Offset1 = jEnd * Scan + Scan \ 2

         ja = 1
         jaa = -1
         For j = 0 To jEnd
            sc = j
            ja = ja + jaa
            jaa = jaa + 2
            ia = ja + 1
            iaa = -1
            For i = Offset1 To Offset1 + iEnd
               ia = ia + iaa
               iaa = iaa + 2
               lBits(i) = lGrad(sc)
               If (ia >= s(sc) - sc) Then
                  sc = sc + 1
               End If
            Next i
            Offset1 = Offset1 - Scan
         Next j

         '-- Mirror first "quadrant"

         iEnd = iEnd - iPad
         Offset1 = 0
         Offset2 = Scan - 1

         For j = 0 To jEnd
            For i = 0 To iEnd
               lBits(Offset1 + i) = lBits(Offset2 - i)
            Next i
            Offset1 = Offset1 + Scan
            Offset2 = Offset2 + Scan
         Next j

         '-- Mirror first "half"

         iEnd = Scan - 1
         jEnd = jEnd - jPad
         Offset1 = (Height - 1) * Scan
         Offset2 = 0

         For j = 0 To jEnd
            For i = 0 To iEnd
               lBits(Offset1 + i) = lBits(Offset2 + i)
            Next i
            Offset1 = Offset1 - Scan
            Offset2 = Offset2 + Scan
         Next j

         '-- Define DIB header
         With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height
         End With

      End If

   End If

End Sub

Private Sub DisplayList(Optional vThumbYPos As Single = -1)

'*************************************************************************
'* controls the display of all visible list items.                       *
'*************************************************************************

   Dim i As Long     ' loop variable.

   VerticalScrollBarActive = (m_ListCount > MaxDisplayItems)

'  determine x coordinate of listitem icons based on checkbox mode and .RightToLeft.
'  I do this every time I display the list in case .RightToLeft changes in code.
   CalculatePicXCoordinate

'  Calculate scroll bar impact on listitem display.  Zero if scroll bar not active.
   lBarWid = ScrollBarButtonWidth * -VerticalScrollBarActive

'  repaint the picture background or gradient.
   SetBackGround

'  safety net.
   If DisplayRange.FirstListItem = -1 Then
      UserControl.Refresh
      Exit Sub
   End If

'  if the entire list will fit in the listbox...
   If m_ListCount <= MaxDisplayItems Then
      DisplayRange.FirstListItem = 0
      DisplayRange.LastListItem = m_ListCount - 1
   Else
'     if not displaying the very end of the list...
      If Not InDisplayedItemRange(m_ListCount - 1) Then
         DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
      Else
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems
      End If
   End If

'  set the .TopIndex property.
   m_TopIndex = DisplayRange.FirstListItem

'  display the appropriate listbox items, as selected or unselected.
   For i = DisplayRange.FirstListItem To DisplayRange.LastListItem
      DisplayListBoxItem i, SelectedArray(i), FocusRectangleNo
   Next i

'  if there's a list entry with a focus rectangle visible, redraw it.
   If m_Style = [Standard] Then
      DisplayListBoxItem ItemWithFocus, SelectedArray(ItemWithFocus), FocusRectangleYes
   Else
      If LastSelectedItem = -1 Then
         LastSelectedItem = 0
         ItemWithFocus = 0
         m_ListIndex = 0
      End If
      DisplayListBoxItem ItemWithFocus, DrawAsSelected, FocusRectangleYes
   End If

'  draw scrollbar if called for.
   If VerticalScrollBarActive Then
'      If vThumbYPos = -1 Then vThumbYPos = m_BorderWidth + VerticalThumbY ' mru
      DisplayVerticalScrollBar vThumbYPos
   End If

'  if there is a picture background instead of a gradient, or any control corners have
'  curvature, the border needs to be redrawn.  Thanks to Light Templer for catching this bug.
   If IsPictureThere(m_ActivePicture) Or (m_CurveTopLeft + m_CurveTopRight + m_CurveBottomLeft + m_CurveBottomRight > 0) Then
      CreateBorder
   End If

   UserControl.Refresh

End Sub

Private Sub DisplayListBoxItem(ByVal Index As Long, ByVal ItemSelected As Boolean, ByVal FocusRectFlag As Boolean)

'*************************************************************************
'* displays one (selected or unselected) listbox entry, using the spec-  *
'* ified ListArray index, in appropriate style (CheckBox or Standard).   *
'* Thanks to Redbird77 for optimizing the hell out of this routine!      *
'*************************************************************************

   Dim r           As RECT    ' the listitem text display rectangle.
   Dim lRet        As Long    ' bitblt function return.
   Dim nDisp       As Long    ' the index in the viewable list area of the listitem.
   Dim yStart      As Long
   Dim SelY        As Long

   If Not InDisplayedItemRange(Index) Or (m_ListFont Is Nothing) Then
      Exit Sub
   End If

   nDisp = GetDisplayIndexFromArrayIndex(Index)

   If nDisp < 0 Then
      nDisp = 0
   End If

   If m_ShowItemImages Then
      If ListItemHeight >= m_ItemImageSize Then
         SelY = YCoords(nDisp)
      Else
         SelY = YCoords(nDisp) + (m_ItemImageSize \ 4)
      End If
   Else
      SelY = YCoords(nDisp)
   End If

'  Draw selection bar background.  However, don't draw it if in checkbox mode
'  and item is selected but not the focused item (Index<> LastSelectedItem).  This is
'  because in CheckBox mode, only the item with focus has the selection bar background.
   If (m_Style = [Standard] And ItemSelected) Or (m_Style = [CheckBox] And ItemSelected And Index = LastSelectedItem) Then

'     if in checkbox mode, repaint gradient under checkbox so that checkbox is
'     "unchecked" when drawing checkbox (if the list item is now unselected).
'     Account for listitem images possibly being displayed.
      If m_Style = [CheckBox] Then
'        center the checkbox.
         If m_ShowItemImages Then
            If ListItemHeight >= m_ItemImageSize Then
               yStart = YCoords(nDisp) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
            Else
               yStart = YCoords(nDisp) + (m_ItemImageSize \ 4)
            End If
         Else
            If ListItemHeight > MIN_FONT_HEIGHT Then
               yStart = YCoords(nDisp) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
            Else
               yStart = YCoords(nDisp)
            End If
         End If
'        erase the checkbox checkmark.
         If Not m_RightToLeft Then
            lRet = BitBlt(hdc, m_BorderWidth, yStart, _
                          m_BorderWidth + 16, 13, _
                          VirtualBackgroundDC, _
                          m_BorderWidth, yStart, _
                          vbSrcCopy)
         Else
            lRet = BitBlt(hdc, ScaleWidth - m_BorderWidth - 16, yStart, _
                          16, 13, _
                          VirtualBackgroundDC, _
                          ScaleWidth - m_BorderWidth - 16, yStart, _
                          vbSrcCopy)
         End If
      End If

      UserControl.ForeColor = TranslateColor(m_ActiveSelTextColor)
'     draw the selection highlight gradient according to .RightToLeft property.
      If Not m_RightToLeft Then
         Call StretchDIBits(hdc, m_BorderWidth + SelBarOffset, SelY, _
                            ScaleWidth - (m_BorderWidth * 2) - lBarWid - SelBarOffset, ListItemHeight, _
                            0, 0, _
                            ScaleWidth - (m_BorderWidth * 2), ListItemHeight, _
                            SellBits(0), SeluBIH, _
                            DIB_RGB_COLORS, vbSrcCopy)
      Else
         Call StretchDIBits(hdc, m_BorderWidth + (ScrollBarButtonWidth * -(VerticalScrollBarActive = True)), SelY, _
                            ScaleWidth - (m_BorderWidth * 2) - lBarWid - SelBarOffset, ListItemHeight, _
                            0, 0, _
                            ScaleWidth - (m_BorderWidth * 2), ListItemHeight, _
                            SellBits(0), SeluBIH, _
                            DIB_RGB_COLORS, vbSrcCopy)
      End If

   Else

'     item is not highlighted; draw regular item background.
      UserControl.ForeColor = TranslateColor(m_ActiveTextColor)
      If Not m_RightToLeft Then
         lRet = BitBlt(hdc, m_BorderWidth, SelY, _
                       ScaleWidth - (m_BorderWidth * 2) - lBarWid, _
                       ListItemHeight + 1, _
                       VirtualBackgroundDC, _
                       m_BorderWidth, SelY - 1, vbSrcCopy)
      Else
         lRet = BitBlt(hdc, m_BorderWidth + (ScrollBarButtonWidth * -(VerticalScrollBarActive = True)), SelY, _
                       ScaleWidth - (m_BorderWidth * 2) - lBarWid, _
                       ListItemHeight + 1, _
                       VirtualBackgroundDC, _
                       m_BorderWidth + (ScrollBarButtonWidth * -(VerticalScrollBarActive = True)), SelY - 1, vbSrcCopy)
      End If

   End If

'  display listitem image if necessary.
   If m_ShowItemImages Then
      If ImageIndexArray(Index) <> -1 Then    ' if it's -1 there is no associated image.
         If m_ItemImageSize = 0 Then
'           if the ItemImageSize property is zero, that means we paint icon
'           in same width/height dimensions as listitem text height.
            UserControl.PaintPicture Images(ImageIndexArray(Index)), PicX, YCoords(nDisp), ListItemHeight - 1, ListItemHeight - 1
         Else
'           otherwise, set the width and height of the icon to ItemImageSize.
'           Determine the Y coordinate based on list item text height.
            If m_ItemImageSize < ListItemHeight Then
               UserControl.PaintPicture Images(ImageIndexArray(Index)), _
                                        PicX, YCoords(nDisp) + (ListItemHeight - m_ItemImageSize) \ 2, _
                                        m_ItemImageSize, m_ItemImageSize
            Else
               UserControl.PaintPicture Images(ImageIndexArray(Index)), _
                                        PicX, YCoords(nDisp), _
                                        m_ItemImageSize, m_ItemImageSize
            End If
         End If
      End If
   End If

'  calculate text rectangle size and position.
   Select Case m_RightToLeft

      Case False
         With r
            .Left = TextClearance
            If m_ShowItemImages Then
               If ListItemHeight >= m_ItemImageSize Then
                  .Top = YCoords(nDisp)
               Else
                  .Top = YCoords(nDisp) + m_ItemImageSize \ 4
               End If
            Else
               .Top = YCoords(nDisp)
            End If
            .Bottom = .Top + ListItemHeight
            .Right = ScaleWidth - m_BorderWidth - lBarWid
         End With

      Case True
         With r
            .Left = m_BorderWidth + lBarWid
            If m_Style = [Standard] Then
               If m_ShowItemImages Then
                  If ListItemHeight >= m_ItemImageSize Then
                     .Top = YCoords(nDisp)
                  Else
                     .Top = YCoords(nDisp) + m_ItemImageSize \ 4
                  End If
                  If m_ItemImageSize = 0 Then
                     .Right = ScaleWidth - m_BorderWidth - ListItemHeight - 4
                  Else
                     .Right = ScaleWidth - m_BorderWidth - m_ItemImageSize - 4
                  End If
               Else
                  .Right = ScaleWidth - m_BorderWidth - 3
                  .Top = YCoords(nDisp)
               End If
            Else
               If m_ShowItemImages Then
                  If ListItemHeight >= m_ItemImageSize Then
                     .Top = YCoords(nDisp)
                  Else
                     .Top = YCoords(nDisp) + m_ItemImageSize \ 4
                  End If
                  If m_ItemImageSize = 0 Then
                     .Right = ScaleWidth - m_BorderWidth - ListItemHeight - 23
                  Else
                     .Right = ScaleWidth - m_BorderWidth - m_ItemImageSize - 23
                  End If
               Else
                  .Right = ScaleWidth - 23
                  .Top = YCoords(nDisp)
               End If
            End If
            .Bottom = .Top + ListItemHeight
         End With

   End Select

'  display the text using DrawText api.
   If Not m_RightToLeft Then
      Call DrawText(UserControl.hdc, ListArray(Index), -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE)
   Else
      Call DrawText(UserControl.hdc, ListArray(Index), -1, r, DT_RIGHT Or DT_NOPREFIX Or DT_SINGLELINE)
   End If

'  display the checkbox, if the .Style property is set to CheckBox.
   If m_Style = [CheckBox] Then
      Call DisplayCheckBox(nDisp, SelectedArray(Index))
   End If

'  if the control and item both have focus, display the item's focus rectangle.
   If FocusRectFlag And HasFocus Then
      Call DisplayFocusRectangle(nDisp)
   End If

End Sub

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)

'*************************************************************************
'* draws the text with Unicode support based on OS version.              *
'* Thanks to Richard Mewett.                                             *
'*************************************************************************

   If mWindowsNT Then
      DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
   Else
      DrawTextA hdc, lpString, nCount, lpRect, wFormat
   End If

End Sub

Private Sub DisplayCheckBox(ByVal Index As Long, ByVal SelectedStatus As Boolean)

'*************************************************************************
'* draws an item-centered, one-pixel wide checkbox next to a list item.  *
'* If item is selected, draws a checkmark in one of three styles.        *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim YCoordStart As Long   ' starting y position of box based on font size.

'  center the checkbox, if the font is higher then the checkbox.
   If m_ShowItemImages Then
      If ListItemHeight >= m_ItemImageSize Then
         YCoordStart = YCoords(Index) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
      Else
         YCoordStart = YCoords(Index) + (m_ItemImageSize \ 4)
      End If
   Else
      If ListItemHeight > MIN_FONT_HEIGHT Then
         YCoordStart = YCoords(Index) + (ListItemHeight - MIN_FONT_HEIGHT) / 2
      Else
         YCoordStart = YCoords(Index)
      End If
   End If

'  display the checkbox.
   If Not m_RightToLeft Then
      DrawRectangle m_BorderWidth + 4, YCoordStart, m_BorderWidth + 18, YCoordStart + 14, m_CheckBoxColor
   Else
      DrawRectangle ScaleWidth - m_BorderWidth - 17, YCoordStart, ScaleWidth - m_BorderWidth - 3, YCoordStart + 14, m_CheckBoxColor
   End If

   If SelectedStatus Then
      Select Case m_CheckStyle
         Case [Arrow]
            DisplayCheckBoxArrow YCoordStart
         Case [Tick]
            DisplayCheckBoxTickMark YCoordStart
         Case [X]
            DisplayCheckBoxX YCoordStart
      End Select
   End If

ErrHandler:
   Exit Sub

End Sub

Private Sub DisplayCheckBoxArrow(ByVal Index As Long)

'*************************************************************************
'* draws an arrow in the checkbox of a selected list item when the       *
'* .Style property is set to Checkbox mode.                              *
'*************************************************************************

   Dim hPO           As Long    ' selected pen object.
   Dim hPN           As Long    ' pen object for drawing checkmark.
   Dim r             As Long    ' loop and result variable for api calls.
   Dim X1            As Long    ' the x coordinate of the start of the checkmark.
   Dim Y1            As Long    ' the y coordinate of the start of the checkmark vertical line.
   Dim Y2            As Long    ' the y coordinate of the end of the checkmark vertical line.
   Dim DrawDirection As Long    ' draw from left to right or right to left?

'  determine x and y coordinates of first part of check arrow to draw and the direction to draw.
   If Not m_RightToLeft Then
      X1 = m_BorderWidth + 9
      DrawDirection = 1
   Else
      X1 = ScaleWidth - m_BorderWidth - 9
      DrawDirection = -1
   End If
   Y1 = Index + 2
   Y2 = 10

'  draw the check arrow.
   hPN = CreatePen(0, 1, m_CheckBoxArrowColor)
   hPO = SelectObject(hdc, hPN)
   MoveTo hdc, X1, Y1, ByVal 0&
   For r = 1 To 6
      LineTo hdc, X1, Y1 + Y2
      X1 = X1 + DrawDirection
      Y1 = Y1 + 1
      Y2 = Y2 - 2
      MoveTo hdc, X1, Y1, ByVal 0&
   Next r

'  delete the pen object.
   r = SelectObject(hdc, hPO)
   r = DeleteObject(hPN)

End Sub

Private Sub DisplayCheckBoxX(ByVal Index As Long)

'*************************************************************************
'* draws an X in the checkbox of a selected list item when the .Style    *
'* property is set to Checkbox mode.                                     *
'*************************************************************************

   Dim i As Long
   Dim X As Long

   If Not m_RightToLeft Then
      X = m_BorderWidth + 7
   Else
      X = ScaleWidth - m_BorderWidth - 14
   End If

   For i = 1 To 2: SetPixelV hdc, X + i, Index + 4, m_CheckBoxArrowColor: Next i
   For i = 6 To 7: SetPixelV hdc, X + i, Index + 4, m_CheckBoxArrowColor: Next i
   For i = 1 To 3: SetPixelV hdc, X + i, Index + 5, m_CheckBoxArrowColor: Next i
   For i = 5 To 7: SetPixelV hdc, X + i, Index + 5, m_CheckBoxArrowColor: Next i
   For i = 2 To 6: SetPixelV hdc, X + i, Index + 6, m_CheckBoxArrowColor: Next i
   For i = 3 To 5: SetPixelV hdc, X + i, Index + 7, m_CheckBoxArrowColor: Next i
   For i = 2 To 6: SetPixelV hdc, X + i, Index + 8, m_CheckBoxArrowColor: Next i
   For i = 1 To 3: SetPixelV hdc, X + i, Index + 9, m_CheckBoxArrowColor: Next i
   For i = 5 To 7: SetPixelV hdc, X + i, Index + 9, m_CheckBoxArrowColor: Next i
   For i = 1 To 2: SetPixelV hdc, X + i, Index + 10, m_CheckBoxArrowColor: Next i
   For i = 6 To 7: SetPixelV hdc, X + i, Index + 10, m_CheckBoxArrowColor: Next i

End Sub

Private Sub DisplayCheckBoxTickMark(ByVal Index As Long)

'*************************************************************************
'* draws a tick mark in the checkbox of a selected list item when the    *
'* .Style property is set to Checkbox mode.                              *
'*************************************************************************

   Dim i As Long
   Dim X As Long

   If Not m_RightToLeft Then
      X = m_BorderWidth + 5
   Else
      X = ScaleWidth - m_BorderWidth - 16
   End If

   For i = 9 To 12: SetPixelV hdc, X + i, Index + 3, m_CheckBoxArrowColor: Next i
   For i = 8 To 11: SetPixelV hdc, X + i, Index + 4, m_CheckBoxArrowColor: Next i
   For i = 7 To 10: SetPixelV hdc, X + i, Index + 5, m_CheckBoxArrowColor: Next i
   For i = 1 To 2: SetPixelV hdc, X + i, Index + 6, m_CheckBoxArrowColor: Next i
   For i = 6 To 9: SetPixelV hdc, X + i, Index + 6, m_CheckBoxArrowColor: Next i
   For i = 1 To 3: SetPixelV hdc, X + i, Index + 7, m_CheckBoxArrowColor: Next i
   For i = 5 To 8: SetPixelV hdc, X + i, Index + 7, m_CheckBoxArrowColor: Next i
   For i = 1 To 7: SetPixelV hdc, X + i, Index + 8, m_CheckBoxArrowColor: Next i
   For i = 2 To 6: SetPixelV hdc, X + i, Index + 9, m_CheckBoxArrowColor: Next i
   For i = 3 To 5: SetPixelV hdc, X + i, Index + 10, m_CheckBoxArrowColor: Next i
   SetPixelV hdc, X + 4, Index + 11, m_CheckBoxArrowColor

End Sub

Private Sub DisplayFocusRectangle(ByVal DispIndex As Long)

'*************************************************************************
'* draws a custom focus rectangle around the specified listbox entry.    *
'* Originally I used the DrawFocusRect API, but found that the default   *
'* dotted focus rectangle was often hard to see against darker back-     *
'* grounds.  So I did this to give the user complete control over color. *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim X1     As Long        ' first x coordinate for focus rectangle.
   Dim Y1     As Long        ' first y coordinate for focus rectangle.
   Dim X2     As Long        ' second x coordinate for focus rectangle.
   Dim Y2     As Long        ' second y coordinate for focus rectange.

'  if the .ShowSelectRect property is set to False, exit.
   If Not m_ShowSelectRect Then
      Exit Sub
   End If

'  calculate left and right x coordinates of focus rectangle, accounting for .RightToLeft property.
   If Not m_RightToLeft Then
      If m_Style = [Standard] Then
         X1 = m_BorderWidth + 1
      Else
         X1 = m_BorderWidth + 21
      End If
      If m_ShowItemImages Then
         If m_ItemImageSize = 0 Then
            X1 = X1 + ListItemHeight
         Else
            X1 = X1 + m_ItemImageSize
         End If
      End If
      If Not VerticalScrollBarActive Then
         X2 = ScaleWidth - m_BorderWidth + 1
      Else
         X2 = ScaleWidth - m_BorderWidth - ScrollBarButtonWidth + 1
      End If
   Else
      X1 = m_BorderWidth + 1
      If VerticalScrollBarActive Then
         X1 = X1 + ScrollBarButtonWidth - 1
      End If
      X2 = ScaleWidth - m_BorderWidth
      If m_Style = [CheckBox] Then
         X2 = X2 - 19
      End If
      If m_ShowItemImages Then
         If m_ItemImageSize = 0 Then
            X2 = X2 - ListItemHeight
         Else
            X2 = X2 - m_ItemImageSize
         End If
      End If
   End If

'  define the top and bottom y coordinates for the focus rectangle.
   If m_ShowItemImages Then
      If ListItemHeight >= m_ItemImageSize Then
         Y1 = YCoords(DispIndex)
      Else
         Y1 = YCoords(DispIndex) + (m_ItemImageSize \ 4)
      End If
   Else
      Y1 = YCoords(DispIndex)
   End If
   Y2 = Y1 + ListItemHeight + 1

   DrawRectangle X1, Y1, X2, Y2, m_FocusRectColor

ErrHandler:
   Exit Sub

End Sub

'******************** Virtual DC Code ***************
Private Sub CreateVirtualDC(TargetDC As Long, vDC As Long, mMB As Long, mOB As Long, ByVal vWidth As Long, ByVal vHeight As Long)

'*************************************************************************
'* creates virtual bitmaps for background and cells.                     *
'*************************************************************************

   If IsCreated(vDC) Then
      DestroyVirtualDC vDC, mMB, mOB
   End If

'  create a memory device context to use.
   vDC = CreateCompatibleDC(TargetDC)

'  define it as a bitmap so that drawing can be performed to the virtual DC.
   mMB = CreateCompatibleBitmap(TargetDC, vWidth, vHeight)
   mOB = SelectObject(vDC, mMB)

End Sub

Private Function IsCreated(ByVal vDC As Long) As Boolean

'*************************************************************************
'* checks the handle of a virtual DC and returns if it exists.           *
'*************************************************************************

   IsCreated = (vDC <> 0)

End Function

Private Sub DestroyVirtualDC(ByRef vDC As Long, ByVal mMB As Long, ByVal mOB As Long)

'*************************************************************************
'* eliminates a virtual dc bitmap on control's termination.              *
'*************************************************************************

   If Not IsCreated(vDC) Then
      Exit Sub
   End If

   Call SelectObject(vDC, mOB)
   Call DeleteObject(mMB)
   Call DeleteDC(vDC)
   vDC = 0

End Sub
'********************************************************************

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<< Public Methods and Method Helper Routines >>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Sub Sort(Optional ByVal SortAscending As Boolean = True)

'*************************************************************************
'* .Sort method.  Uses a slightly tweaked version of Phillipe Lord's     *
'* TriQuickSort for maximum sorting speed for both badly unsorted and    *
'* nearly-sorted lists.  (I only tweaked it to fit seamlessly into this  *
'* control; I did not modify the sort algorithms in any way.)            *
'*************************************************************************

   Dim iLBound As Long
   Dim iUBound As Long

   iLBound = 0                'LBound(listArray)
   iUBound = m_ListCount - 1  'UBound(listArray)

'  *NOTE*  the value 4 is VERY important here !!! DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortString 4, iLBound, iUBound
   InsertionSortString iLBound, iUBound

   If Not SortAscending Then
      ReverseStringArray
   End If

'  added by MRU - when list is sorted, must redisplay from the top of the list and
'  reset ListIndex, etc
   DisplayRange.FirstListItem = 0
   If m_ListCount < MaxDisplayItems Then
      DisplayRange.LastListItem = m_ListCount - 1
   Else
      DisplayRange.LastListItem = MaxDisplayItems - 1
   End If
   If m_Style = [CheckBox] Then
      LastSelectedItem = 0
   Else
      LastSelectedItem = -1
   End If

   If m_MultiSelect = vbMultiSelectNone Then
      m_ListIndex = -1
   Else
      m_ListIndex = 0
   End If

   ItemWithFocus = 0
   DisplayList

End Sub

Private Sub TriQuickSortString(ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)

'*************************************************************************
'* the QuickSort portion of Phillipe Lord's TriQuickSort algorithm.      *
'*************************************************************************

   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   If (iMax - iMin) > iSplit Then

      i = (iMax + iMin) / 2

      If m_SortAsNumeric Then
         If Val(ListArray(iMin)) > Val(ListArray(i)) Then SwapListBoxData iMin, i ' SwapStrings ListArray(iMin), ListArray(i)
         If Val(ListArray(iMin)) > Val(ListArray(iMax)) Then SwapListBoxData iMin, iMax ' SwapStrings ListArray(iMin), ListArray(iMax)
         If Val(ListArray(i)) > Val(ListArray(iMax)) Then SwapListBoxData i, iMax ' SwapStrings ListArray(i), ListArray(iMax)
      Else 'keep
         If ListArray(iMin) > ListArray(i) Then SwapListBoxData iMin, i ' SwapStrings ListArray(iMin), ListArray(i)
         If ListArray(iMin) > ListArray(iMax) Then SwapListBoxData iMin, iMax ' SwapStrings ListArray(iMin), ListArray(iMax)
         If ListArray(i) > ListArray(iMax) Then SwapListBoxData i, iMax ' SwapStrings ListArray(i), ListArray(iMax)
      End If

      j = iMax - 1
      SwapListBoxData i, j
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(ListArray(j)), 4 ' sTemp = ListArray(j)

      Do
         If m_SortAsNumeric Then
            Do
               i = i + 1
            Loop While Val(ListArray(i)) < Val(sTemp)
            Do
               j = j - 1
            Loop While Val(ListArray(j)) > Val(sTemp)
         Else 'keep
            Do
               i = i + 1
            Loop While ListArray(i) < sTemp
            Do
               j = j - 1
            Loop While ListArray(j) > sTemp
         End If

         If j < i Then Exit Do
         SwapListBoxData i, j
      Loop

      SwapListBoxData i, iMax - 1

      TriQuickSortString iSplit, iMin, j
      TriQuickSortString iSplit, i + 1, iMax

   End If

'  clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4

End Sub

Private Sub InsertionSortString(ByVal iMin As Long, ByVal iMax As Long)

'*************************************************************************
'* the Insertion Sort portion of Phillipe Lord's TriQuickSort algorithm. *
'*************************************************************************

   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(ListArray(i)), 4 'stemp = listarray(i)
      j = i
      Do While j > iMin
         If m_SortAsNumeric Then
            If Val(ListArray(j - 1)) <= Val(sTemp) Then Exit Do
         Else
            If ListArray(j - 1) <= sTemp Then Exit Do 'keep
         End If
         CopyMemory ByVal VarPtr(ListArray(j)), ByVal VarPtr(ListArray(j - 1)), 4 'listarray(j)=listarray(j-1)
         SelectedArray(j) = SelectedArray(j - 1)
         ItemDataArray(j) = ItemDataArray(j - 1)
         If m_ShowItemImages Then
            ImageIndexArray(j) = ImageIndexArray(j - 1)
         End If
         j = j - 1
      Loop
      CopyMemory ByVal VarPtr(ListArray(j)), ByVal VarPtr(sTemp), 4 'listarray(j)=stemp
   Next i

'  clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4

End Sub

Private Sub SwapListBoxData(i As Long, j As Long)

'*************************************************************************
'* takes care of swapping all relevant listbox data based on item sort.  *
'*************************************************************************

   SwapStrings ListArray(i), ListArray(j)              ' swap .List array values.
   SwapBooleans SelectedArray(i), SelectedArray(j)     ' swap .Selected array values.
   SwapLongs ItemDataArray(i), ItemDataArray(j)        ' swap .ItemData array values.
   If m_ShowItemImages Then
      SwapLongs ImageIndexArray(i), ImageIndexArray(j)    ' swap image indexes.
   End If

End Sub

Private Sub SwapBooleans(ByRef b1 As Boolean, ByRef b2 As Boolean)

'*************************************************************************
'* swaps .Selected array values.                                         *
'*************************************************************************

   Dim bTemp As Boolean

   bTemp = b1
   b1 = b2
   b2 = bTemp

End Sub

Private Sub SwapLongs(ByRef l1 As Long, ByRef l2 As Long)

'*************************************************************************
'* swaps .Selected array values.                                         *
'*************************************************************************

   Dim lTemp As Long

   lTemp = l1
   l1 = l2
   l2 = lTemp

End Sub

Private Sub SwapStrings(ByRef s1 As String, ByRef s2 As String)

'*************************************************************************
'* helper procedure for sorting that quickly swaps two strings.          *
'*************************************************************************

   Dim i As Long

'  StrPtr() returns 0 (null) if string is not initialized
'  But StrPtr() is 5% faster than using CopyMemory, so I used that workaround, which is safe and fast.
   i = StrPtr(s1)
   If i = 0 Then
      CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4
   End If

   CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
   CopyMemory ByVal VarPtr(s2), i, 4

End Sub

Public Sub ReverseStringArray()

'*************************************************************************
'* .Sort method helper procedure for arranging list in descending order. *
'*************************************************************************

   Dim iLBound As Long
   Dim iUBound As Long

   iLBound = 0               'LBound(listArray)
   iUBound = m_ListCount - 1 'UBound(listArray)

   While iLBound < iUBound
      SwapStrings ListArray(iLBound), ListArray(iUBound)
      iLBound = iLBound + 1
      iUBound = iUBound - 1
   Wend

End Sub

Public Sub DisplayFrom(ByVal ItemIndex As Long)

'*************************************************************************
'* .DisplayFrom method.  Displays list items from ItemIndex to maximum   *
'* number of displayable list items.  If ItemIndex is anywhere within    *
'* last displayable page of the list, the entire last page is displayed. *
'*************************************************************************

   If m_Enabled Then

'     if the entire list can fit in the display area, just exit.
      If m_ListCount <= MaxDisplayItems Then
         Exit Sub
      End If

'     make sure ItemIndex actually points to an existing list item before continuing.
      If ItemIndex < 0 Or ItemIndex > m_ListCount - 1 Then
         Exit Sub
      End If

'     calculate the display range.
      DisplayRange.LastListItem = ItemIndex + MaxDisplayItems - 1
      If DisplayRange.LastListItem > m_ListCount - 1 Then
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
      Else
         DisplayRange.FirstListItem = ItemIndex
      End If

'     redisplay the new range of list items.
      DisplayList

   End If

End Sub

Public Function FindIndex(ByVal sStringToMatch As String, Optional CaseSensitive As Boolean = False) As Long

'*************************************************************************
'* .FindIndex method.  Returns the .List() index for the supplied string *
'* or -1 if the string is not found in the list.  Uses a binary search   *
'* algorithm if the .Sorted property is true; otherwise has to rely on a *
'* much slower sequential search.  If the optional CaseSensitive boolean *
'* parameter is set to True, case of supplied string must match the case *
'* of the intended target in the .List array for match to be successful. *
'*************************************************************************

   Dim i      As Long      ' loop variable.
   Dim tmpStr As String    ' string that holds current .List array item for equality comparison.
   Dim iLBound As Long     ' lower bound of list array portion currently being searched.
   Dim iUBound As Long     ' upper bound of list array portion currently being searched.
   Dim iMiddle As Long     ' middle of list array portion currently being searched.

   If m_Enabled Then

'     if we don't care about case sensitivity, make the source and target strings lower case.
      If Not CaseSensitive Then
         sStringToMatch = LCase(sStringToMatch)
      End If

      If m_Sorted Then

'        if list is sorted, a binary search can be used.
         iLBound = 0
         iUBound = m_ListCount - 1
         Do
            iMiddle = (iLBound + iUBound) \ 2
            tmpStr = ListArray(iMiddle)
            If Not CaseSensitive Then
               tmpStr = LCase(tmpStr)
            End If
            If tmpStr = sStringToMatch Then
               FindIndex = iMiddle
               Exit Function
            ElseIf tmpStr < sStringToMatch Then
               iLBound = iMiddle + 1
            Else
               iUBound = iMiddle - 1
            End If
         Loop Until iLBound > iUBound

      Else

'        if list is not sorted, a sequential search must be performed.
         For i = 0 To m_ListCount - 1
            tmpStr = ListArray(i)
            If Not CaseSensitive Then
               tmpStr = LCase(tmpStr)
            End If
            If tmpStr = sStringToMatch Then
               FindIndex = i
               Exit Function
            End If
         Next i

      End If

'     if we get here a match has not been found.
      FindIndex = -1

   Else

'     control is disabled; return -1.
      FindIndex = -1

   End If

End Function

Public Function MouseOverIndex(ByVal YPos As Single) As Long
Attribute MouseOverIndex.VB_Description = "The index of the listitem the mouse pointer is hovering over.  If the mouse pointer is not over an item, returns -1."

'*************************************************************************
'* .MouseOverIndex method.  Returns the .List() index of the item the    *
'* mouse pointer is over, based on the mouse y-coordinate and the first  *
'* displayed item's index.  Is also used internally by other usercontrol *
'* routines.  Returns -1 if mouse cursor is not over populated part of   *
'* the list or is not in list portion of control (e.g. over scrollbar).  *
'*************************************************************************

   Dim DisplayIndex As Long     ' display position in listbox.

   If m_Enabled Then

'     determine the display order index based on mouse Y coordinate.
      DisplayIndex = GetDisplayOrderIndex(YPos)

'     add that index to the index of the first displayed value.
      MouseOverIndex = DisplayRange.FirstListItem + DisplayIndex

'     safety net for below last item in list, no items at all, and mouse not in list portion.
'     the "If Not ScrollFlag" ensures that a -1 is not returned when drag scrolling.
      If Not ScrollFlag And (MouseOverIndex > m_ListCount - 1 Or Not IsInList(MouseX, MouseY)) Then
         MouseOverIndex = -1
      End If

   Else

'     control is disabled; return -1.
      MouseOverIndex = -1

   End If

End Function

Private Function GetDisplayOrderIndex(ByVal YPos As Single) As Long

'*************************************************************************
'* determines the display (YCoords array) index of the desired displayed *
'* list item, given the mouse Y coordinate.  Helper function for the     *
'* MouseOverIndex method function.                                       *
'*************************************************************************

   Dim iLBound As Long      ' lower bound of list array portion currently being searched.
   Dim iUBound As Long      ' upper bound of list array portion currently being searched.
   Dim iMiddle As Long      ' middle of list array portion currently being searched.
   Dim Done    As Boolean   ' while loop finished flag.

   iLBound = LBound(YCoords)
   iUBound = MaxDisplayItems - 2

   Done = False
   While Not Done
      iMiddle = (iLBound + iUBound) / 2
      If YPos >= YCoords(iMiddle) And YPos < YCoords(iMiddle + 1) Then
         GetDisplayOrderIndex = iMiddle
         Done = True
      ElseIf iLBound > iUBound Then
         GetDisplayOrderIndex = iLBound
         Done = True
      Else
         If YCoords(iMiddle) < YPos Then
            iLBound = iMiddle + 1
         Else
            iUBound = iMiddle - 1
         End If
      End If
   Wend

End Function

Public Sub Refresh()

'*************************************************************************
'* allows user to refresh the graphics of the control if ever necessary. *
'*************************************************************************

   If m_Enabled Then
      UserControl.Refresh
   End If

End Sub

Public Sub AddImage(ByVal ImagePath As String)

'*************************************************************************
'* .AddImage method.  Allows user to add an image to the list of images  *
'* that can be displayed next to listitems.  Adapted from a routine in   *
'* Jim Jose's "McImageList" submission at PSC, txtCodeId=62417.  Thanks  *
'* to Jim.  Note:  It is up to the programmer to keep track of the image *
'* order in project code so that it is known which image is which.       *
'*************************************************************************

   Dim mArray() As StdPicture    ' temporary image array.
   Dim NewImage As StdPicture    ' the image to add, loaded using ImagePath parameter.
   Dim i        As Long          ' loop variable.

   Set NewImage = LoadPicture(ImagePath)

   If ImageCount = 0 Then
      ReDim Images(0)
      Set Images(0) = NewImage
      ImageCount = 1
   Else
      mArray = Images
      Erase Images
      ImageCount = ImageCount + 1
      ReDim Images(0 To ImageCount - 1)
      For i = 0 To ImageCount - 2
         Set Images(i) = mArray(i)
      Next i
      Set Images(ImageCount - 1) = NewImage
   End If

End Sub

Public Sub AddItem(ByVal ItemToAdd As String, Optional ByVal Index As Long = -1)

'*************************************************************************
'* .AddItem method - adds an item/ItemData item to the list, optionally  *
'* to the given index.  If the Sorted property is True, adds to the list *
'* in the appropriate spot.  If Index parameter is supplied, this takes  *
'* precedence over Sorted property (as is the case with the intrinsic VB *
'* textbox).  If Sorted is False and no index is supplied, appends item  *
'* to end of list.                                                       *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   m_ListCount = m_ListCount + 1
   RecalculateThumbHeight = True    ' we need to recalculate thumb height since size of list has changed.

'  add the item to the list and set the .NewItem property.
   If Index = -1 Then
'     if the index is -1 (i.e. not supplied), and .Sorted = False, just append the item.
      If Not m_Sorted Then
         AddToList ItemToAdd, m_ListCount - 1
         AddToLongPropertyArray ItemDataArray(), m_ListCount - 1, 0
         AddToSelected m_ListCount - 1
         If m_ShowItemImages Then
            AddToLongPropertyArray ImageIndexArray(), m_ListCount - 1, -1
         End If
         m_NewIndex = m_ListCount - 1
      Else
'        if the index is -1 (i.e. not supplied), and .Sorted = True, insert item alphabetically.
'        if .SortAsNumeric property is .True, treat strings as numbers.
         If Not m_SortAsNumeric Then
            Index = AddToSortedList(ItemToAdd)
         Else
            Index = AddToSortedListAsNumeric(ItemToAdd)
         End If
         AddToLongPropertyArray ItemDataArray(), Index, 0
         AddToSelected Index
         If m_ShowItemImages Then
            AddToLongPropertyArray ImageIndexArray(), Index, -1
         End If
         m_NewIndex = Index
      End If
   Else
'     index has been supplied; insert the item at the indicated position.
'     NOTE: A supplied index overrides the .Sorted property (by design).
'     Therefore it is the programmer's responsibility to remember this fact and
'     to realize proper sort order will be lost if an index is supplied when .Sorted
'     is True.  Search using the .FindIndex method is also adversely affected.
      AddToList ItemToAdd, Index
      AddToLongPropertyArray ItemDataArray, Index, 0
      AddToSelected Index
      If m_ShowItemImages Then
         AddToLongPropertyArray ImageIndexArray(), Index, -1
      End If
      m_NewIndex = Index
   End If

'  if the item is not just being appended to the list, we may have to adjust the following
'  variables if they point to items that come on or after the supplied index of added item.
'  Ignored if .ListIndex = -1 or 0 and no item clicked on (i.e. list is newly initialized).
   If Not (m_ListIndex <= 0 And m_SelCount = 0) Then
      If m_ListIndex >= Index Then
         m_ListIndex = m_ListIndex + 1
      End If
      If ItemWithFocus >= Index Then
         ItemWithFocus = ItemWithFocus + 1
      End If
      If LastSelectedItem >= Index Then
         LastSelectedItem = LastSelectedItem + 1
      End If
   End If

'  since there's at least one item in the list now, activate the first display
'  item index.  The last display item index is calculated in the DrawText routine.
   If DisplayRange.FirstListItem = -1 Then
      DisplayRange.FirstListItem = 0
   End If

'  the .RedrawFlag property is used to postpone redrawing if large numbers
'  of items are added at one time.  .RemoveItem method also uses this property.
   If m_RedrawFlag Then
      DisplayList
   End If

End Sub

Private Function AddToSortedList(ByVal sToAdd As String) As Long

'*************************************************************************
'* helper routine for the .AddItem method.  Places a new list entry      *
'* into the proper place in an already-sorted listbox array.             *
'*************************************************************************

   Dim Lower As Long     ' lower bound of list array portion currently being searched.
   Dim Middle As Long    ' middle of list array portion currently being searched.
   Dim Upper As Long     ' upper bound of list array portion currently being searched.

   If m_ListCount = 1 Then ' already incremented m_listcount in AddItem method so this means empty list.
      AddToSortedList = 0
      AddToList sToAdd, 0
      Exit Function
   End If

   Lower = LBound(ListArray)
   Upper = UBound(ListArray) - 1

'  find the appropriate index to place new list item into.
   While (True)
      Middle = (Lower + Upper) / 2
      If ListArray(Middle) = sToAdd Then
         AddToSortedList = Middle
         AddToList sToAdd, Middle
         Exit Function
      ElseIf Lower > Upper Then
         AddToSortedList = Lower
         AddToList sToAdd, Lower
         Exit Function
      Else
         If ListArray(Middle) < sToAdd Then
            Lower = Middle + 1
         Else
            Upper = Middle - 1
         End If
      End If
   Wend

End Function

Private Function AddToSortedListAsNumeric(ByVal sToAdd As String) As Long

'*************************************************************************
'* helper routine for the .AddItem method.  Places a new list entry      *
'* into the proper place in an already-sorted listbox array as a numeric *
'* sort when .SortAsNumeric property is True.  Suggestion by Jeff Mayes. *
'* Note:  This is considerably slower than normal string comparison.     *
'* However, it should still load much faster than a standard VB listbox. *
'*************************************************************************

   Dim Lower    As Long     ' lower bound of list array portion currently being searched.
   Dim Middle   As Long     ' middle of list array portion currently being searched.
   Dim Upper    As Long     ' upper bound of list array portion currently being searched.
   Dim nToAdd   As Double   ' must account for large or non-whole numbers.
   Dim nCompare As Double   ' treat values already in list as doubles also.

   nToAdd = Val(sToAdd)

   If m_ListCount = 1 Then ' already incremented m_listcount in AddItem method so this means empty list.
      AddToSortedListAsNumeric = 0
      AddToList sToAdd, 0
      Exit Function
   End If

   Lower = LBound(ListArray)
   Upper = UBound(ListArray) - 1

'  find the appropriate index to place new list item into.
   While (True)
      Middle = (Lower + Upper) / 2
      nCompare = Val(ListArray(Middle))
      If nCompare = nToAdd Then
         AddToSortedListAsNumeric = Middle
         AddToList sToAdd, Middle
         Exit Function
      ElseIf Lower > Upper Then
         AddToSortedListAsNumeric = Lower
         AddToList sToAdd, Lower
         Exit Function
      Else
         If nCompare < nToAdd Then
            Lower = Middle + 1
         Else
            Upper = Middle - 1
         End If
      End If
   Wend

End Function

Private Sub AddToList(ByVal sStringToAdd As String, Optional ByVal iPos As Long = -1)

'*************************************************************************
'* helper routine for the .AddItem method.  Places a new list entry      *
'* into the specified position in a listbox array. If no index is spec-  *
'* ified, item is appended to the end of the list.  Modification of a    *
'* routine by Philippe Lord.                                             *
'*************************************************************************

   Dim iUBound As Long    ' upper bound of the List array.
   Dim iTemp   As Long    ' don't really know :)

   iUBound = UBound(ListArray)

'  if array is empty.
   If iUBound = -1 Then
      ReDim ListArray(0)
      ListArray(0) = sStringToAdd
      Exit Sub
   End If

'  if adding at the end.
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve ListArray(iUBound + 1)
      ListArray(iUBound + 1) = sStringToAdd
      Exit Sub
   End If

'  in case a negative less than -1 is erroneously passed.
   If iPos < 0 Then
      iPos = 0
   End If

'  increase size of array by one element.
   iUBound = iUBound + 1
   ReDim Preserve ListArray(iUBound)

   CopyMemory ByVal VarPtr(ListArray(iPos + 1)), ByVal VarPtr(ListArray(iPos)), (iUBound - iPos) * 4

   iTemp = 0 ' view this as String(4, Chr(0)) or a NULL value
   CopyMemory ByVal VarPtr(ListArray(iPos)), iTemp, 4

   ListArray(iPos) = sStringToAdd

End Sub

Private Sub AddToLongPropertyArray(PropArray() As Long, ByVal iPos As Long, ByVal InitialValue As Long)

'*************************************************************************
'* helper routine for the .AddItem method.  places a new ItemData entry  *
'* into the specified position in an ItemData or ImageIndex long array.  *
'* Modification of a routine by Philippe Lord.                           *
'*************************************************************************

   Dim iUBound As Long    ' upper bound of the ItemData or ImageIndex array.

   iUBound = UBound(PropArray)

'  if array is empty.
   If iUBound = -1 Then
      ReDim PropArray(0)
      PropArray(0) = InitialValue
      Exit Sub
   End If

'  if adding at the end.
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve PropArray(iUBound + 1)
      PropArray(iUBound + 1) = InitialValue
      Exit Sub
   End If

'  in case a negative index is erroneously passed.
   If iPos < 0 Then
      iPos = 0
   End If

'  increase size of array by one element.
   iUBound = iUBound + 1
   ReDim Preserve PropArray(iUBound)

   CopyMemory PropArray(iPos + 1), PropArray(iPos), (iUBound - LBound(PropArray) - iPos) * Len(PropArray(iPos))
   PropArray(iPos) = InitialValue

End Sub

Private Sub AddToSelected(ByVal iPos As Long)

'*************************************************************************
'* helper routine for the .AddItem method.  Adds a new .Selected() entry *
'* into the specified position in a Selected boolean array. Modification *
'* of a routine by Philippe Lord.                                        *
'*************************************************************************

   Dim iUBound As Long    ' upper bound of the Selected array.

   iUBound = UBound(SelectedArray)

   If iUBound = -1 Then
      ReDim SelectedArray(0)
      SelectedArray(0) = False
      Exit Sub
   End If

' if adding at the end.
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve SelectedArray(iUBound + 1)
      SelectedArray(iUBound + 1) = False
      Exit Sub
   End If

'  in case a negative is erroneously passed.
   If iPos < 0 Then
      iPos = 0
   End If

'  increase size of array by one element.
   iUBound = iUBound + 1
   ReDim Preserve SelectedArray(iUBound)

   CopyMemory SelectedArray(iPos + 1), SelectedArray(iPos), (iUBound - LBound(SelectedArray) - iPos) * Len(SelectedArray(iPos))
   SelectedArray(iPos) = False

End Sub

Public Sub RemoveItem(ByVal Index As Long)

'*************************************************************************
'* .RemoveItem method - removes the specified item from the List,        *
'* ItemData, Selected, and ImageIndex property arrays.                   *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   If m_ListCount > 0 Then

      If Index >= LBound(ListArray) And Index <= UBound(ListArray) Then

'        reduce the .ListCount property variable.
         m_ListCount = m_ListCount - 1
         RecalculateThumbHeight = True   ' we need to recalculate thumb height since size of list has changed.

'        if the item to remove is selected, decrease the m_SelCount property variable.
         If SelectedArray(Index) Then
            m_SelCount = m_SelCount - 1
         End If

'        remove from the property arrays.
         RemoveFromListArray Index
         RemoveFromLongPropertyArray ItemDataArray(), Index ' RemoveFromItemDataArray Index
         RemoveFromSelectedArray Index
         RemoveFromLongPropertyArray ImageIndexArray(), Index ' RemoveFromImageIndexArray Index

'        adjust the .ListIndex property variable.
'        if .ListIndex points to the item that was just removed, set it to cleared status per mode.
         If m_ListIndex = Index Then
            If m_Style = [CheckBox] Or m_MultiSelect <> vbMultiSelectNone Then
               m_ListIndex = 0
            Else
               m_ListIndex = -1
            End If
         Else
'           if .ListIndex comes after deleted item we must decrement
'           it to reflect the ListIndex item's new array position.
            If m_ListIndex > Index Then
               m_ListIndex = m_ListIndex - 1
            End If
         End If

'        adjust the LastSelectedItem internal variable.
'        if it points to the item that was just removed, clear it.
         If LastSelectedItem = Index Then
            LastSelectedItem = -1
         Else
'           if it comes after deleted item we must decrement
'           it to reflect the item's new array position.
            If LastSelectedItem > Index Then
               LastSelectedItem = LastSelectedItem - 1
            End If
         End If

'        adjust the ItemWithFocus internal variable.
'        if it points to the item that was just removed, clear it.
         If ItemWithFocus = Index Then
            ItemWithFocus = 0
         Else
'           if it comes after deleted item we must decrement
'           it to reflect the item's new array position.
            If ItemWithFocus > Index Then
               ItemWithFocus = ItemWithFocus - 1
            End If
         End If

'        since an item has been removed, we must set the .NewIndex property to -1.
         m_NewIndex = -1

'        as with the .AddItem method, the .RedrawFlag property can be used to postpone
'        redrawing the control if a large number of items are being removed from the list.
         If m_RedrawFlag Then
            DisplayList
         End If

      End If

   End If

End Sub

Private Sub RemoveFromListArray(ByVal iPos As Long)

'*************************************************************************
'* helper routine for the RemoveItem method - removes the specified item *
'* from the List array.  Modification of a routine by Philippe Lord.     *
'*************************************************************************

   Dim iLBound As Long     ' lower bound of List array.
   Dim iUBound As Long     ' upper bound of List array.
   Dim iTemp   As Long     ' pointer to address of List array element.

   iLBound = LBound(ListArray)
   iUBound = UBound(ListArray)

'  if we only have one element in array.
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then
      Erase ListArray
      ReDim ListArray(0)
      Exit Sub
   End If

'  if invalid iPos - might not need 1st two checks now.
   If (iPos > iUBound) Or (iPos = -1) Then
      iPos = iUBound
   End If
   If iPos < iLBound Then
      iPos = iLBound
   End If
   If iPos = iUBound Then
      ReDim Preserve ListArray(iUBound - 1)
      Exit Sub
   End If

   iTemp = StrPtr(ListArray(iPos))
   CopyMemory ByVal VarPtr(ListArray(iPos)), ByVal VarPtr(ListArray(iPos + 1)), (iUBound - iPos) * 4

'  do this to have VB deallocate the string; avoids memory leaks.
   CopyMemory ByVal VarPtr(ListArray(iUBound)), iTemp, 4

   ReDim Preserve ListArray(iUBound - 1)

End Sub

Private Sub RemoveFromLongPropertyArray(PropArray() As Long, ByVal iPos As Long)

'*************************************************************************
'* helper routine for the .RemoveItem method - removes specified item    *
'* from the ItemData or ImageIndex arrays.  Modification of a routine by *
'* Philippe Lord.                                                        *
'*************************************************************************

   Dim iLBound As Long   ' lower bound of ItemData or ImageIndex array.
   Dim iUBound As Long   ' upper bound of ItemData or ImageIndex array.

   iLBound = LBound(PropArray)
   iUBound = UBound(PropArray)

'  if we only have one element in array.
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then
      Erase PropArray
      ReDim PropArray(0)
      Exit Sub
   End If

'  if invalid iPos.
   If (iPos > iUBound) Or (iPos = -1) Then
      iPos = iUBound
   End If
   If iPos < iLBound Then
      iPos = iLBound
   End If
   If iPos = iUBound Then
      ReDim Preserve PropArray(iUBound - 1)
      Exit Sub
   End If

   CopyMemory PropArray(iPos), PropArray(iPos + 1), (iUBound - iLBound - iPos) * Len(PropArray(iPos))

   ReDim Preserve PropArray(iUBound - 1)

End Sub

Private Sub RemoveFromSelectedArray(ByVal iPos As Long)

'*************************************************************************
'* helper routine for the .RemoveItem method - removes specified item    *
'* from the Selected array.  Modification of a routine by Philippe Lord. *
'*************************************************************************

   Dim iLBound As Long    ' lower bound of .Selected() array.
   Dim iUBound As Long    ' upper bound of .Selected() array.

   iLBound = LBound(SelectedArray)
   iUBound = UBound(SelectedArray)

'  if we only have one element in array.
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then
      Erase SelectedArray
      ReDim SelectedArray(0)
      Exit Sub
   End If

'  if invalid iPos.
   If (iPos > iUBound) Or (iPos = -1) Then
      iPos = iUBound
   End If
   If iPos < iLBound Then
      iPos = iLBound
   End If
   If iPos = iUBound Then
      ReDim Preserve SelectedArray(iUBound - 1)
      Exit Sub
   End If

   CopyMemory SelectedArray(iPos), SelectedArray(iPos + 1), (iUBound - iLBound - iPos) * Len(SelectedArray(iPos))

   ReDim Preserve SelectedArray(iUBound - 1)

End Sub

Public Sub Clear()

'*************************************************************************
'* the .Clear method for the listbox; removes all entries.               *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

'  re-initialize the four property arrays.
   ReDim ListArray(0)
   ReDim ItemDataArray(0)
   ReDim SelectedArray(0)
   ReDim ImageIndexArray(0)

'  set the appropriate property and internal values to initialized state.
   m_SelCount = 0
   m_ListCount = 0
   LastSelectedItem = -1
   ItemWithFocus = 0
   DisplayRange.FirstListItem = -1
   DisplayRange.LastListItem = -1
   m_RedrawFlag = True
   VerticalScrollBarActive = False
   RecalculateThumbHeight = True    ' we need to recalculate thumb height since size of list has changed.

'  in CheckBox, MultiSelect Simple and MultiSelect Extended modes, the
'  ListIndex property is 0 in a cleared list but -1 in MultiSelect None mode.
   If m_Style = [CheckBox] Or m_MultiSelect <> vbMultiSelectNone Then
      m_ListIndex = 0
   Else
      m_ListIndex = -1
   End If

'  since the listbox has been cleared, we must set the .NewIndex property to -1.
   m_NewIndex = -1

'  redraw the background (and border, if picture) onto the usercontrol DC.
   SetBackGround
   If IsPictureThere(m_ActivePicture) Then
      CreateBorder
   End If
   UserControl.Refresh

End Sub

Public Sub ClearOrSelect(Optional ByVal StartIdx As Long = -1, Optional ByVal EndIdx As Long = -1, Optional ByVal SelectFlag As Boolean = False)

'*************************************************************************
'* public method that allows user to select or deselect all listitems    *
'* within a given range.  If SelectFlag parameter is True, range is sel- *
'* ected; otherwise it is deselected.  If no range is supplied, all      *
'* listitems are selected or deselected.  User is responsible for supp-  *
'* lying start and end indices that account for list being zero-based.   *
'*************************************************************************

   Dim Temp           As Long       ' swap variable.
   Dim NumberSelected As Long       ' num selected in range.
   Dim NumNotSelected As Long       ' num not selected in range.
   Dim i              As Long       ' loop variable.
   Dim Subset         As Boolean    ' range selected flag.

'  catch situation where one index was not set.
   If StartIdx * EndIdx < 0 Then
      Exit Sub    ' a MsgBox error message could be supplied here if desired.
   End If

'  make sure start index is less than end index.
   If StartIdx > -1 And EndIdx > -1 Then
      If StartIdx > EndIdx Then
         Temp = EndIdx
         EndIdx = StartIdx
         StartIdx = Temp
      End If
   End If

'  make sure end index is no larger than the number of listitems (accounting for zero-based array).
   If EndIdx > m_ListCount - 1 Then
      EndIdx = m_ListCount - 1
   End If

'  if start and end indices are -1, (i.e. not supplied by user) adjust to entire list.
   If StartIdx = -1 And EndIdx = -1 Then
      StartIdx = 0
      EndIdx = m_ListCount - 1
   End If

   If m_ListCount > 0 And ((Not SelectFlag And m_SelCount > 0) Or SelectFlag) Then

'     if a subset of the list has been specified for selection/deselection, we must count the number
'     of selected/deselected items within the range to properly reset m_SelCount property variable.
'     For very large ranges this can be time consuming.  If anyone knows of an API-based alternative
'     (I couldn't find one in API Guide) please leave feedback on this control's Planet Source Code page!
      If EndIdx - StartIdx + 1 < m_ListCount Then
         Subset = True
         If SelectFlag Then
            For i = StartIdx To EndIdx
               If Not SelectedArray(i) Then
                  NumNotSelected = NumNotSelected + 1
               End If
            Next i
         Else
            For i = StartIdx To EndIdx
               If SelectedArray(i) Then
                  NumberSelected = NumberSelected + 1
               End If
            Next i
         End If
      End If

'     reinitialize the selected array to all False.
      If SelectFlag Then
         SetSelectedArrayRange StartIdx, EndIdx, True
      Else
         SetSelectedArrayRange StartIdx, EndIdx, False
      End If

'     adjust the .SelCount property variable.
      If Subset Then
         If SelectFlag Then
            m_SelCount = m_SelCount + NumNotSelected
         Else
            m_SelCount = m_SelCount - NumberSelected
         End If
      Else
         If SelectFlag Then
            m_SelCount = m_ListCount
         Else
            m_SelCount = 0
         End If
      End If

'     redisplay the list.
      DisplayList

   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Miscellaneous ListBox Helper Functions >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ProcessSelectedItem()

'*************************************************************************
'* handles pointers and painting of a newly selected listbox item        *
'* according to the .MultiSelect property (Single, Simple, Extended).    *
'*************************************************************************

   Select Case m_MultiSelect

      Case vbMultiSelectNone
         ProcessSelected_MultiSelectNone

      Case vbMultiSelectSimple
         ProcessSelected_MultiSelectSimple

      Case vbMultiSelectExtended
         ProcessSelected_MultiSelectExtended

   End Select

End Sub

Private Sub ProcessSelected_MultiSelectNone()

'*************************************************************************
'* performs operations to select an item in MultiSelect None mode.       *
'*************************************************************************

'  reinitialize the Selected array to all False. (Remember, no multiselect here.)
   SetSelectedArrayRange 0, m_ListCount - 1, False

'  set the .ListIndex property to the index of the selected item.
   m_ListIndex = LastSelectedItem

'  set the clicked item's selected status to True.
   SelectedArray(LastSelectedItem) = True
   ItemWithFocus = LastSelectedItem
   m_SelCount = 1

   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessSelected_MultiSelectSimple()

'*************************************************************************
'* performs operations to select an item in MultiSelect Simple mode.     *
'*************************************************************************

'  set the clicked item's selected status to True.
   SelectedArray(LastSelectedItem) = True
   ItemWithFocus = LastSelectedItem ' can take out?

'  in Simple mode, the .ListIndex property is ALWAYS the index of the focused item, selected or not.
   m_ListIndex = ItemWithFocus
   m_SelCount = m_SelCount + 1

   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessSelected_MultiSelectExtended(Optional ByVal SelCountFlag As Boolean = True)

'*************************************************************************
'* performs operations to select an item in MultiSelect Extended mode.   *
'*************************************************************************

'  set the .ListIndex property to the index of the selected item.
   ItemWithFocus = LastSelectedItem
   m_ListIndex = ItemWithFocus

'  set the clicked item's selected status to True.
   SelectedArray(LastSelectedItem) = True
   If SelCountFlag Then
      m_SelCount = m_SelCount + 1
   End If

   DisplayListBoxItem LastSelectedItem, DrawAsSelected, FocusRectangleYes
   UserControl.Refresh

End Sub

Private Sub ProcessContinuousScroll(ScrollAmount As Long)

'*************************************************************************
'* performs a continuous vertical scroll, scrolling by specified amount. *
'* this routine handles all vertical scrollbar continuous scrolling,     *
'* (up arrow, down arrow, mouse above/below listbox, and trackbar).      *
'*************************************************************************

   Dim OriginalTickCount As Long   ' comparison tick count for calculating elapsed time.
   Dim CurrentTickCount  As Long   ' current time (tick count).
   Dim LastItemIndex     As Long   ' last item for down scroll, first item for up scroll.

'  based on scroll direction, determine the end item to check for during scroll.
   LastItemIndex = IIf(ScrollAmount > 0, m_ListCount - 1, 0)

'  create a preliminary delay before list starts scrolling, to give the user time to unclick the
'  mouse button and prevent scrolling.  If user clicked in list area and is scrolling by dragging
'  the mouse above or below the listbox, ScrollFlag will be True and this initial delay is ignored.
   If Not ScrollFlag Then
      OriginalTickCount = GetTickCount
      CurrentTickCount = OriginalTickCount
      While MouseAction <> MOUSE_NOACTION And CurrentTickCount - OriginalTickCount < INITIAL_SCROLL_DELAY
         CurrentTickCount = GetTickCount
         DoEvents
      Wend
   End If

'  if the DoEvents from the above delay loop did not reveal a MouseUp event, the display-scrolling
'  loop below can execute (MouseAction will be <> MOUSE_NOACTION).  Loop until the mouse button is
'  unclicked or the top or bottom of the list has been reached, depending on scroll direction.
'  Scroll interval (SCROLL_TICKCOUNT) is 50 milliseconds.
   OriginalTickCount = GetTickCount
   While MouseAction <> MOUSE_NOACTION And (Not InDisplayedItemRange(LastItemIndex))

'     allow an opportunity for a MouseUp event to stop the scrolling.
      DoEvents
'     keeps cpu usage from maxing out.  Thanks to Mike Douglas for the tip.
      Sleep 25

'     this 'If' statement allows control to process trackbar scrolling like the VB listbox -
'     scrolling will continue until thumb is under mouse cursor, at which point scrolling stops.
      If ScrollFlag Or MouseLocation = OVER_VTRACKBAR Or MouseLocation = OVER_UPBUTTON Or MouseLocation = OVER_DOWNBUTTON Then

'        get the current time.
         CurrentTickCount = GetTickCount

'        perform a scroll if at least SCROLL_TICKCOUNT milliseconds have passed.
         If CurrentTickCount - OriginalTickCount >= SCROLL_TICKCOUNT Then

'           adjust the display range up or down according to scroll direction.
            DisplayRange.FirstListItem = DisplayRange.FirstListItem + ScrollAmount
            If DisplayRange.FirstListItem < 0 Then
               DisplayRange.FirstListItem = 0
            End If
            DisplayRange.LastListItem = DisplayRange.LastListItem + ScrollAmount
            If DisplayRange.LastListItem > m_ListCount - 1 Then
               DisplayRange.LastListItem = m_ListCount - 1
            End If

'           if we're page scrolling (i.e. scrolling using trackbar), make sure end of list is
'           displayed correctly with the last list item at the physical bottom of the control.
            If Abs(ScrollAmount) > 1 And DisplayRange.FirstListItem + MaxDisplayItems - 1 > m_ListCount Then
               DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
            End If

'           if page scrolling via trackbar, and mouse pointer moves out of section of trackbar
'           that determined scroll direction (e.g. mouse is moved below thumb when scrolling up
'           due to mouse being originally clicked in track above thumb) stop scrolling.
            If Abs(ScrollAmount) > 1 Then
               If MouseCursorIsAboveThumb(MouseY) And MouseAction = MOUSE_DOWNED_IN_LOWERTRACK Then
                  Exit Sub
               End If
               If Not MouseCursorIsAboveThumb(MouseY) And MouseAction = MOUSE_DOWNED_IN_UPPERTRACK Then
                  Exit Sub
               End If
            End If

'           if we're scrolling by clicking on the list and then dragging the mouse above
'           or below the listbox, the top or bottom displayed list item (depending on
'           scrolling direction) always has the highlight gradient (unless in MultiSelect
'           Simple mode, where just the focus rectangle is used).
            If ScrollFlag Then
               If ScrollAmount > 0 Then   ' if scrolling down (mouse dragged below listbox)...
                  If m_MultiSelect <> vbMultiSelectSimple Then
                     LastSelectedItem = DisplayRange.LastListItem
                  End If
               Else                       ' if scrolling up (mouse dragged above listbox)...
                  If m_MultiSelect <> vbMultiSelectSimple Then
                     LastSelectedItem = DisplayRange.FirstListItem
                  End If
               End If
               ProcessMouseMoveItemSelection
            End If

'           time to redisplay the list after display range adjustments.
            DisplayList

'           store the current time and wait for the next SCROLL_TICKCOUNT milliseconds.
            OriginalTickCount = GetTickCount

         End If

      End If

   Wend

End Sub

Private Sub SetSelectedArrayRange(ByVal FirstValue As Long, ByVal LastValue As Long, ByVal bSelectedStatus As Boolean)

'*************************************************************************
'* this procedure sets the given range of elements in the SelectedArray  *
'* array to either True or False using the FillMemory API.  It is used   *
'* when .MultiSelect is Extended and extremely large ranges of list      *
'* items' Selected status must be set very quickly (for example, when    *
'* 25000 list items must all be selected at once due clicking the first  *
'* item and then pressing Shift-End).  It is also used to instantly set  *
'* the entire array to unselected (False).                               *
'*************************************************************************

   Dim Temp As Long     ' swap temporary variable.

'  the way I do things in this control, FirstValue might be greater than LastValue.
'  For example, when doing a Shift-PageUp in MultiSelect Extended mode, this will be
'  the case.  To correctly use the FillMemory API,  we must swap FirstValue and
'  LastValue in these circumstances.
   If FirstValue > LastValue Then
      Temp = FirstValue
      FirstValue = LastValue
      LastValue = Temp
   End If
On Error Resume Next
   FillMemory SelectedArray(FirstValue), 2 * (LastValue - FirstValue + 1), bSelectedStatus

End Sub

Private Sub AdjustDisplayRange()

'*************************************************************************
'* the following code helps emulate the vb listbox when the focused item *
'* is out of displayed range (such as when scrollbar is used to navigate *
'* up or down the list), and an arrow key, page key, etc. is pressed.    *
'* MorphListBox display range is adjusted according to vb listbox rules. *
'*************************************************************************

   If ItemWithFocus < DisplayRange.FirstListItem Then
'     if the list item with the focus is above the first displayed item, the display
'     will adjust so that the focused item is at the top of the displayed range.
      DisplayRange.FirstListItem = ItemWithFocus
      If DisplayRange.FirstListItem + MaxDisplayItems - 1 <= m_ListCount Then
         DisplayRange.LastListItem = ItemWithFocus + MaxDisplayItems - 1
      Else
         DisplayRange.LastListItem = m_ListCount - 1
      End If
      DisplayList
   Else
      If ItemWithFocus > DisplayRange.LastListItem Then
'        if the list item with the focus is below the last displayed item, the display
'        will adjust so that the focused item is at the bottom of the displayed range.
         DisplayRange.LastListItem = ItemWithFocus
         If DisplayRange.LastListItem - MaxDisplayItems + 1 >= 0 Then
            DisplayRange.FirstListItem = DisplayRange.LastListItem - MaxDisplayItems + 1
         Else
            DisplayRange.FirstListItem = 0
         End If
         DisplayList
      End If
   End If

End Sub

Private Function GetDisplayIndexFromArrayIndex(ArrayIndex As Long) As Long

'*************************************************************************
'* given the item's array index, returns the display (YCoords) index,    *
'* or returns -1 if the item is not in the display range.                *
'*************************************************************************

   Dim iLBound      As Long        ' lower bound of display range currently being searched.
   Dim iUBound      As Long        ' upper bound of display range currently being searched.
   Dim iMiddle      As Long        ' middle of display range currently being searched.
   Dim CompareIndex As Long        ' current display range index currently being examined.

   iLBound = 0
   iUBound = DisplayRange.LastListItem - DisplayRange.FirstListItem + 1

   Do
      iMiddle = (iLBound + iUBound) \ 2
      CompareIndex = DisplayRange.FirstListItem + iMiddle - 1
      If CompareIndex = ArrayIndex Then
         GetDisplayIndexFromArrayIndex = iMiddle - 1
         Exit Function
      ElseIf CompareIndex < ArrayIndex Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If
   Loop Until iLBound > iUBound

   GetDisplayIndexFromArrayIndex = -1

End Function

Private Sub CalculateSelCount()

'*************************************************************************
'* I don't like recalculating m_SelCount with this brute force method.   *
'* But (for example) in MultiSelect Extended mode using the Shift-Home   *
'* or Shift-End  keys, with other items or groups of items possibly hav- *
'* ing been selected in other parts of the list (or even the part of the *
'* list affected by a Shift-Home/Shift-End), this is the most logical    *
'* way.  Still very fast though.                                         *
'*************************************************************************

   Dim i      As Long    ' loop variable.
   Dim EndVal As Long    ' last element of the .Selected array.

   m_SelCount = 0
   EndVal = m_ListCount - 1

   For i = 0 To EndVal
'     remember that True = -1 and False = 0.  So the absolute value
'     of the sum of the elements of the SelectedArray can be used.
      m_SelCount = m_SelCount + SelectedArray(i)
   Next i

   m_SelCount = Abs(m_SelCount)

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Vertical ScrollBar >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub DisplayVerticalScrollBar(Optional vThumbYPos As Single = -1)

'**************************************************************************
'* displays the vertical scrollbar.                                       *
'**************************************************************************

   DisplayVerticalTrackBar
   DisplayTrackBarButton UPBUTTON
   DisplayTrackBarButton DOWNBUTTON
   DisplayVerticalThumb vThumbYPos

'  get the information so that mouse cursor location in scrollbar can be determined.
   GetVScrollbarLocationInfo

End Sub

Private Sub GetVScrollbarLocationInfo()

'*************************************************************************
'* gets position info for all parts of vertical scrollbar except the     *
'* thumb, which is calculated on-the-fly when the thumb is drawn.        *
'*************************************************************************

   Dim vLeft As Long     ' x position of left side of scrollbar.
   Dim vRight As Long    ' x position of right side of scrollbar.

   If Not m_RightToLeft Then
      vLeft = ScaleWidth - m_BorderWidth - ScrollBarButtonWidth - 1
   Else
      vLeft = m_BorderWidth
   End If
   vRight = vLeft + ScrollBarButtonWidth - 1

   vScrollBarLocation.UpButtonLocation.Top = m_BorderWidth
   vScrollBarLocation.UpButtonLocation.Left = vLeft
   vScrollBarLocation.UpButtonLocation.Bottom = vScrollBarLocation.UpButtonLocation.Top + ScrollBarButtonHeight - 1
   vScrollBarLocation.UpButtonLocation.Right = vRight

   vScrollBarLocation.DownButtonLocation.Top = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight
   vScrollBarLocation.DownButtonLocation.Left = vLeft
   vScrollBarLocation.DownButtonLocation.Bottom = vScrollBarLocation.DownButtonLocation.Top + ScrollBarButtonHeight - 1
   vScrollBarLocation.DownButtonLocation.Right = vRight

   vScrollBarLocation.ScrollTrackLocation.Top = m_BorderWidth + ScrollBarButtonHeight
   vScrollBarLocation.ScrollTrackLocation.Left = vLeft '- m_BorderWidth
   vScrollBarLocation.ScrollTrackLocation.Bottom = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight - 1
   vScrollBarLocation.ScrollTrackLocation.Right = vRight

End Sub

Private Function MouseCursorIsAboveThumb(YPos As Single) As Boolean

'*************************************************************************
'* when mouse is clicked on vertical scroll bar trackbar, need to deter- *
'* mine if the mouse is above or below the scroll thumb so a page up or  *
'* page down can be performed.                                           *
'*************************************************************************

   If YPos < vScrollBarLocation.ScrollThumbLocation.Top Then
      MouseCursorIsAboveThumb = True
   End If

End Function

Private Sub DisplayVerticalThumb(Optional Y As Single = -1)

'*************************************************************************
'* displays the vertical scrollbar's thumb scroller.                     *
'*************************************************************************

   Dim XPos As Long    ' the x position for left side of thumb.
   Dim YPos As Long    ' the y position for the top of the thumb.

   If Not m_RightToLeft Then
      XPos = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth
   Else
      XPos = m_BorderWidth
   End If

'  obtain the height of the scroller thumb.  Save processing
'  time by only recalculating when list size has changed.
   If RecalculateThumbHeight Then
      vThumbHeight = CalculateVScrollThumbHeight
   End If

'  get the top y pos of the vertical scroller thumb.
   If Y = -1 Then
'     if no cursor position in thumb is passed, then just calculate
'     the thumb's top y coordinate based on where we are in the list.
      YPos = ThumbYPos ' calculated in DisplayVerticalTrackbar.
   Else
'     otherwise, calculate based on the current mouse pos and where
'     it is in the thumb.  This is only done when dragging the thumb.
      YPos = MouseY - Y
'     these 'if' statements ensure the thumb stays between the up and down buttons.
      If YPos + vThumbHeight > vScrollBarLocation.DownButtonLocation.Top Then
         YPos = vScrollBarLocation.DownButtonLocation.Top - vThumbHeight
      End If
      If YPos < vScrollBarLocation.UpButtonLocation.Bottom Then
         YPos = vScrollBarLocation.UpButtonLocation.Bottom + 1
      End If
   End If

'  define the vertical thumb's rectangle coordinates.
   vScrollBarLocation.ScrollThumbLocation.Top = YPos
   vScrollBarLocation.ScrollThumbLocation.Left = XPos
   vScrollBarLocation.ScrollThumbLocation.Bottom = YPos + vThumbHeight - 1
   vScrollBarLocation.ScrollThumbLocation.Right = vScrollBarLocation.ScrollThumbLocation.Left + ScrollBarButtonWidth - 1

'  display the thumb.
   Call StretchDIBits(hdc, _
                      vScrollBarLocation.ScrollThumbLocation.Left, YPos, _
                      ScrollBarButtonWidth, vThumbHeight, _
                      0, 0, _
                      ScrollBarButtonWidth, vThumbHeight, _
                      vThumblBits(0), vThumbuBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

'  draw the thumb's border.
   DisplayVerticalThumbBorder

End Sub

Private Sub DisplayVerticalThumbBorder()

'*************************************************************************
'* draws a one-pixel wide border around the vertical scrollbar thumb.    *
'*************************************************************************

   DrawRectangle vScrollBarLocation.ScrollThumbLocation.Left + 1, _
                 vScrollBarLocation.ScrollThumbLocation.Top, _
                 vScrollBarLocation.ScrollThumbLocation.Right + 1, _
                 vScrollBarLocation.ScrollThumbLocation.Bottom + 1, _
                 m_ActiveThumbBorderColor

End Sub

Private Function VerticalThumbY() As Long

'*************************************************************************
'* determines the y coordinate of the top of the vertical scrollbar's    *
'* thumb, so thumb will be displayed in the right place in the track.    *
'*************************************************************************

   Dim PixelsPerScroll       As Single  ' how many pixels involved in scrolling one list item.
   Dim NumClicks             As Long    ' basically, how many clicks it takes to get to list end.

'  calculate the thumb middle pixel's range of motion in the track.
   vThumbRange.Top = m_BorderWidth + ScrollBarButtonHeight + (vThumbHeight / 2) + 1
   vThumbRange.Bottom = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight - (vThumbHeight / 2) + 1

'  if we're at the very top or bottom of the list, this is easy.
   If DisplayRange.FirstListItem = 0 Then
      VerticalThumbY = m_BorderWidth + ScrollBarButtonHeight
      Exit Function
   End If
   If DisplayRange.LastListItem = m_ListCount - 1 Then
      VerticalThumbY = ScaleHeight - m_BorderWidth - vThumbHeight - ScrollBarButtonHeight
      If VerticalThumbY < m_BorderWidth + ScrollBarButtonHeight Then
         VerticalThumbY = m_BorderWidth + ScrollBarButtonHeight
      End If
      Exit Function
   End If

'  determine how by many items the list can be scrolled.
   NumClicks = m_ListCount - MaxDisplayItems

'  how many pixels of scroll thumb motion per list item scroll?
   PixelsPerScroll = (vThumbRange.Bottom - vThumbRange.Top) / NumClicks
'  how many pixels into the thumb's middle pixel motion range are we?
   PixelsPerScroll = PixelsPerScroll * DisplayRange.FirstListItem

   VerticalThumbY = (vThumbRange.Top + PixelsPerScroll) - vThumbHeight / 2

End Function

Private Sub DisplayVerticalTrackBar()

'*************************************************************************
'* displays the vertical scrollbar trackbar between up and down buttons. *
'*************************************************************************

   Dim tbX As Long

   If Not m_RightToLeft Then
      tbX = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth
   Else
      tbX = m_BorderWidth
   End If

'  determine the y position of the top of the vertical scrollbar thumb.
'  This is also used in the DisplayVerticalThumb routine. It is calculated
'  here so that if the scroll track under the thumb is being clicked down,
'  the correct height of trackbar under the thumb is highlighted.
   ThumbYPos = VerticalThumbY

   Call StretchDIBits(hdc, _
                      tbX, _
                      m_BorderWidth + ScrollBarButtonHeight, _
                      ScrollBarButtonWidth, _
                      vScrollTrackHeight, _
                      0, 0, _
                      ScrollBarButtonWidth, _
                      vScrollTrackHeight, _
                      VTracklBits(0), _
                      VTrackuBIH, _
                      DIB_RGB_COLORS, _
                      vbSrcCopy)

'  if the mouse is currently clicked down in the track, repaint that portion accordingly.
   If MouseAction = MOUSE_DOWNED_IN_UPPERTRACK Then

      Call StretchDIBits(hdc, _
                         tbX, _
                         m_BorderWidth + ScrollBarButtonHeight, _
                         ScrollBarButtonWidth, _
                         vScrollBarLocation.ScrollThumbLocation.Top - m_BorderWidth - ScrollBarButtonHeight, _
                         0, 0, _
                         ScrollBarButtonWidth, _
                         vScrollTrackHeight, _
                         vClickTracklBits(0), _
                         vClickTrackuBIH, _
                         DIB_RGB_COLORS, _
                         vbSrcCopy)

   ElseIf MouseAction = MOUSE_DOWNED_IN_LOWERTRACK Then    'y,dy

      Call StretchDIBits(hdc, _
                         tbX, _
                         ThumbYPos, _
                         ScrollBarButtonWidth, _
                         ScaleHeight - ThumbYPos - ScrollBarButtonHeight - m_BorderWidth, _
                         0, 0, _
                         ScrollBarButtonWidth, _
                         vScrollTrackHeight, _
                         vClickTracklBits(0), _
                         vClickTrackuBIH, _
                         DIB_RGB_COLORS, _
                         vbSrcCopy)

   End If

End Sub

Private Sub DisplayTrackBarButton(ByVal WhichButton As Long)

'*************************************************************************
'* displays the appropriate scrollbar button.                            *
'*************************************************************************

   Dim tbX As Long

   If Not m_RightToLeft Then
      tbX = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth
   Else
      tbX = m_BorderWidth
   End If

   Select Case WhichButton

      Case UPBUTTON
         Call StretchDIBits(hdc, _
                            tbX, _
                            m_BorderWidth, _
                            ScrollBarButtonWidth, _
                            ScrollBarButtonHeight, _
                            0, 0, _
                            ScrollBarButtonWidth, _
                            ScrollBarButtonHeight, _
                            TrackButtonlBits(0), _
                            TrackButtonuBIH, _
                            DIB_RGB_COLORS, _
                            vbSrcCopy)
         DrawScrollButtonArrow UPBUTTON
   
      Case DOWNBUTTON
         Call StretchDIBits(hdc, _
                            tbX, _
                            ScaleHeight - m_BorderWidth - ScrollBarButtonHeight, _
                            ScrollBarButtonWidth, _
                            ScrollBarButtonHeight, _
                            0, 0, _
                            ScrollBarButtonWidth, _
                            ScrollBarButtonHeight, _
                            TrackButtonlBits(0), _
                            TrackButtonuBIH, _
                            DIB_RGB_COLORS, _
                            vbSrcCopy)
         DrawScrollButtonArrow DOWNBUTTON

   End Select

End Sub

Private Sub DrawScrollButtonArrow(ByVal WhichButton As Long)

'*************************************************************************
'* draws both the up and down arrows on vertical scrollbar buttons.      *
'*************************************************************************

   Dim hPO           As Long       ' selected pen object.
   Dim hPN           As Long       ' pen object for drawing checkmark.
   Dim r             As Long       ' loop and result variable for api calls.
   Dim X1            As Long       ' the x coordinate of the start of the checkmark.
   Dim Y1            As Long       ' the y coordinate of the start of the checkmark vertical line.
   Dim X2            As Long       ' the y coordinate of the end of the checkmark vertical line.
   Dim DrawDirection As Long       ' draw from left to right or right to left?
   Dim ArrowColor    As Long       ' up color or down color?

   If Not m_RightToLeft Then
      X1 = ScaleWidth - ScrollBarButtonWidth - m_BorderWidth + 3
   Else
      X1 = m_BorderWidth + 3
   End If

'  determine x coordinate of first part of check arrow to draw and the direction to draw.
   If WhichButton = UPBUTTON Then
      Y1 = m_BorderWidth + ScrollBarButtonHeight - 6
      DrawDirection = -1
   Else
      Y1 = ScaleHeight - m_BorderWidth - ScrollBarButtonHeight + 6
      DrawDirection = 1
   End If

'  select the correct button color.
   If WhichButton = UPBUTTON Then
      If MouseAction = MOUSE_DOWNED_IN_UPBUTTON Then
         ArrowColor = m_ActiveArrowDownColor
      Else
         ArrowColor = m_ActiveArrowUpColor
      End If
   Else
      If MouseAction = MOUSE_DOWNED_IN_DOWNBUTTON Then
         ArrowColor = m_ActiveArrowDownColor
      Else
         ArrowColor = m_ActiveArrowUpColor
      End If
   End If

'  draw the arrow.
   X2 = 9
   hPN = CreatePen(0, 1, ArrowColor)
   hPO = SelectObject(hdc, hPN)
   MoveTo hdc, X1, Y1, ByVal 0&
   For r = 1 To 5
      LineTo hdc, X1 + X2, Y1
      X1 = X1 + 1
      Y1 = Y1 + DrawDirection
      X2 = X2 - 2
      MoveTo hdc, X1, Y1, ByVal 0&
   Next r

'  delete the pen object.
   r = SelectObject(hdc, hPO)
   r = DeleteObject(hPN)

End Sub

Private Function CalculateScrollTrackHeight() As Long

'**************************************************************************
'* calculates the vertical scrollbar's track height (the distance in      *
'* pixels between the bottom of the top arrow button and the top of the   *
'* bottom arrow button).  Borders are accounted for.                      *
'**************************************************************************

  CalculateScrollTrackHeight = ScaleHeight - (2 * ScrollBarButtonHeight) - (2 * m_BorderWidth)

End Function

Private Function CalculateVScrollThumbHeight() As Long

'*************************************************************************
'* returns the proper height of the vertical scrollbar thumb (in pixels) *
'* based on the number of items in the listbox, the number displayable   *
'* at one time, and the minimum allowable height of the thumb.  Makes    *
'* sure thumb is an odd number of pixels in height, so that the middle   *
'* of the thumb doesn't fall between two pixels.                         *
'*************************************************************************

   Dim VisiblePercentage As Single    ' percentage of list that can fit in display area.
   Dim THeight           As Long      ' preliminary thumb height.

'  no vertical scrollbar needed if the entire list can fit in the display area.  Just a safety net.
   If m_ListCount <= MaxDisplayItems Then
      CalculateVScrollThumbHeight = -1
      Exit Function
   End If

'  calculate the percentage of the list that can be displayed at one time.
   VisiblePercentage = MaxDisplayItems / m_ListCount

'  calculate the corresponding VisiblePercentage of the vertical scrollbar track height.
   THeight = Int(vScrollTrackHeight * VisiblePercentage)

'  if the thumb height is under the defined minimum, change it to the minimum.
   If THeight < vScrollMinThumbHeight Then
      CalculateVScrollThumbHeight = vScrollMinThumbHeight
   Else
'     make sure the height is an odd number of pixels so middle of thumb isn't between two pixels.
      If THeight Mod 2 = 0 Then
         THeight = THeight - 1
      End If
      CalculateVScrollThumbHeight = THeight
   End If

'  reset the flag so scrollbar thumb hieght doesn't get calculated unnecessarily.
   RecalculateThumbHeight = False

End Function

Private Sub ProcessVThumbScroll()

'*************************************************************************
'* allows the user to scroll through the list by dragging the vertical   *
'* scrollbar thumb up or down.                                           *
'*************************************************************************

   Dim ThumbRange     As Long        ' number of pixels middle of thumb can move in scroll track.
   Dim YPosPct        As Single      ' the percentage of the total list to move.
   Dim MouseMovement  As Long        ' how many pixels thumb was dragged from original position.
   Dim NumItemsToMove As Long        ' number of items to move list display by.

   ThumbScrolling = True
   MousePosInVThumb = MouseY - vScrollBarLocation.ScrollThumbLocation.Top

   While MouseAction <> MOUSE_NOACTION And DraggingVThumb

'     detect MouseMove for thumb positioning, and MouseUp to change value of MouseAction when it happens.
      DoEvents
'     keeps cpu usage from maxing out.  Thanks to Mike Douglas for the tip.
      Sleep 25

'     if mouse has moved below permissible scrolling range, display bottom of list and exit.
      If MouseY > ScaleHeight - m_BorderWidth - ScrollBarButtonHeight Then
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
         DisplayList
         ThumbScrolling = False ' used in MouseMove to see if we can resume thumb scrolling.
         Exit Sub
      End If

'     if mouse has moved above permissible scrolling range, display top of list and exit.
      If MouseY <= m_BorderWidth + ScrollBarButtonHeight Then
         DisplayRange.FirstListItem = 0
         DisplayRange.LastListItem = MaxDisplayItems - 1
         DisplayList
         ThumbScrolling = False ' used in MouseMove to see if we can resume thumb scrolling.
         Exit Sub
      End If

'     calculate how far the mouse has moved from the original mousedown y position.
      MouseMovement = MouseY - MouseDownYPos ' could be pos. or neg., depends on direction mouse moves.

'     determine the thumb's middle-pixel movement range.
      ThumbRange = vThumbRange.Bottom - vThumbRange.Top + 1 '# of pixels in the range.
'     determine how far, percentagewise, the list should be moved.
      YPosPct = (Abs(MouseMovement) / ThumbRange)
'     how many items is that?
      NumItemsToMove = Int(m_ListCount * YPosPct)

'     change it to negative if necessary.
      If MouseMovement < 0 Then
         NumItemsToMove = 0 - NumItemsToMove
      End If

'     only do the scroll if the range to display has changed.
      If NumItemsToMove <> 0 Then
         DisplayRange.FirstListItem = DisplayRange.FirstListItem + NumItemsToMove
         If DisplayRange.FirstListItem < 0 Then
            DisplayRange.FirstListItem = 0
         End If
         If DisplayRange.FirstListItem + MaxDisplayItems - 1 > m_ListCount Then
            DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
         End If
         DisplayRange.LastListItem = DisplayRange.FirstListItem + MaxDisplayItems - 1
         DisplayList MousePosInVThumb
         MousePosInVThumb = MouseY - vScrollBarLocation.ScrollThumbLocation.Top
         MouseDownYPos = MouseY
      End If

   Wend

End Sub

Private Sub ProcessMouseDragThumbOutOfRange(DidIt As Boolean)

'*************************************************************************
'* this code accounts for when mouse drags thumb above or below maximum  *
'* vertical scroll range.  The top or bottom of the list is then         *
'* automatically displayed.                                              *
'*************************************************************************

   DidIt = False

   If DraggingVThumb Then
      If MouseY <= m_BorderWidth + ScrollBarButtonHeight Then
         DisplayRange.FirstListItem = 0
         DisplayRange.LastListItem = MaxDisplayItems - 1
         DisplayList
         DidIt = True
      ElseIf MouseY >= vScrollBarLocation.DownButtonLocation.Top Then
         DisplayRange.FirstListItem = m_ListCount - MaxDisplayItems + 1
         DisplayRange.LastListItem = m_ListCount - 1
         DisplayList
         DidIt = True
      End If
   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to the default constants.                       *
'*************************************************************************

   Set m_Picture = LoadPicture("")
   Set m_DisPicture = LoadPicture("")
   Set m_ListFont = Ambient.Font
   m_ArrowUpColor = m_def_ArrowUpColor
   m_ArrowDownColor = m_def_ArrowDownColor
   m_AutoRedraw = UserControl.AutoRedraw
   m_BackAngle = m_def_BackAngle
   m_BackColor2 = m_def_BackColor2
   m_BackColor1 = m_def_BackColor1
   m_BackMiddleOut = m_def_BackMiddleOut
   m_BorderColor1 = m_def_BorderColor1
   m_BorderColor2 = m_def_BorderColor2
   m_BorderMiddleOut = m_def_BorderMiddleOut
   m_BorderWidth = m_def_BorderWidth
   m_ButtonColor1 = m_def_ButtonColor1
   m_ButtonColor2 = m_def_ButtonColor2
   m_CheckBoxArrowColor = m_def_CheckBoxArrowColor
   m_CheckBoxColor = m_def_CheckBoxColor
   m_CheckStyle = m_def_CheckStyle
   m_CircularGradient = m_def_CircularGradient
   m_CurveTopLeft = m_def_CurveTopLeft
   m_CurveTopRight = m_def_CurveTopRight
   m_CurveBottomLeft = m_def_CurveBottomLeft
   m_CurveBottomRight = m_def_CurveBottomRight
   m_DblClickBehavior = m_def_DblClickBehavior
   m_DisArrowDownColor = m_def_DisArrowDownColor
   m_DisArrowUpColor = m_def_DisArrowUpColor
   m_DisBackColor1 = m_def_DisBackColor1
   m_DisBackColor2 = m_def_DisBackColor2
   m_DisBorderColor1 = m_def_DisBorderColor1
   m_DisBorderColor2 = m_def_DisBorderColor2
   m_DisButtonColor1 = m_def_DisButtonColor1
   m_DisButtonColor2 = m_def_DisButtonColor2
   m_DisCheckboxArrowColor = m_def_DisCheckboxArrowColor
   m_DisCheckboxColor = m_def_DisCheckboxColor
   m_DisFocusRectColor = m_def_DisFocusRectColor
   m_DisPictureMode = m_def_DisPictureMode
   m_DisSelColor1 = m_def_DisSelColor1
   m_DisSelColor2 = m_def_DisSelColor2
   m_DisSelTextColor = m_def_DisSelTextColor
   m_DisTextColor = m_def_DisTextColor
   m_DisThumbBorderColor = m_def_DisThumbBorderColor
   m_DisThumbColor1 = m_def_DisThumbColor1
   m_DisThumbColor2 = m_def_DisThumbColor2
   m_DisTrackbarColor1 = m_def_DisTrackbarColor1
   m_DisTrackbarColor2 = m_def_DisTrackbarColor2
   m_DragEnabled = m_def_DragEnabled
   m_Enabled = m_def_Enabled
   m_FocusBorderColor1 = m_def_FocusBorderColor1
   m_FocusBorderColor2 = m_def_FocusBorderColor2
   m_FocusRectColor = m_def_FocusRectColor
   m_ItemImageSize = m_def_ItemImageSize
   m_ListIndex = m_def_ListIndex
   m_MultiSelect = m_def_MultiSelect
   m_NewIndex = m_def_NewIndex
   m_PictureMode = m_def_PictureMode
   m_RedrawFlag = m_def_RedrawFlag
   m_RightToLeft = m_def_RightToLeft
   m_ScaleHeight = UserControl.ScaleHeight
   m_ScaleMode = UserControl.ScaleMode
   m_ScaleWidth = UserControl.ScaleWidth
   m_SelColor1 = m_def_SelColor1
   m_SelColor2 = m_def_SelColor2
   m_SelCount = m_def_SelCount
   m_SelTextColor = m_def_SelTextColor
   m_ShowItemImages = m_def_ShowItemImages
   m_ShowSelectRect = m_def_ShowSelectRect
   m_SortAsNumeric = m_def_SortAsNumeric
   m_Sorted = m_def_Sorted
   m_Style = m_def_Style
   m_Text = m_def_Text
   m_TextColor = m_def_TextColor
   m_Theme = m_def_Theme
   m_ThumbBorderColor = m_def_ThumbBorderColor
   m_ThumbColor1 = m_def_ThumbColor1
   m_ThumbColor2 = m_def_ThumbColor2
   m_TopIndex = m_def_TopIndex
   m_TrackBarColor1 = m_def_TrackBarColor1
   m_TrackBarColor2 = m_def_TrackBarColor2
   m_TrackClickColor1 = m_def_TrackClickColor1
   m_TrackClickColor2 = m_def_TrackClickColor2

'  initialize appropriate display colors.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read properties in the property bag.                                  *
'*************************************************************************

   With PropBag
      Set m_ListFont = .ReadProperty("ListFont", Ambient.Font)
      Set UserControl.Font = m_ListFont
      Set m_Picture = .ReadProperty("Picture", Nothing)
      Set m_DisPicture = PropBag.ReadProperty("DisPicture", Nothing)
      m_ArrowDownColor = .ReadProperty("ArrowDownColor", m_def_ArrowDownColor)
      m_ArrowUpColor = .ReadProperty("ArrowUpColor", m_def_ArrowUpColor)
      m_AutoRedraw = UserControl.AutoRedraw
      m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
      m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
      m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
      m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
      m_BorderColor1 = .ReadProperty("BorderColor1", m_def_BorderColor1)
      m_BorderColor2 = .ReadProperty("BorderColor2", m_def_BorderColor2)
      m_BorderMiddleOut = .ReadProperty("BorderMiddleOut", m_def_BorderMiddleOut)
      m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
      m_ButtonColor1 = .ReadProperty("ButtonColor1", m_def_ButtonColor1)
      m_ButtonColor2 = .ReadProperty("ButtonColor2", m_def_ButtonColor2)
      m_CheckBoxArrowColor = .ReadProperty("CheckboxArrowColor", m_def_CheckBoxArrowColor)
      m_CheckBoxColor = .ReadProperty("CheckBoxColor", m_def_CheckBoxColor)
      m_CheckStyle = .ReadProperty("CheckStyle", m_def_CheckStyle)
      m_CircularGradient = .ReadProperty("CircularGradient", m_def_CircularGradient)
      m_CurveTopLeft = .ReadProperty("CurveTopLeft", m_def_CurveTopLeft)
      m_CurveTopRight = .ReadProperty("CurveTopRight", m_def_CurveTopRight)
      m_CurveBottomLeft = .ReadProperty("CurveBottomLeft", m_def_CurveBottomLeft)
      m_CurveBottomRight = .ReadProperty("CurveBottomRight", m_def_CurveBottomRight)
      m_DblClickBehavior = .ReadProperty("DblClickBehavior", m_def_DblClickBehavior)
      m_DisArrowDownColor = .ReadProperty("DisArrowDownColor", m_def_DisArrowDownColor)
      m_DisArrowUpColor = .ReadProperty("DisArrowUpColor", m_def_DisArrowUpColor)
      m_DisBackColor1 = .ReadProperty("DisBackColor1", m_def_DisBackColor1)
      m_DisBackColor2 = .ReadProperty("DisBackColor2", m_def_DisBackColor2)
      m_DisBorderColor1 = .ReadProperty("DisBorderColor1", m_def_DisBorderColor1)
      m_DisBorderColor2 = .ReadProperty("DisBorderColor2", m_def_DisBorderColor2)
      m_DisButtonColor1 = .ReadProperty("DisButtonColor1", m_def_DisButtonColor1)
      m_DisButtonColor2 = .ReadProperty("DisButtonColor2", m_def_DisButtonColor2)
      m_DisCheckboxArrowColor = .ReadProperty("DisCheckboxArrowColor", m_def_DisCheckboxArrowColor)
      m_DisCheckboxColor = .ReadProperty("DisCheckboxColor", m_def_DisCheckboxColor)
      m_DisFocusRectColor = .ReadProperty("DisFocusRectColor", m_def_DisFocusRectColor)
      m_DisPictureMode = .ReadProperty("DisPictureMode", m_def_DisPictureMode)
      m_DisSelColor1 = .ReadProperty("DisSelColor1", m_def_DisSelColor1)
      m_DisSelColor2 = .ReadProperty("DisSelColor2", m_def_DisSelColor2)
      m_DisSelTextColor = .ReadProperty("DisSelTextColor", m_def_DisSelTextColor)
      m_DisTextColor = .ReadProperty("DisTextColor", m_def_DisTextColor)
      m_DisThumbBorderColor = .ReadProperty("DisThumbBorderColor", m_def_DisThumbBorderColor)
      m_DisThumbColor1 = .ReadProperty("DisThumbColor1", m_def_DisThumbColor1)
      m_DisThumbColor2 = .ReadProperty("DisThumbColor2", m_def_DisThumbColor2)
      m_DisTrackbarColor1 = .ReadProperty("DisTrackbarColor1", m_def_DisTrackbarColor1)
      m_DisTrackbarColor2 = .ReadProperty("DisTrackbarColor2", m_def_DisTrackbarColor2)
      m_DragEnabled = .ReadProperty("DragEnabled", m_def_DragEnabled)
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      m_FocusBorderColor1 = .ReadProperty("FocusBorderColor1", m_def_FocusBorderColor1)
      m_FocusBorderColor2 = .ReadProperty("FocusBorderColor2", m_def_FocusBorderColor2)
      m_FocusRectColor = .ReadProperty("FocusRectColor", m_def_FocusRectColor)
      m_ItemImageSize = .ReadProperty("ItemImageSize", m_def_ItemImageSize)
      m_MultiSelect = .ReadProperty("MultiSelect", m_def_MultiSelect)
      m_NewIndex = .ReadProperty("NewIndex", m_def_NewIndex)
      m_PictureMode = .ReadProperty("PictureMode", m_def_PictureMode)
      m_RedrawFlag = .ReadProperty("RedrawFlag", m_def_RedrawFlag)
      m_RightToLeft = .ReadProperty("RightToLeft", m_def_RightToLeft)
      m_ScaleHeight = UserControl.ScaleHeight
      m_ScaleMode = UserControl.ScaleMode
      m_ScaleWidth = UserControl.ScaleWidth
      m_SelColor1 = .ReadProperty("SelColor1", m_def_SelColor1)
      m_SelColor2 = .ReadProperty("SelColor2", m_def_SelColor2)
      m_SelCount = .ReadProperty("SelCount", m_def_SelCount)
      m_SelTextColor = .ReadProperty("SelTextColor", m_def_SelTextColor)
      m_ShowItemImages = .ReadProperty("ShowItemImages", m_def_ShowItemImages)
      m_ShowSelectRect = .ReadProperty("ShowSelectRect", m_def_ShowSelectRect)
      m_SortAsNumeric = .ReadProperty("SortAsNumeric", m_def_SortAsNumeric)
      m_Sorted = .ReadProperty("Sorted", m_def_Sorted)
      m_Style = .ReadProperty("Style", m_def_Style)
      m_Text = .ReadProperty("Text", m_def_Text)
      m_TextColor = .ReadProperty("TextColor", m_def_TextColor)
      m_Theme = .ReadProperty("Theme", m_def_Theme)
      m_ThumbBorderColor = .ReadProperty("ThumbBorderColor", m_def_ThumbBorderColor)
      m_ThumbColor1 = .ReadProperty("ThumbColor1", m_def_ThumbColor1)
      m_ThumbColor2 = .ReadProperty("ThumbColor2", m_def_ThumbColor2)
      m_TopIndex = .ReadProperty("TopIndex", m_def_TopIndex)
      m_TrackBarColor1 = .ReadProperty("TrackBarColor1", m_def_TrackBarColor1)
      m_TrackBarColor2 = .ReadProperty("TrackBarColor2", m_def_TrackBarColor2)
      m_TrackClickColor1 = .ReadProperty("TrackClickColor1", m_def_TrackClickColor1)
      m_TrackClickColor2 = .ReadProperty("TrackClickColor2", m_def_TrackClickColor2)
   End With

   DragFlag = m_DragEnabled
   If m_Sorted Then
      ListIsSorted = True
   End If

'  initially, the .ListIndex property is determined by the value of the .MultiSelect property.
   m_ListIndex = IIf(m_MultiSelect = vbMultiSelectNone, -1, 0)

'  initialize appropriate display colors.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

'  initialize the list item display range structure.
   DisplayRange.FirstListItem = -1
   DisplayRange.LastListItem = -1

'  initialize gradients, list item height, display coordinates.
   InitListBoxDisplayCharacteristics

'  start up the subclassing.
   StartSubclassing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write the properties in the property bag.                             *
'*************************************************************************

   With PropBag
      .WriteProperty "ArrowUpColor", m_ArrowUpColor, m_def_ArrowUpColor
      .WriteProperty "ArrowDownColor", m_ArrowDownColor, m_def_ArrowDownColor
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "BorderColor1", m_BorderColor1, m_def_BorderColor1
      .WriteProperty "BorderColor2", m_BorderColor2, m_def_BorderColor2
      .WriteProperty "BorderMiddleOut", m_BorderMiddleOut, m_def_BorderMiddleOut
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "ButtonColor1", m_ButtonColor1, m_def_ButtonColor1
      .WriteProperty "ButtonColor2", m_ButtonColor2, m_def_ButtonColor2
      .WriteProperty "CheckboxArrowColor", m_CheckBoxArrowColor, m_def_CheckBoxArrowColor
      .WriteProperty "CheckBoxColor", m_CheckBoxColor, m_def_CheckBoxColor
      .WriteProperty "CheckStyle", m_CheckStyle, m_def_CheckStyle
      .WriteProperty "CircularGradient", m_CircularGradient, m_def_CircularGradient
      .WriteProperty "CurveTopLeft", m_CurveTopLeft, m_def_CurveTopLeft
      .WriteProperty "CurveTopRight", m_CurveTopRight, m_def_CurveTopRight
      .WriteProperty "CurveBottomLeft", m_CurveBottomLeft, m_def_CurveBottomLeft
      .WriteProperty "CurveBottomRight", m_CurveBottomRight, m_def_CurveBottomRight
      .WriteProperty "DblClickBehavior", m_DblClickBehavior, m_def_DblClickBehavior
      .WriteProperty "DisArrowDownColor", m_DisArrowDownColor, m_def_DisArrowDownColor
      .WriteProperty "DisArrowUpColor", m_DisArrowUpColor, m_def_DisArrowUpColor
      .WriteProperty "DisBackColor1", m_DisBackColor1, m_def_DisBackColor1
      .WriteProperty "DisBackColor2", m_DisBackColor2, m_def_DisBackColor2
      .WriteProperty "DisBorderColor1", m_DisBorderColor1, m_def_DisBorderColor1
      .WriteProperty "DisBorderColor2", m_DisBorderColor2, m_def_DisBorderColor2
      .WriteProperty "DisButtonColor1", m_DisButtonColor1, m_def_DisButtonColor1
      .WriteProperty "DisButtonColor2", m_DisButtonColor2, m_def_DisButtonColor2
      .WriteProperty "DisCheckboxArrowColor", m_DisCheckboxArrowColor, m_def_DisCheckboxArrowColor
      .WriteProperty "DisCheckboxColor", m_DisCheckboxColor, m_def_DisCheckboxColor
      .WriteProperty "DisFocusRectColor", m_DisFocusRectColor, m_def_DisFocusRectColor
      .WriteProperty "DisPicture", m_DisPicture, Nothing
      .WriteProperty "DisPictureMode", m_DisPictureMode, m_def_DisPictureMode
      .WriteProperty "DisSelColor1", m_DisSelColor1, m_def_DisSelColor1
      .WriteProperty "DisSelColor2", m_DisSelColor2, m_def_DisSelColor2
      .WriteProperty "DisSelTextColor", m_DisSelTextColor, m_def_DisSelTextColor
      .WriteProperty "DisTextColor", m_DisTextColor, m_def_DisTextColor
      .WriteProperty "DisThumbBorderColor", m_DisThumbBorderColor, m_def_DisThumbBorderColor
      .WriteProperty "DisThumbColor1", m_DisThumbColor1, m_def_DisThumbColor1
      .WriteProperty "DisThumbColor2", m_DisThumbColor2, m_def_DisThumbColor2
      .WriteProperty "DisTrackbarColor1", m_DisTrackbarColor1, m_def_DisTrackbarColor1
      .WriteProperty "DisTrackbarColor2", m_DisTrackbarColor2, m_def_DisTrackbarColor2
      .WriteProperty "DragEnabled", m_DragEnabled, m_def_DragEnabled
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "FocusBorderColor1", m_FocusBorderColor1, m_def_FocusBorderColor1
      .WriteProperty "FocusBorderColor2", m_FocusBorderColor2, m_def_FocusBorderColor2
      .WriteProperty "FocusRectColor", m_FocusRectColor, m_def_FocusRectColor
      .WriteProperty "ItemImageSize", m_ItemImageSize, m_def_ItemImageSize
      .WriteProperty "ListIndex", m_ListIndex, m_def_ListIndex
      .WriteProperty "ListFont", m_ListFont, Ambient.Font
      .WriteProperty "MultiSelect", m_MultiSelect, m_def_MultiSelect
      .WriteProperty "NewIndex", m_NewIndex, m_def_NewIndex
      .WriteProperty "RedrawFlag", m_RedrawFlag, m_def_RedrawFlag
      .WriteProperty "RightToLeft", m_RightToLeft, m_def_RightToLeft
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "PictureMode", m_PictureMode, m_def_PictureMode
      .WriteProperty "SelColor1", m_SelColor1, m_def_SelColor1
      .WriteProperty "SelColor2", m_SelColor2, m_def_SelColor2
      .WriteProperty "SelCount", m_SelCount, m_def_SelCount
      .WriteProperty "SelTextColor", m_SelTextColor, m_def_SelTextColor
      .WriteProperty "ShowItemImages", m_ShowItemImages, m_def_ShowItemImages
      .WriteProperty "ShowSelectRect", m_ShowSelectRect, m_def_ShowSelectRect
      .WriteProperty "SortAsNumeric", m_SortAsNumeric, m_def_SortAsNumeric
      .WriteProperty "Sorted", m_Sorted, m_def_Sorted
      .WriteProperty "Style", m_Style, m_def_Style
      .WriteProperty "Text", m_Text, m_def_Text
      .WriteProperty "TextColor", m_TextColor, m_def_TextColor
      .WriteProperty "Theme", m_Theme, m_def_Theme
      .WriteProperty "ThumbBorderColor", m_ThumbBorderColor, m_def_ThumbBorderColor
      .WriteProperty "ThumbColor1", m_ThumbColor1, m_def_ThumbColor1
      .WriteProperty "ThumbColor2", m_ThumbColor2, m_def_ThumbColor2
      .WriteProperty "TopIndex", m_TopIndex, m_def_TopIndex
      .WriteProperty "TrackBarColor1", m_TrackBarColor1, m_def_TrackBarColor1
      .WriteProperty "TrackBarColor2", m_TrackBarColor2, m_def_TrackBarColor2
      .WriteProperty "TrackClickColor1", m_TrackClickColor1, m_def_TrackClickColor1
      .WriteProperty "TrackClickColor2", m_TrackClickColor2, m_def_TrackClickColor2
   End With

End Sub

Public Property Get ArrowDownColor() As OLE_COLOR
Attribute ArrowDownColor.VB_Description = "The color of the scrollbar up/down arrow symbols when scrollbar buttons are clicked down."
   ArrowDownColor = m_ArrowDownColor
End Property

Public Property Let ArrowDownColor(ByVal New_ArrowDownColor As OLE_COLOR)
   m_ArrowDownColor = New_ArrowDownColor
   PropertyChanged "ArrowDownColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ArrowUpColor() As OLE_COLOR
Attribute ArrowUpColor.VB_Description = "The color of the scrollbar up/down arrow symbols when scrollbar buttons are not clicked down (default state)."
   ArrowUpColor = m_ArrowUpColor
End Property

Public Property Let ArrowUpColor(ByVal New_ArrowUpColor As OLE_COLOR)
   m_ArrowUpColor = New_ArrowUpColor
   PropertyChanged "ArrowUpColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get AutoRedraw() As Boolean
   AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
   m_AutoRedraw = New_AutoRedraw
   UserControl.AutoRedraw = m_AutoRedraw
   PropertyChanged "AutoRedraw"
End Property

Public Property Get BackAngle() As Single
Attribute BackAngle.VB_Description = "The angle of the listbox background gradient."
Attribute BackAngle.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
'  do some bounds checking.
   If New_BackAngle > 360 Then
      New_BackAngle = 360
   ElseIf New_BackAngle < 0 Then
      New_BackAngle = 0
   End If
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first color of the listbox background gradient."
Attribute BackColor1.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second color of the listbox background gradient."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BackMiddleOut() As Boolean
Attribute BackMiddleOut.VB_Description = "Allows the background gradient to be middle-out (Color 1> Color 2 > Color 1)."
Attribute BackMiddleOut.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   CalculateBackGroundGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BorderColor1() As OLE_COLOR
Attribute BorderColor1.VB_Description = "The first gradient color of the MorphListBox border."
Attribute BorderColor1.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderColor1 = m_BorderColor1
End Property

Public Property Let BorderColor1(ByVal New_BorderColor1 As OLE_COLOR)
   m_BorderColor1 = New_BorderColor1
   m_ActiveBorderColor1 = m_BorderColor1
   PropertyChanged "BorderColor1"
   InitBorder
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BorderColor2() As OLE_COLOR
Attribute BorderColor2.VB_Description = "The second gradient color of the MorphListBox border."
   BorderColor2 = m_BorderColor2
End Property

Public Property Let BorderColor2(ByVal New_BorderColor2 As OLE_COLOR)
   m_BorderColor2 = New_BorderColor2
   m_ActiveBorderColor2 = m_BorderColor2
   PropertyChanged "BorderColor2"
   InitBorder
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get BorderMiddleOut() As Boolean
Attribute BorderMiddleOut.VB_Description = "Allows the border gradient to be middle-out (Color 1> Color 2 > Color 1)."
   BorderMiddleOut = m_BorderMiddleOut
End Property

Public Property Let BorderMiddleOut(ByVal New_BorderMiddleOut As Boolean)
   m_BorderMiddleOut = New_BorderMiddleOut
   PropertyChanged "BorderMiddleOut"
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "The width, in pixels, of the MorphListBox border."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
   InitBorder                            ' generate new gradient borders.
   CalculateVerticalTrackbarGradients    ' generate new vertical scrollbar trackbar.
   InitTextDisplayCharacteristics        ' recalculate text display boundaries to accommodate new borders.
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ButtonColor1() As OLE_COLOR
Attribute ButtonColor1.VB_Description = "The first gradient color of the scrollbar up/down buttons."
   ButtonColor1 = m_ButtonColor1
End Property

Public Property Let ButtonColor1(ByVal New_ButtonColor1 As OLE_COLOR)
   m_ButtonColor1 = New_ButtonColor1
   PropertyChanged "ButtonColor1"
   CalculateScrollbarButtonGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ButtonColor2() As OLE_COLOR
Attribute ButtonColor2.VB_Description = "The second gradient color of the scrollbar up/down buttons."
   ButtonColor2 = m_ButtonColor2
End Property

Public Property Let ButtonColor2(ByVal New_ButtonColor2 As OLE_COLOR)
   m_ButtonColor2 = New_ButtonColor2
   PropertyChanged "ButtonColor2"
   CalculateScrollbarButtonGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CheckboxArrowColor() As OLE_COLOR
Attribute CheckboxArrowColor.VB_Description = "The color of the checkbox arrow, tick or 'X' when .Style property is set to Checkbox."
   CheckboxArrowColor = m_CheckBoxArrowColor
End Property

Public Property Let CheckboxArrowColor(ByVal New_CheckboxArrowColor As OLE_COLOR)
   m_CheckBoxArrowColor = New_CheckboxArrowColor
   PropertyChanged "CheckboxArrowColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CheckBoxColor() As OLE_COLOR
Attribute CheckBoxColor.VB_Description = "The color of the checkbox border when the .Style property is set to Checkbox."
   CheckBoxColor = m_CheckBoxColor
End Property

Public Property Let CheckBoxColor(ByVal New_CheckBoxColor As OLE_COLOR)
   m_CheckBoxColor = New_CheckBoxColor
   PropertyChanged "CheckBoxColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CheckStyle() As CheckStyleOptions
Attribute CheckStyle.VB_Description = "The style of the listitem selected indicator in Checkbox mode: Tick, Check or Arrow."
   CheckStyle = m_CheckStyle
End Property

Public Property Let CheckStyle(ByVal New_CheckStyle As CheckStyleOptions)
   m_CheckStyle = New_CheckStyle
   RedrawControl
   UserControl.Refresh
   PropertyChanged "CheckStyle"
End Property

Public Property Get CircularGradient() As Boolean
Attribute CircularGradient.VB_Description = "If True, control background gradient is displayed in circular (radiant) fashion as opposed to linear."
   CircularGradient = m_CircularGradient
End Property

Public Property Let CircularGradient(ByVal New_CircularGradient As Boolean)
   m_CircularGradient = New_CircularGradient
   CalculateBackGroundGradient
   RedrawControl
   PropertyChanged "CircularGradient"
End Property

Public Property Get CurveBottomLeft() As Long
Attribute CurveBottomLeft.VB_Description = "The amount of curve of the bottom left corner of the ListBox. Only valid if BorderWidth property is <= 2 (pixels)."
Attribute CurveBottomLeft.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveBottomLeft = m_CurveBottomLeft
End Property

Public Property Let CurveBottomLeft(ByVal New_CurveBottomLeft As Long)
   m_CurveBottomLeft = New_CurveBottomLeft
   PropertyChanged "CurveBottomLeft"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveBottomRight() As Long
Attribute CurveBottomRight.VB_Description = "The amount of curve of the bottom right corner of the ListBox. Only valid if BorderWidth property is <= 2 (pixels)."
Attribute CurveBottomRight.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveBottomRight = m_CurveBottomRight
End Property

Public Property Let CurveBottomRight(ByVal New_CurveBottomRight As Long)
   m_CurveBottomRight = New_CurveBottomRight
   PropertyChanged "CurveBottomRight"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveTopLeft() As Long
Attribute CurveTopLeft.VB_Description = "The amount of curve of the top left corner of the ListBox. Only valid if BorderWidth property is <= 2 (pixels)."
Attribute CurveTopLeft.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveTopLeft = m_CurveTopLeft
End Property

Public Property Let CurveTopLeft(ByVal New_CurveTopLeft As Long)
   m_CurveTopLeft = New_CurveTopLeft
   PropertyChanged "CurveTopLeft"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get CurveTopRight() As Long
Attribute CurveTopRight.VB_Description = "The amount of curve of the top right corner of the ListBox. Only valid if BorderWidth property is <= 2 (pixels)."
Attribute CurveTopRight.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveTopRight = m_CurveTopRight
End Property

Public Property Let CurveTopRight(ByVal New_CurveTopRight As Long)
   m_CurveTopRight = New_CurveTopRight
   PropertyChanged "CurveTopRight"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get DblClickBehavior() As DblClickBehaviorOptions
Attribute DblClickBehavior.VB_Description = "Allows user to determine mouse left-double-click behavior: Either return a DblClick event or two rapid single click events."
   DblClickBehavior = m_DblClickBehavior
End Property

Public Property Let DblClickBehavior(ByVal New_DblClickBehavior As DblClickBehaviorOptions)
'  this makes the property read-only at runtime.
   If Ambient.UserMode Then Err.Raise 382
   m_DblClickBehavior = New_DblClickBehavior
   PropertyChanged "DblClickBehavior"
End Property

Public Property Get DisArrowDownColor() As OLE_COLOR
Attribute DisArrowDownColor.VB_Description = "The color of the scrollbar button directional arrow symbols when the control is disabled."
Attribute DisArrowDownColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisArrowDownColor = m_DisArrowDownColor
End Property

Public Property Let DisArrowDownColor(ByVal New_DisArrowDownColor As OLE_COLOR)
   m_DisArrowDownColor = New_DisArrowDownColor
   PropertyChanged "DisArrowDownColor"
End Property

Public Property Get DisArrowUpColor() As OLE_COLOR
Attribute DisArrowUpColor.VB_Description = "The color of the scrollbar button directional arrow symbols when the control is disabled."
Attribute DisArrowUpColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisArrowUpColor = m_DisArrowUpColor
End Property

Public Property Let DisArrowUpColor(ByVal New_DisArrowUpColor As OLE_COLOR)
   m_DisArrowUpColor = New_DisArrowUpColor
   PropertyChanged "DisArrowUpColor"
End Property

Public Property Get DisBackColor1() As OLE_COLOR
Attribute DisBackColor1.VB_Description = "The first color of the background gradient when the control is disabled."
Attribute DisBackColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisBackColor1 = m_DisBackColor1
End Property

Public Property Let DisBackColor1(ByVal New_DisBackColor1 As OLE_COLOR)
   m_DisBackColor1 = New_DisBackColor1
   PropertyChanged "DisBackColor1"
End Property

Public Property Get DisBackColor2() As OLE_COLOR
Attribute DisBackColor2.VB_Description = "The second color of the background gradient when the control is disabled."
Attribute DisBackColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisBackColor2 = m_DisBackColor2
End Property

Public Property Let DisBackColor2(ByVal New_DisBackColor2 As OLE_COLOR)
   m_DisBackColor2 = New_DisBackColor2
   PropertyChanged "DisBackColor2"
End Property

Public Property Get DisBorderColor1() As OLE_COLOR
Attribute DisBorderColor1.VB_Description = "The first color of the border gradient when the control is disabled."
Attribute DisBorderColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisBorderColor1 = m_DisBorderColor1
End Property

Public Property Let DisBorderColor1(ByVal New_DisBorderColor1 As OLE_COLOR)
   m_DisBorderColor1 = New_DisBorderColor1
   PropertyChanged "DisBorderColor1"
End Property

Public Property Get DisBorderColor2() As OLE_COLOR
Attribute DisBorderColor2.VB_Description = "The second color of the border gradient when the control is disabled."
   DisBorderColor2 = m_DisBorderColor2
End Property

Public Property Let DisBorderColor2(ByVal New_DisBorderColor2 As OLE_COLOR)
   m_DisBorderColor2 = New_DisBorderColor2
   PropertyChanged "DisBorderColor2"
End Property

Public Property Get DisButtonColor1() As OLE_COLOR
Attribute DisButtonColor1.VB_Description = "The first color of the scrollbar buttons gradient when the control is disabled."
Attribute DisButtonColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisButtonColor1 = m_DisButtonColor1
End Property

Public Property Let DisButtonColor1(ByVal New_DisButtonColor1 As OLE_COLOR)
   m_DisButtonColor1 = New_DisButtonColor1
   PropertyChanged "DisButtonColor1"
End Property

Public Property Get DisButtonColor2() As OLE_COLOR
Attribute DisButtonColor2.VB_Description = "The second color of the scrollbar buttons gradient when the control is disabled."
Attribute DisButtonColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisButtonColor2 = m_DisButtonColor2
End Property

Public Property Let DisButtonColor2(ByVal New_DisButtonColor2 As OLE_COLOR)
   m_DisButtonColor2 = New_DisButtonColor2
   PropertyChanged "DisButtonColor2"
End Property

Public Property Get DisCheckboxArrowColor() As OLE_COLOR
Attribute DisCheckboxArrowColor.VB_Description = "The color of .Checkbox mode listitem selected checkbox arrow when control is disabled."
Attribute DisCheckboxArrowColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisCheckboxArrowColor = m_DisCheckboxArrowColor
End Property

Public Property Let DisCheckboxArrowColor(ByVal New_DisCheckboxArrowColor As OLE_COLOR)
   m_DisCheckboxArrowColor = New_DisCheckboxArrowColor
   PropertyChanged "DisCheckboxArrowColor"
End Property

Public Property Get DisCheckboxColor() As OLE_COLOR
Attribute DisCheckboxColor.VB_Description = "The color of .Checkbox mode checkbox border when control is disabled."
Attribute DisCheckboxColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisCheckboxColor = m_DisCheckboxColor
End Property

Public Property Let DisCheckboxColor(ByVal New_DisCheckboxColor As OLE_COLOR)
   m_DisCheckboxColor = New_DisCheckboxColor
   PropertyChanged "DisCheckboxColor"
End Property

Public Property Get DisFocusRectColor() As OLE_COLOR
Attribute DisFocusRectColor.VB_Description = "The color of the listitem focus rectangle when control is disabled."
Attribute DisFocusRectColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisFocusRectColor = m_DisFocusRectColor
End Property

Public Property Let DisFocusRectColor(ByVal New_DisFocusRectColor As OLE_COLOR)
   m_DisFocusRectColor = New_DisFocusRectColor
   PropertyChanged "DisFocusRectColor"
End Property

Public Property Get DisPicture() As Picture
Attribute DisPicture.VB_Description = "The background image to use when control is disabled."
Attribute DisPicture.VB_ProcData.VB_Invoke_Property = ";Disabled"
   Set DisPicture = m_DisPicture
End Property

Public Property Set DisPicture(ByVal New_DisPicture As Picture)
   Set m_DisPicture = New_DisPicture
   PropertyChanged "DisPicture"
'  this flag tells Redraw to re-blit the new background to the virtual DC.
   ChangingPicture = True
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get DisPictureMode() As MLB_PictureModeOptions
Attribute DisPictureMode.VB_Description = "The image display mode to use when control is disabled (normal, stretched or tiled)."
Attribute DisPictureMode.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisPictureMode = m_DisPictureMode
End Property

Public Property Let DisPictureMode(ByVal New_DisPictureMode As MLB_PictureModeOptions)
   m_DisPictureMode = New_DisPictureMode
   PropertyChanged "DisPictureMode"
End Property

Public Property Get DisSelColor1() As OLE_COLOR
Attribute DisSelColor1.VB_Description = "The first gradient color of the listitem selection bar when control is disabled."
Attribute DisSelColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisSelColor1 = m_DisSelColor1
End Property

Public Property Let DisSelColor1(ByVal New_DisSelColor1 As OLE_COLOR)
   m_DisSelColor1 = New_DisSelColor1
   PropertyChanged "DisSelColor1"
End Property

Public Property Get DisSelColor2() As OLE_COLOR
Attribute DisSelColor2.VB_Description = "The second gradient color of the listitem selection bar when control is disabled."
Attribute DisSelColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisSelColor2 = m_DisSelColor2
End Property

Public Property Let DisSelColor2(ByVal New_DisSelColor2 As OLE_COLOR)
   m_DisSelColor2 = New_DisSelColor2
   PropertyChanged "DisSelColor2"
End Property

Public Property Get DisSelTextColor() As OLE_COLOR
Attribute DisSelTextColor.VB_Description = "The color of listitem with selection bar when control is disabled."
Attribute DisSelTextColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisSelTextColor = m_DisSelTextColor
End Property

Public Property Let DisSelTextColor(ByVal New_DisSelTextColor As OLE_COLOR)
   m_DisSelTextColor = New_DisSelTextColor
   PropertyChanged "DisSelTextColor"
End Property

Public Property Get DisTextColor() As OLE_COLOR
Attribute DisTextColor.VB_Description = "listitem text color when control is disabled."
Attribute DisTextColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisTextColor = m_DisTextColor
End Property

Public Property Let DisTextColor(ByVal New_DisTextColor As OLE_COLOR)
   m_DisTextColor = New_DisTextColor
   PropertyChanged "DisTextColor"
End Property

Public Property Get DisThumbBorderColor() As OLE_COLOR
Attribute DisThumbBorderColor.VB_Description = "Color of scrollbar thumb border when control is disabled."
Attribute DisThumbBorderColor.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisThumbBorderColor = m_DisThumbBorderColor
End Property

Public Property Let DisThumbBorderColor(ByVal New_DisThumbBorderColor As OLE_COLOR)
   m_DisThumbBorderColor = New_DisThumbBorderColor
   PropertyChanged "DisThumbBorderColor"
End Property

Public Property Get DisThumbColor1() As OLE_COLOR
Attribute DisThumbColor1.VB_Description = "First color of scrollbar thumb gradient when control is disabled."
Attribute DisThumbColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisThumbColor1 = m_DisThumbColor1
End Property

Public Property Let DisThumbColor1(ByVal New_DisThumbColor1 As OLE_COLOR)
   m_DisThumbColor1 = New_DisThumbColor1
   PropertyChanged "DisThumbColor1"
End Property

Public Property Get DisThumbColor2() As OLE_COLOR
Attribute DisThumbColor2.VB_Description = "Second color of scrollbar thumb gradient when control is disabled."
Attribute DisThumbColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisThumbColor2 = m_DisThumbColor2
End Property

Public Property Let DisThumbColor2(ByVal New_DisThumbColor2 As OLE_COLOR)
   m_DisThumbColor2 = New_DisThumbColor2
   PropertyChanged "DisThumbColor2"
End Property

Public Property Get DisTrackbarColor1() As OLE_COLOR
Attribute DisTrackbarColor1.VB_Description = "First color of scrollbar trackbar gradient when control is disabled."
Attribute DisTrackbarColor1.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisTrackbarColor1 = m_DisTrackbarColor1
End Property

Public Property Let DisTrackbarColor1(ByVal New_DisTrackbarColor1 As OLE_COLOR)
   m_DisTrackbarColor1 = New_DisTrackbarColor1
   PropertyChanged "DisTrackbarColor1"
End Property

Public Property Get DisTrackbarColor2() As OLE_COLOR
Attribute DisTrackbarColor2.VB_Description = "Second color of scrollbar trackbar gradient when control is disabled."
Attribute DisTrackbarColor2.VB_ProcData.VB_Invoke_Property = ";Disabled"
   DisTrackbarColor2 = m_DisTrackbarColor2
End Property

Public Property Let DisTrackbarColor2(ByVal New_DisTrackbarColor2 As OLE_COLOR)
   m_DisTrackbarColor2 = New_DisTrackbarColor2
   PropertyChanged "DisTrackbarColor2"
End Property

Public Property Get DragEnabled() As Boolean
Attribute DragEnabled.VB_Description = "If True, drag-and drop of listitems is enabled."
   DragEnabled = m_DragEnabled
End Property

Public Property Let DragEnabled(ByVal New_DragEnabled As Boolean)
   m_DragEnabled = New_DragEnabled
   PropertyChanged "DragEnabled"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Allows user to enable or disable the control."
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If
   InitListBoxDisplayCharacteristics
   PropertyChanged "Enabled"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get FocusBorderColor1() As OLE_COLOR
   FocusBorderColor1 = m_FocusBorderColor1
End Property

Public Property Let FocusBorderColor1(ByVal New_FocusBorderColor1 As OLE_COLOR)
   m_FocusBorderColor1 = New_FocusBorderColor1
   PropertyChanged "FocusBorderColor1"
   InitBorder
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get FocusBorderColor2() As OLE_COLOR
   FocusBorderColor2 = m_FocusBorderColor2
End Property

Public Property Let FocusBorderColor2(ByVal New_FocusBorderColor2 As OLE_COLOR)
   m_FocusBorderColor2 = New_FocusBorderColor2
   PropertyChanged "FocusBorderColor2"
   InitBorder
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get FocusRectColor() As OLE_COLOR
Attribute FocusRectColor.VB_Description = "The color of the listitem focus rectangle."
   FocusRectColor = m_FocusRectColor
End Property

Public Property Let FocusRectColor(ByVal New_FocusRectColor As OLE_COLOR)
   m_FocusRectColor = New_FocusRectColor
   PropertyChanged "FocusRectColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hwnd = UserControl.hwnd
End Property

Public Property Get ImageIndex(ByVal Index As Long) As Long
   ImageIndex = ImageIndexArray(Index)
End Property

Public Property Let ImageIndex(ByVal Index As Long, New_ImageIndex As Long)
   ImageIndexArray(Index) = New_ImageIndex
   If m_RedrawFlag And InDisplayedItemRange(Index) Then
      DisplayList
   End If
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
   ItemData = ItemDataArray(Index)
End Property

Public Property Let ItemData(ByVal Index As Long, NewValue As Long)
   ItemDataArray(Index) = NewValue
End Property

Public Property Get ItemImageSize() As Long
Attribute ItemImageSize.VB_Description = "If 0, listitem image width/height is the same as the height of item text.  Otherwise specifies height and width of listitem image."
   ItemImageSize = m_ItemImageSize
End Property

Public Property Let ItemImageSize(ByVal New_ItemImageSize As Long)
   m_ItemImageSize = New_ItemImageSize
   InitListBoxDisplayCharacteristics
   RedrawControl
   PropertyChanged "ItemImageSize"
End Property

Public Property Get List(ByVal Index As Long) As String
   If Index >= 0 And Index <= UBound(ListArray) Then
      List = ListArray(Index)
   End If
End Property

Public Property Get ListCount() As Long
   ListCount = m_ListCount
End Property

Public Property Get ListFont() As Font
Attribute ListFont.VB_Description = "The font used to display listitem text."
   Set ListFont = m_ListFont
End Property

Public Property Set ListFont(ByVal New_ListFont As Font)
   Set m_ListFont = New_ListFont
   Set UserControl.Font = m_ListFont
'  get the height range of characters in the current font.
   ListItemHeight = TextHeight("^j")
   InitTextDisplayCharacteristics
   PropertyChanged "ListFont"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "The index of the most recently selected list item."
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = ";Text"
   ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)

'*************************************************************************
'* processes programmatic alteration of the .ListIndex property.         *
'*************************************************************************

'  can't modify this property in design mode.
   If Ambient.UserMode = False Then Err.Raise 387

   m_ListIndex = New_ListIndex

   If m_Style = [Standard] And m_MultiSelect = vbMultiSelectNone Then
'     in MultiSelect None mode, if the new .ListIndex is -1,  clear all selections.  If
'     it's > -1, clear any existing selection and select the item pointed to by .ListIndex.
      If m_ListIndex = -1 Then
         SetSelectedArrayRange 0, m_ListCount - 1, False
         LastSelectedItem = m_ListIndex
         m_SelCount = 0
         ItemWithFocus = 0
      Else
         SetSelectedArrayRange 0, m_ListCount - 1, False
         LastSelectedItem = m_ListIndex
         m_SelCount = 1
         SelectedArray(m_ListIndex) = True
         ItemWithFocus = m_ListIndex
      End If
   Else
'     for MultiSelect Simple, MultiSelect Extended and CheckBox modes,
'     move the focus rectangle to the item pointed to by .ListIndex.
      ItemWithFocus = m_ListIndex
      If m_Style = [CheckBox] Then
'        in CheckBox mode, selection bar always moves with the focus rectangle.
         LastSelectedItem = m_ListIndex
      End If
   End If

   DisplayList

   PropertyChanged "ListIndex"

End Property

Public Property Get MultiSelect() As SelectionOptions
Attribute MultiSelect.VB_Description = "The listitem multiple selection mode  (None, Simple, or Extended)."
   MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As SelectionOptions)
'  can't change at runtime.
   If Ambient.UserMode Then Err.Raise 382
   m_MultiSelect = New_MultiSelect
   PropertyChanged "MultiSelect"
End Property

Public Property Get NewIndex() As Long
   NewIndex = m_NewIndex
End Property

Public Property Let NewIndex(ByVal New_NewIndex As Long)
'  this makes the property unavailable at design time.
   If Ambient.UserMode = False Then Err.Raise 387
   m_NewIndex = New_NewIndex
   PropertyChanged "NewIndex"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "The bitmap to display in lieu of a gradient in the ListBox background."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set m_Picture = New_Picture
   Set m_ActivePicture = m_Picture
   PropertyChanged "Picture"
'  this flag tells Redraw to re-blit the new background to the virtual DC.
   ChangingPicture = True
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get PictureMode() As MLB_PictureModeOptions
Attribute PictureMode.VB_Description = "The method used to render the background image: Normal, Stretched or Tiled."
   If Ambient.UserMode Then Err.Raise 393
   PictureMode = m_PictureMode
End Property

Public Property Let PictureMode(ByVal New_PictureMode As MLB_PictureModeOptions)
'  not available at runtime.
   If Ambient.UserMode Then Err.Raise 382
   m_PictureMode = New_PictureMode
   m_ActivePictureMode = m_PictureMode
   RedrawControl
   UserControl.Refresh
   PropertyChanged "PictureMode"
End Property

Public Property Get RedrawFlag() As Boolean
   If Ambient.UserMode Then Err.Raise 393
   RedrawFlag = m_RedrawFlag
End Property

Public Property Let RedrawFlag(ByVal New_RedrawFlag As Boolean)
   If Ambient.UserMode = False Then Err.Raise 387    ' read-only at design time.
   m_RedrawFlag = New_RedrawFlag
   PropertyChanged "RedrawFlag"
   RedrawControl ' if RedrawFlag is now True, the control will redraw.
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "If True, MorphListBox graphics and text are arranged to provide a natural feel for those users whose written language is read from right to left."
   RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
   m_RightToLeft = New_RightToLeft
   RedrawControl
   PropertyChanged "RightToLeft"
End Property

Public Property Get ScaleHeight() As Long
   ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScaleMode() As Integer
   ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
   m_ScaleMode = New_ScaleMode
   UserControl.ScaleMode = m_ScaleMode
   PropertyChanged "ScaleMode"
End Property

Public Property Get ScaleWidth() As Long
   ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get SelColor1() As OLE_COLOR
Attribute SelColor1.VB_Description = "The first gradient color of the listitem selection bar."
   SelColor1 = m_SelColor1
End Property

Public Property Let SelColor1(ByVal New_SelColor1 As OLE_COLOR)
   m_SelColor1 = New_SelColor1
   PropertyChanged "SelColor1"
   CalculateHighlightBarGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get SelColor2() As OLE_COLOR
Attribute SelColor2.VB_Description = "The second gradient color of the listitem selection bar."
   SelColor2 = m_SelColor2
End Property

Public Property Let SelColor2(ByVal New_SelColor2 As OLE_COLOR)
   m_SelColor2 = New_SelColor2
   PropertyChanged "SelColor2"
   CalculateHighlightBarGradient
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get SelCount() As Long
Attribute SelCount.VB_MemberFlags = "400"
   SelCount = m_SelCount
End Property

Public Property Get Selected(ByVal Index As Long) As Boolean
   Selected = SelectedArray(Index)
End Property

Public Property Let Selected(ByVal Index As Long, NewValue As Boolean)

'*************************************************************************
'* processes programmatic selection/deselection of specified list item.  *
'* all 3 multiple selection modes (MultiSelect Simple, MultiSelect Exten-*
'* ded and CheckBox) are processed the same way here; only MultiSelect   *
'* None mode is treated differently.                                     *
'*************************************************************************

   Dim PreviouslySelected As Boolean

   PreviouslySelected = SelectedArray(Index)

   If m_Style = [Standard] And m_MultiSelect = vbMultiSelectNone Then
'     in MultiSelect None mode, if Selected(Index) is set to True, all other list items
'     are deselected and variables are set to the index of the newly selected item.
      SetSelectedArrayRange 0, m_ListCount - 1, False
      SelectedArray(Index) = NewValue
      If NewValue Then
         m_ListIndex = Index
         m_SelCount = 1
         LastSelectedItem = Index
         ItemWithFocus = Index
      Else
'        if Selected(Index) is set to False, all items are deselected.  .ListIndex
'        property is set to -1, which the default in this mode for no selected items.
         If PreviouslySelected Then
            m_ListIndex = -1
            m_SelCount = 0
            LastSelectedItem = -1
            ItemWithFocus = 0
         End If
      End If
   Else
'     in modes that allow multiple selections, the item is selected or deselected, and
'     the .SelCount property is adjusted accordingly.  .ListIndex points to Index item.
      SelectedArray(Index) = NewValue
      m_ListIndex = Index
      ItemWithFocus = Index
      LastSelectedItem = Index
      If PreviouslySelected And Not NewValue Then
         m_SelCount = m_SelCount - 1
      ElseIf Not PreviouslySelected And NewValue Then
         m_SelCount = m_SelCount + 1
      End If
   End If

   If m_RedrawFlag Then ' added if 01/25/06 to account for many selections via code.
      DisplayList
   End If

End Property

Public Property Get SelTextColor() As OLE_COLOR
Attribute SelTextColor.VB_Description = "The color of listitem text covered by selection bar."
   SelTextColor = m_SelTextColor
End Property

Public Property Let SelTextColor(ByVal New_SelTextColor As OLE_COLOR)
   m_SelTextColor = New_SelTextColor
   PropertyChanged "SelTextColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get ShowItemImages() As Boolean
Attribute ShowItemImages.VB_Description = "If True, listitem images (small bitmaps or icons) can be loaded and assigned to listitems."
   ShowItemImages = m_ShowItemImages
End Property

Public Property Let ShowItemImages(ByVal New_ShowItemImages As Boolean)
   m_ShowItemImages = New_ShowItemImages
   InitListBoxDisplayCharacteristics
   RedrawControl
   PropertyChanged "ShowItemImages"
End Property

Public Property Get ShowSelectRect() As Boolean
Attribute ShowSelectRect.VB_Description = "If True, listitem focus rectangle is displayed."
   ShowSelectRect = m_ShowSelectRect
End Property

Public Property Let ShowSelectRect(ByVal New_ShowSelectRect As Boolean)
   m_ShowSelectRect = New_ShowSelectRect
   PropertyChanged "ShowSelectRect"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get SortAsNumeric() As Boolean
Attribute SortAsNumeric.VB_Description = "If True, listitems are sorted in numerical order (numbers sort improperly as strings otherwise)."
   SortAsNumeric = m_SortAsNumeric
End Property

Public Property Let SortAsNumeric(ByVal New_SortAsNumeric As Boolean)
   m_SortAsNumeric = New_SortAsNumeric
   PropertyChanged "SortAsNumeric"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "When True, ListBox items are automatically maintained in ascending order."
Attribute Sorted.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Sorted = m_Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
   m_Sorted = New_Sorted
   If m_Sorted And Not ListIsSorted Then
      Sort
   ElseIf Not m_Sorted Then
      ListIsSorted = False
   End If
   PropertyChanged "Sorted"
End Property

Public Property Get Style() As ListItemOptions
Attribute Style.VB_Description = "Sets the display and operation of the ListBox to either Standard or CheckBox styles.  CheckBox style supercedes all .MultiSelect operation modes."
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Operation Modes"
   Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As ListItemOptions)
   If Ambient.UserMode Then Err.Raise 382    ' not supported at runtime.
   m_Style = New_Style
   PropertyChanged "Style"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "The text of the most recently selected listitem."
   Text = ListArray(m_ListIndex)
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "The color of non-selected listitem text."
   TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
   m_TextColor = New_TextColor
   PropertyChanged "TextColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get Theme() As ThemeOptions
Attribute Theme.VB_Description = "Selects a color scheme to apply to the control.  Eight schemes are built in; user may add more."
   Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As ThemeOptions)

'*************************************************************************
'* changes color scheme of listbox to one of eight predefined themes.    *
'*************************************************************************

   m_Theme = New_Theme

   Select Case m_Theme

      Case [Cyan Eyed]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H404000
         m_ArrowUpColor = &HFFFF00
         m_BackAngle = 45
         m_BackColor1 = &H808000
         m_BackColor2 = &HFFFF80
         m_BackMiddleOut = True
         m_BorderColor1 = &H404000
         m_BorderColor2 = &HFFFF00
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H404000
         m_ButtonColor2 = &H808000
         m_CheckBoxColor = &H404000
         m_CheckBoxArrowColor = &H404000
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H404000
         m_FocusBorderColor2 = &HC0C000
         m_FocusRectColor = &HFFFFC0
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H808000
         m_SelColor2 = &H808000
         m_SelTextColor = &HFFFFC0
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HFFFF00
         m_ThumbColor1 = &H404000
         m_ThumbColor2 = &H808000
         m_TrackBarColor1 = &H808000
         m_TrackBarColor2 = &HFFFFC0
         m_TrackClickColor1 = &H404000
         m_TrackClickColor2 = &HFFFF80

      Case [Gunmetal Grey]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H0
         m_ArrowUpColor = &HE0E0E0
         m_BackAngle = 45
         m_BackColor1 = &H606060
         m_BackColor2 = &HE0E0E0
         m_BackMiddleOut = True
         m_BorderColor1 = &H0
         m_BorderColor2 = &HE0E0E0
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H0
         m_ButtonColor2 = &HC0C0C0
         m_CheckBoxColor = &H0
         m_CheckBoxArrowColor = &H0
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H0
         m_FocusBorderColor2 = &H808080
         m_FocusRectColor = &HFFFFFF
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H404040
         m_SelColor2 = &H404040
         m_SelTextColor = &HE0E0E0
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HE0E0E0
         m_ThumbColor1 = &H0
         m_ThumbColor2 = &H909090
         m_TrackBarColor1 = &H606060
         m_TrackBarColor2 = &HE0E0E0
         m_TrackClickColor1 = &H0
         m_TrackClickColor2 = &HE0E0E0

      Case [Blue Moon]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H400000
         m_ArrowUpColor = &HFFC0C0
         m_BackAngle = 45
         m_BackColor1 = &HC00000
         m_BackColor2 = &HFFC0C0
         m_BackMiddleOut = True
         m_BorderColor1 = &H400000
         m_BorderColor2 = &HFF8080
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H400000
         m_ButtonColor2 = &HFF8080
         m_CheckBoxColor = &H400000
         m_CheckBoxArrowColor = &H400000
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H400000
         m_FocusBorderColor2 = &HFF0000
         m_FocusRectColor = &HFFC0C0
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H800000
         m_SelColor2 = &H800000
         m_SelTextColor = &HFFC0C0
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HFFC0C0
         m_ThumbColor1 = &H400000
         m_ThumbColor2 = &HFF8080
         m_TrackBarColor1 = &H800000
         m_TrackBarColor2 = &HFFC0C0
         m_TrackClickColor1 = &H400000
         m_TrackClickColor2 = &HFF8080

      Case [Red Rum]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H40
         m_ArrowUpColor = &HC0C0FF
         m_BackAngle = 45
         m_BackColor1 = &H80
         m_BackColor2 = &HC0C0FF
         m_BackMiddleOut = True
         m_BorderColor1 = &H40
         m_BorderColor2 = &H8080FF
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H40
         m_ButtonColor2 = &H8080FF
         m_CheckBoxColor = &H40
         m_CheckBoxArrowColor = &H40&
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H40
         m_FocusBorderColor2 = &HFF&
         m_FocusRectColor = &HC0C0FF
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H80&
         m_SelColor2 = &H80&
         m_SelTextColor = &HC0C0FF
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HC0C0FF
         m_ThumbColor1 = &H40
         m_ThumbColor2 = &H8080FF
         m_TrackBarColor1 = &H80
         m_TrackBarColor2 = &HC0C0FF
         m_TrackClickColor1 = &H40&
         m_TrackClickColor2 = &H8080FF

      Case [Green With Envy]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H4000&
         m_ArrowUpColor = &HC0FFC0
         m_BackAngle = 45
         m_BackColor1 = &H8000&
         m_BackColor2 = &HC0FFC0
         m_BackMiddleOut = True
         m_BorderColor1 = &H4000&
         m_BorderColor2 = &H80FF80
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H4000&
         m_ButtonColor2 = &H80FF80
         m_CheckBoxColor = &H4000&
         m_CheckBoxArrowColor = &H4000&
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H4000&
         m_FocusBorderColor2 = &HC000&
         m_FocusRectColor = &HC0FFC0
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H8000&
         m_SelColor2 = &H8000&
         m_SelTextColor = &HC0FFC0
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HC0FFC0
         m_ThumbColor1 = &H4000&
         m_ThumbColor2 = &HFF00&
         m_TrackBarColor1 = &H8000&
         m_TrackBarColor2 = &HC0FFC0
         m_TrackClickColor1 = &H4000&
         m_TrackClickColor2 = &H80FF80

      Case [Purple People Eater]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H800080
         m_ArrowUpColor = &HFF00FF
         m_BackAngle = 45
         m_BackColor1 = &H800080
         m_BackColor2 = &HFF80FF
         m_BackMiddleOut = True
         m_BorderColor1 = &H400040
         m_BorderColor2 = &HFF80FF
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H400040
         m_ButtonColor2 = &H800080
         m_CheckBoxColor = &H400040
         m_CheckBoxArrowColor = &H400040
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H400040
         m_FocusBorderColor2 = &HC000C0
         m_FocusRectColor = &HFFC0FF
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H800080
         m_SelColor2 = &H800080
         m_SelTextColor = &HFFC0FF
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HFF00FF
         m_ThumbColor1 = &H400040
         m_ThumbColor2 = &H800080
         m_TrackBarColor1 = &H800080
         m_TrackBarColor2 = &HFFC0FF
         m_TrackClickColor1 = &H400040
         m_TrackClickColor2 = &HFF80FF

      Case [Golden Goose]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H8080&
         m_ArrowUpColor = &HFFFF&
         m_BackAngle = 45
         m_BackColor1 = &H8080&
         m_BackColor2 = &H80FFFF
         m_BackMiddleOut = True
         m_BorderColor1 = &H4040&
         m_BorderColor2 = &H80FFFF
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H4040&
         m_ButtonColor2 = &H8080&
         m_CheckBoxColor = &H4040&
         m_CheckBoxArrowColor = &H4040&
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H4040&
         m_FocusBorderColor2 = &HC0C0&
         m_FocusRectColor = &HC0FFFF
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H8080&
         m_SelColor2 = &H8080&
         m_SelTextColor = &HC0FFFF
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &HFFFF&
         m_ThumbColor1 = &H4040&
         m_ThumbColor2 = &H8080&
         m_TrackBarColor1 = &H8080&
         m_TrackBarColor2 = &HC0FFFF
         m_TrackClickColor1 = &H4040&
         m_TrackClickColor2 = &H80FFFF

      Case [Penny Wise]
         Set m_Picture = Nothing
         Set m_DisPicture = Nothing
         m_ArrowDownColor = &H4080&
         m_ArrowUpColor = &HC0E0FF
         m_BackAngle = 45
         m_BackColor1 = &H4080&
         m_BackColor2 = &H80C0FF
         m_BackMiddleOut = True
         m_BorderColor1 = &H404080
         m_BorderColor2 = &H80C0FF
         m_BorderMiddleOut = True
         m_BorderWidth = 16
         m_ButtonColor1 = &H404080
         m_ButtonColor2 = &H80FF&
         m_CheckBoxColor = &H0
         m_CheckBoxArrowColor = &H0
         m_CheckStyle = Tick
         m_CircularGradient = False
         m_CurveTopLeft = 0
         m_CurveTopRight = 0
         m_CurveBottomLeft = 0
         m_CurveBottomRight = 0
         m_DisArrowDownColor = &HC0C0C0
         m_DisArrowUpColor = &HC0C0C0
         m_DisBackColor1 = &H808080
         m_DisBackColor2 = &HC0C0C0
         m_DisBorderColor1 = &H0
         m_DisBorderColor2 = &H0
         m_DisButtonColor1 = &H404040
         m_DisButtonColor2 = &H808080
         m_DisCheckboxArrowColor = &H0
         m_DisCheckboxColor = &H0
         m_DisFocusRectColor = &H808080
         m_DisSelColor1 = &H808080
         m_DisSelColor2 = &HC0C0C0
         m_DisSelTextColor = &H808080
         m_DisTextColor = &H404040
         m_DisThumbBorderColor = &H808080
         m_DisThumbColor1 = &H404040
         m_DisThumbColor2 = &H808080
         m_DisTrackbarColor1 = &H808080
         m_DisTrackbarColor2 = &HC0C0C0
         m_FocusBorderColor1 = &H404080
         m_FocusBorderColor2 = &H80FF&
         m_FocusRectColor = &HC0E0FF
         m_MultiSelect = 0
         m_RightToLeft = False
         m_SelColor1 = &H4080&
         m_SelColor2 = &H4080&
         m_SelTextColor = &HC0E0FF
         m_ShowItemImages = False
         m_ShowSelectRect = True
         m_SortAsNumeric = False
         m_Sorted = False
         m_Style = Standard
         m_TextColor = &H0
         m_ThumbBorderColor = &H80FF&
         m_ThumbColor1 = &H404080
         m_ThumbColor2 = &H40C0&
         m_TrackBarColor1 = &H4080&
         m_TrackBarColor2 = &HC0E0FF
         m_TrackClickColor1 = &H404080
         m_TrackClickColor2 = &H80C0FF

   End Select

   PropertyChanged "Theme"

   GetEnabledDisplayProperties
   InitListBoxDisplayCharacteristics
   RedrawControl
   UserControl.Refresh

End Property

Public Property Get ThumbColor1() As OLE_COLOR
Attribute ThumbColor1.VB_Description = "The first gradient color of the scrollbar thumb."
   ThumbColor1 = m_ThumbColor1
End Property

Public Property Let ThumbColor1(ByVal New_ThumbColor1 As OLE_COLOR)
   m_ThumbColor1 = New_ThumbColor1
   PropertyChanged "ThumbColor1"
   CalculateScrollbarThumbGradient
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get ThumbColor2() As OLE_COLOR
Attribute ThumbColor2.VB_Description = "The second gradient color of the scrollbar thumb."
   ThumbColor2 = m_ThumbColor2
End Property

Public Property Let ThumbColor2(ByVal New_ThumbColor2 As OLE_COLOR)
   m_ThumbColor2 = New_ThumbColor2
   PropertyChanged "ThumbColor2"
   CalculateScrollbarThumbGradient
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get ThumbBorderColor() As OLE_COLOR
Attribute ThumbBorderColor.VB_Description = "The color of the scrollbar thumb border."
   ThumbBorderColor = m_ThumbBorderColor
End Property

Public Property Let ThumbBorderColor(ByVal New_ThumbBorderColor As OLE_COLOR)
   m_ThumbBorderColor = New_ThumbBorderColor
   PropertyChanged "ThumbBorderColor"
   RedrawControl
   UserControl.Refresh
End Property

Public Property Get TopIndex() As Long
'  Note:  In the standard VB listbox, the .TopIndex property is read/write, and you can set
'  the property to display list items beginning with the specified index.  In this control,
'  the .DisplayFrom method replaces the .TopIndex write functionality. Therefore, .TopIndex
'  here is read-only and returns the index of the first displayed list item.
   TopIndex = m_TopIndex
End Property

Public Property Get TrackBarColor1() As OLE_COLOR
   TrackBarColor1 = m_TrackBarColor1
End Property

Public Property Let TrackBarColor1(ByVal New_TrackBarColor1 As OLE_COLOR)
Attribute TrackBarColor1.VB_Description = "The first gradient color of the vertical scrollbar trackbar."
Attribute TrackBarColor1.VB_ProcData.VB_Invoke_PropertyPut = ";Vertical Scrollbar"
   m_TrackBarColor1 = New_TrackBarColor1
   PropertyChanged "TrackBarColor1"
   CalculateVerticalTrackbarGradientUnclicked
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get TrackBarColor2() As OLE_COLOR
Attribute TrackBarColor2.VB_Description = "The second gradient color of the vertical scrollbar trackbar."
Attribute TrackBarColor2.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackBarColor2 = m_TrackBarColor2
End Property

Public Property Let TrackBarColor2(ByVal New_TrackBarColor2 As OLE_COLOR)
   m_TrackBarColor2 = New_TrackBarColor2
   PropertyChanged "TrackBarColor2"
   CalculateVerticalTrackbarGradientUnclicked
   DisplayVerticalScrollBar
   UserControl.Refresh
End Property

Public Property Get TrackClickColor1() As OLE_COLOR
Attribute TrackClickColor1.VB_Description = "First gradient color of portion of scroll track above or below thumb when it is clicked."
Attribute TrackClickColor1.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackClickColor1 = m_TrackClickColor1
End Property

Public Property Let TrackClickColor1(ByVal New_TrackClickColor1 As OLE_COLOR)
   m_TrackClickColor1 = New_TrackClickColor1
   PropertyChanged "TrackClickColor1"
   CalculateVerticalTrackbarGradientClicked
   DisplayVerticalTrackBar
   UserControl.Refresh
End Property

Public Property Get TrackClickColor2() As OLE_COLOR
Attribute TrackClickColor2.VB_Description = "Second gradient color of portion of scroll track above or below thumb when it is clicked."
Attribute TrackClickColor2.VB_ProcData.VB_Invoke_Property = ";Vertical Scrollbar"
   TrackClickColor2 = m_TrackClickColor2
End Property

Public Property Let TrackClickColor2(ByVal New_TrackClickColor2 As OLE_COLOR)
   m_TrackClickColor2 = New_TrackClickColor2
   PropertyChanged "TrackClickColor2"
   CalculateVerticalTrackbarGradientClicked
   DisplayVerticalTrackBar
   UserControl.Refresh
End Property

Private Sub GetEnabledDisplayProperties()

'*************************************************************************
'* applies enabled graphics properties to the active display properties. *
'*************************************************************************

   Set m_ActivePicture = m_Picture
   m_ActiveArrowDownColor = m_ArrowDownColor
   m_ActiveArrowUpColor = m_ArrowUpColor
   m_ActiveBackColor1 = m_BackColor1
   m_ActiveBackColor2 = m_BackColor2
   m_ActiveBorderColor1 = m_BorderColor1
   m_ActiveBorderColor2 = m_BorderColor2
   m_ActiveButtonColor1 = m_ButtonColor1
   m_ActiveButtonColor2 = m_ButtonColor2
   m_ActiveCheckboxArrowColor = m_CheckBoxArrowColor
   m_ActiveCheckBoxColor = m_CheckBoxColor
   m_ActiveFocusRectColor = m_FocusRectColor
   m_ActivePictureMode = m_PictureMode
   m_ActiveSelColor1 = m_SelColor1
   m_ActiveSelColor2 = m_SelColor2
   m_ActiveSelTextColor = m_SelTextColor
   m_ActiveTextColor = m_TextColor
   m_ActiveThumbBorderColor = m_ThumbBorderColor
   m_ActiveThumbColor1 = m_ThumbColor1
   m_ActiveThumbColor2 = m_ThumbColor2
   m_ActiveTrackBarColor1 = m_TrackBarColor1
   m_ActiveTrackBarColor2 = m_TrackBarColor2

End Sub

Private Sub GetDisabledDisplayProperties()

'*************************************************************************
'* applies disabled graphics properties to active display properties.    *
'*************************************************************************

   Set m_ActivePicture = m_DisPicture
   m_ActiveArrowDownColor = m_DisArrowDownColor
   m_ActiveArrowUpColor = m_DisArrowUpColor
   m_ActiveBackColor1 = m_DisBackColor1
   m_ActiveBackColor2 = m_DisBackColor2
   m_ActiveBorderColor1 = m_DisBorderColor1
   m_ActiveBorderColor2 = m_DisBorderColor2
   m_ActiveButtonColor2 = m_DisButtonColor2
   m_ActiveButtonColor2 = m_DisButtonColor2
   m_ActiveCheckboxArrowColor = m_DisCheckboxArrowColor
   m_ActiveCheckBoxColor = m_DisCheckboxColor
   m_ActiveFocusRectColor = m_DisFocusRectColor
   m_ActivePictureMode = m_DisPictureMode
   m_ActiveSelColor1 = m_DisSelColor1
   m_ActiveSelColor2 = m_DisSelColor2
   m_ActiveSelTextColor = m_DisSelTextColor
   m_ActiveTextColor = m_DisTextColor
   m_ActiveThumbBorderColor = m_DisThumbBorderColor
   m_ActiveThumbColor1 = m_DisThumbColor1
   m_ActiveThumbColor2 = m_DisThumbColor2
   m_ActiveTrackBarColor1 = m_DisTrackbarColor1
   m_ActiveTrackBarColor2 = m_DisTrackbarColor2

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Subclassing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<< All subclassing code by Paul Caton. >>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub StartSubclassing()

'*************************************************************************
'* starts up Paul Caton's self-subclassing code.                         *
'*************************************************************************

   If Ambient.UserMode Then                                    ' if we're not in design mode.
      With UserControl
         Call Subclass_Start(.hwnd)                            ' Start subclassing.
         Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)  ' for mouse enter detect.
         Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER) ' for mouse leave detect.
         Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER) ' for mouse wheel detect.
         Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)   ' for got focus detect.
         Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)  ' for lost focus detect.
      End With
   End If

End Sub

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

'*************************************************************************
'* track the mouse leaving the indicated window.                         *
'*************************************************************************

   Dim tme As TRACKMOUSEEVENT_STRUCT

   With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
   End With

   Call TrackMouseEventComCtl(tme)

End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
   With sc_aSubData(zIdx(lng_hWnd))
      If When And eMsgWhen.MSG_BEFORE Then
         Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
      End If
      If When And eMsgWhen.MSG_AFTER Then
         Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
      End If
   End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
   Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long

'  Parameters:
'  lng_hWnd  - The handle of the window to be subclassed
'  Returns;
'  The sc_aSubData() index
   Const CODE_LEN              As Long = 204                      'Length of the machine code in bytes
   Const FUNC_CWP              As String = "CallWindowProcA"      'We use CallWindowProc to call the original WndProc
   Const FUNC_EBM              As String = "EbMode"               'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
   Const FUNC_SWL              As String = "SetWindowLongA"       'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
   Const MOD_USER              As String = "user32"               'Location of the SetWindowLongA & CallWindowProc functions
   Const MOD_VBA5              As String = "vba5"                 'Location of the EbMode function if running VB5
   Const MOD_VBA6              As String = "vba6"                 'Location of the EbMode function if running VB6
   Const PATCH_01              As Long = 18                       'Code buffer offset to the location of the relative address to EbMode
   Const PATCH_02              As Long = 68                       'Address of the previous WndProc
   Const PATCH_03              As Long = 78                       'Relative address of SetWindowsLong
   Const PATCH_06              As Long = 116                      'Address of the previous WndProc
   Const PATCH_07              As Long = 121                      'Relative address of CallWindowProc
   Const PATCH_0A              As Long = 186                      'Address of the owner object
   Static aBuf(1 To CODE_LEN)  As Byte                            'Static code buffer byte array
   Static pCWP                 As Long                            'Address of the CallWindowsProc
   Static pEbMode              As Long                            'Address of the EbMode IDE break/stop/running function
   Static pSWL                 As Long                            'Address of the SetWindowsLong function
   Dim i                       As Long                            'Loop index
   Dim j                       As Long                            'Loop index
   Dim nSubIdx                 As Long                            'Subclass data index
   Dim sHex                    As String                          'Hex code string

'  If it's the first time through here..
   If aBuf(1) = 0 Then

'     The hex pair machine code representation.
      sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
             "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
             "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
             "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'     Convert the string from hex pairs to bytes and store in the static machine code buffer
      i = 1
      Do While j < CODE_LEN
         j = j + 1
         aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                   'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
         i = i + 2
      Loop                                                        'Next pair of hex characters

'     Get API function addresses
      If Subclass_InIDE Then                                      'If we're running in the VB IDE
         aBuf(16) = &H90                                          'Patch the code buffer to enable the IDE state code
         aBuf(17) = &H90                                          'Patch the code buffer to enable the IDE state code
         pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                  'Get the address of EbMode in vba6.dll
         If pEbMode = 0 Then                                      'Found?
            pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)               'VB5 perhaps
         End If
      End If

      pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                        'Get the address of the CallWindowsProc function
      pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                        'Get the address of the SetWindowLongA function
      ReDim sc_aSubData(0 To 0) As tSubData                       'Create the first sc_aSubData element

   Else

      nSubIdx = zIdx(lng_hWnd, True)
      If nSubIdx = -1 Then                                        'If an sc_aSubData element isn't being re-cycled
         nSubIdx = UBound(sc_aSubData()) + 1                      'Calculate the next element
         ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData     'Create a new sc_aSubData element
      End If

      Subclass_Start = nSubIdx

   End If

   With sc_aSubData(nSubIdx)
      .hwnd = lng_hWnd                                            'Store the hWnd
      .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)               'Allocate memory for the machine code WndProc
      .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)  'Set our WndProc in place
      Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)      'Copy the machine code from the static byte array to the code array in sc_aSubData
      Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
      Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)             'Original WndProc address for CallWindowProc, call the original WndProc
      Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                   'Patch the relative address of the SetWindowLongA api function
      Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
      Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                   'Patch the relative address of the CallWindowProc api function
      Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))             'Patch the address of this object instance into the static machine code buffer
   End With

End Function

'Stop all subclassing
Private Sub Subclass_StopAll()

   Dim i As Long

   i = UBound(sc_aSubData())                                      'Get the upper bound of the subclass data array
   Do While i >= 0                                                'Iterate through each element
      With sc_aSubData(i)
         If .hwnd <> 0 Then                                       'If not previously Subclass_Stop'd
            Call Subclass_Stop(.hwnd)                             'Subclass_Stop
         End If
      End With
      i = i - 1                                                   'Next element
   Loop

End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)

'  Parameters:
'  lng_hWnd  - The handle of the window to stop being subclassed
   With sc_aSubData(zIdx(lng_hWnd))
      Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)         'Restore the original WndProc
      Call zPatchVal(.nAddrSub, PATCH_05, 0)                      'Patch the Table B entry count to ensure no further 'before' callbacks
      Call zPatchVal(.nAddrSub, PATCH_09, 0)                      'Patch the Table A entry count to ensure no further 'after' callbacks
      Call GlobalFree(.nAddrSub)                                  'Release the machine code memory
      .hwnd = 0                                                   'Mark the sc_aSubData element as available for re-use
      .nMsgCntB = 0                                               'Clear the before table
      .nMsgCntA = 0                                               'Clear the after table
      Erase .aMsgTblB                                             'Erase the before table
      Erase .aMsgTblA                                             'Erase the after table
   End With

End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)

   Dim nEntry  As Long                                            'Message table entry index
   Dim nOff1   As Long                                            'Machine code buffer offset 1
   Dim nOff2   As Long                                            'Machine code buffer offset 2

   If uMsg = ALL_MESSAGES Then                                    'If all messages
      nMsgCnt = ALL_MESSAGES                                      'Indicates that all messages will callback
   Else                                                           'Else a specific message number
      Do While nEntry < nMsgCnt                                   'For each existing entry. NB will skip if nMsgCnt = 0
         nEntry = nEntry + 1
         If aMsgTbl(nEntry) = 0 Then                              'This msg table slot is a deleted entry
            aMsgTbl(nEntry) = uMsg                                'Re-use this entry
            Exit Sub                                              'Bail
         ElseIf aMsgTbl(nEntry) = uMsg Then                       'The msg is already in the table!
            Exit Sub                                              'Bail
         End If
      Loop                                                        'Next entry
      nMsgCnt = nMsgCnt + 1                                       'New slot required, bump the table entry count
      ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                'Bump the size of the table.
      aMsgTbl(nMsgCnt) = uMsg                                     'Store the message number in the table
   End If

   If When = eMsgWhen.MSG_BEFORE Then                             'If before
      nOff1 = PATCH_04                                            'Offset to the Before table
      nOff2 = PATCH_05                                            'Offset to the Before table entry count
   Else                                                           'Else after
      nOff1 = PATCH_08                                            'Offset to the After table
      nOff2 = PATCH_09                                            'Offset to the After table entry count
   End If

   If uMsg <> ALL_MESSAGES Then
      Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
   End If

   Call zPatchVal(nAddr, nOff2, nMsgCnt)                          'Patch the appropriate table entry count

End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
   zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
   Debug.Assert zAddrFunc                                         'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

'  Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
   zIdx = UBound(sc_aSubData)
   Do While zIdx >= 0                                             'Iterate through the existing sc_aSubData() elements
      With sc_aSubData(zIdx)
         If .hwnd = lng_hWnd Then                                 'If the hWnd of this element is the one we're looking for
            If Not bAdd Then                                      'If we're searching not adding
               Exit Function                                      'Found
            End If
         ElseIf .hwnd = 0 Then                                    'If this an element marked for reuse.
            If bAdd Then                                          'If we're adding
               Exit Function                                      'Re-use it
            End If
         End If
      End With
      zIdx = zIdx - 1                                             'Decrement the index
   Loop

  If Not bAdd Then
    Debug.Assert False                                            'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
   Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
   Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
   zSetTrue = True
   bValue = True
End Function
