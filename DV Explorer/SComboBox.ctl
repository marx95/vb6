VERSION 5.00
Begin VB.UserControl SComboBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   KeyPreview      =   -1  'True
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ToolboxBitmap   =   "SComboBox.ctx":0000
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   405
      Top             =   1035
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   -1800
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   -1800
      TabIndex        =   0
      Top             =   135
      Width           =   1155
   End
End
Attribute VB_Name = "SComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2004       *'
'******************************************************'
'*                   Version 1.0.4                    *'
'******************************************************'
'* Control:       SComboBox                           *'
'******************************************************'
'* Author:        Heriberto Mantilla Santamaría       *'
'******************************************************'
'* Collaboration: fred.cpp                            *'
'*                                                    *'
'*                So many thanks for his contribution *'
'*                for this project, some styles and   *'
'*                Traduction to English of some       *'
'*                comments.                           *'
'******************************************************'
'* Description:   This usercontrol simulates a Combo- *'
'*                Box But adds new an great features  *'
'*                like:                               *'
'*                                                    *'
'*                - The first ComboBox On PSC that    *'
'*                  actually works in a single file   *'
'*                  control.                          *'
'*                - When the list is shown doesn't    *'
'*                  deactivate the parent form.       *'
'*                - More than 20 Visual Styles; no    *'
'*                  images Everything done by code.   *'
'*                - Some extra cool properties.       *'
'******************************************************'
'* Started on:    Friday, 11-jun-2004.                *'
'******************************************************'
'*                   Version 1.0.0                    *'
'*                                                    *'
'* Fixes:         - List.                  (18/06/04) *'
'*                - Control Appearance.    (20/06/04) *'
'*                - Standard Appearance.   (21/06/04) *'
'*                - MAC Appearance.        (24/06/04) *'
'*                - XP Appearance.         (25/06/04) *'
'*                - List Elements.         (27/06/04) *'
'*                - List Events.           (27/06/04) *'
'*                - Control properties.    (28/08/09) *'
'*                - Control properties.    (29/08/09) *'
'*                - List.                  (29/06/09) *'
'*                - List.                  (01/07/09) *'
'*                - JAVA Appearance.       (02/07/09) *'
'*                - List.                  (03/07/09) *'
'*                - Soft Style Appearance. (03/07/09) *'
'*                - Ardent Appearance.     (04/07/09) *'
'*                - List.                  (04/07/09) *'
'*                - MAC Appearance.        (04/07/04) *'
'******************************************************'
'*       Errors corrected after the publication       *'
'*                   Version 1.0.1                    *'
'*                                                    *'
'*  - ScrollBar Slider.                    (08/07/04) *'
'*  - ListIndex Property.                  (08/07/04) *'
'*  - Down or Up list when press keys.     (09/07/04) *'
'*  - Drop down List.                      (09/07/04) *'
'*  - AddItem parameters.                  (09/07/04) *'
'*  - ChangeItem parameters.               (09/07/04) *'
'*  - SeparatorLine for Item.              (10/07/09) *'
'*  - Reorganize Code.                     (11/07/09) *'
'*  - ListIndex Property.                  (11/10/09) *'
'******************************************************'
'*       Errors corrected after the publication       *'
'*                   Version 1.0.2                    *'
'*                                                    *'
'*  - Reorganize Code.                     (14/07/09) *'
'*  - Additional Xp Appearance.            (14/07/09) *'
'*  - NiaWBSS Appearance.                  (14/07/04) *'
'*  - Style Arrow Appearance.              (15/07/04) *'
'*  - Ardent Appearance.                   (15/07/04) *'
'*  - Office 2000 Appearance.              (16/07/04) *'
'*  - Comments.                            (16/07/04) *'
'*  - Optimization code.                   (16/07/04) *'
'*  - Chocolate Appearance.                (16/07/04) *'
'*  - Limit the Height of Control.         (16/07/04) *'
'*  - Text in the control.                 (16/07/04) *'
'*  - Gral correction of the Appearances.  (17/07/04) *'
'******************************************************'
'*       Errors corrected after the publication       *'
'*                   Version 1.0.3                    *'
'*                                                    *'
'*  - Remove the reference to the parameter MouseIcon *'
'*    when RemoveItem uses.                (22/07/04) *'
'*  - Remove the reference to the parameter Separa-   *'
'*    torLine when RemoveItem uses.        (22/07/04) *'
'*  - Debug of the comments.               (22/07/04) *'
'*  - ScrollBar Position.                  (27/07/04) *'
'*  - Appearance Windows Xp theme.         (28/07/04) *'
'******************************************************'
'*       Errors corrected after the publication       *'
'*                   Version 1.0.4                    *'
'*                                                    *'
'*  - ItemFocus correction.                (04/08/04) *'
'*  - Clear function.                      (04/08/04) *'
'*  - Border Select.                       (04/08/04) *'
'*  - ScrollBar.                           (04/08/04) *'
'*  - Width Text.                          (07/08/04) *'
'*  - Move for the list.                   (11/08/04) *'
'*  - Office 2003 Appearance.              (15/08/04) *'
'*  - Sub DrawAppearance.                  (19/08/04) *'
'*  - Sub CreateText.                      (21/08/04) *'
'******************************************************'
'*                   Version 1.0.0                    *'
'*                                                    *'
'* Enhancements:  - Office Xp.             (13/06/04) *'
'*                - Win98.                 (13/06/04) *'
'*                - Control Properties.    (14/06/04) *'
'*                - Appearance.            (15/06/04) *'
'*                - WinXp.                 (16/06/04) *'
'*                - Office 2000.           (16/06/04) *'
'*                - Soft Style.            (16/06/04) *'
'*                - ItemTag Property.      (16/06/04) *'
'*                - JAVA.                  (17/06/04) *'
'*                - GradientV.             (18/06/04) *'
'*                - GradientH.             (18/06/04) *'
'*                - OrderList.             (18/06/04) *'
'*                - Color properties.      (19/06/04) *'
'*                - Explorer Bar.          (19/06/04) *'
'*                - Picture.               (19/06/04) *'
'*                - Mac.                   (21/06/04) *'
'*                - Special Border.        (22/06/04) *'
'*                - Rounded.               (22/06/04) *'
'*                - Search.                (23/06/04) *'
'*                - Style Arrow.           (25/06/04) *'
'*                - Light Blue.            (26/06/04) *'
'*                - KDE.                   (29/06/04) *'
'*                - Style Arrow.           (29/06/04) *'
'*                - NiaWBSS.               (30/06/04) *'
'*                - Rhombus.               (30/06/04) *'
'*                - Additional Xp.         (01/07/04) *'
'*                - Ardent.                (03/07/04) *'
'******************************************************'
'* Release date: Sunday, 04-07-2004.                  *'
'******************************************************'
'*          Enhancements after the publication        *'
'*                   Version 1.0.1                    *'
'*                                                    *'
'*  - Drop down when you press F4.         (08/07/04) *'
'*  - Press Enter hidden list.             (08/07/04) *'
'*  - MouseIcon property.                  (08/07/04) *'
'*  - MousePointer property.               (08/07/04) *'
'*  - Down or Up list when press keys.     (08/07/04) *'
'*  - Set ListIndex when the text change.  (09/07/04) *'
'*  - AutoCompleteWord property.           (09/07/04) *'
'*  - Const VK_LBUTTON.                    (09/07/04) *'
'*  - Const VK_RBUTTON.                    (09/07/04) *'
'*  - New comments.                        (09/07/04) *'
'*  - Add SeparatorLine for item.          (09/07/04) *'
'*  - Add MouseIcon for item.              (09/07/04) *'
'******************************************************'
'* Release date: Sunday, 11-07-2004                   *'
'******************************************************'
'*          Enhancements after the publication        *'
'*                   Version 1.0.2                    *'
'*                                                    *'
'*  - New Appearance: Chocolate.           (15/07/04) *'
'*  - New Appearance: Button Download.     (16/07/04) *'
'*  - Add controls the Usercontrol.        (16/07/04) *'
'*  - Added Windows XP Themed Style.       (17/07/04) *'
'******************************************************'
'* Release date: Sunday, 18-07-2004.                  *'
'******************************************************'
'*          Enhancements after the publication        *'
'*                   Version 1.0.3                    *'
'*                                                    *'
'*  - New Comments.                        (22/07/04) *'
'*  - ListPositionShow property.           (23/07/04) *'
'*  - Office 2003 Appearance.              (25/07/04) *'
'*  - Reorganize Office Appearance.        (27/07/04) *'
'*  - New Event TotalItems.                (27/07/04) *'
'*  - Office 2003 Appearance.              (04/08/04) *'
'******************************************************'
'* Release date: Sunday, 01-08-2004.                  *'
'******************************************************'
'*          Enhancements after the publication        *'
'*                   Version 1.0.4                    *'
'*                                                    *'
'*  - Now if the Font change works.        (04/08/04) *'
'*  - Now the temporary directory of                  *'
'*    Windows is used to manipulate                   *'
'*    the images.                          (04/08/04) *'
'*  - Parameter ShadowText.                (11/08/04) *'
'*  - Property ListGradient.               (13/08/04) *'
'*  - Function CalcTextWidth.              (15/08/04) *'
'******************************************************'
'* Release date: Sunday, 13-09-2004.                  *'
'******************************************************'
'* Note:     Comments, suggestions, doubts or bug     *'
'*           reports are wellcome to these e-mail     *'
'*           addresses:                               *'
'*                                                    *'
'*                  heri_05-hms@mixmail.com or        *'
'*                  hcammus@hotmail.com               *'
'*                                                    *'
'*        Please rate my work on this control.        *'
'*    That lives the Soccer and the América of Cali   *'
'*             Of Colombia for the world.             *'
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2004       *'
'******************************************************'
Option Explicit
 
 '****************************'
 '* English: Private Type.   *'
 '* Español: Tipos Privados. *'
 '****************************'
 Private Type GRADIENT_RECT
  UpperLeft   As Long
  LowerRight  As Long
 End Type
 
 Private Type RECT
  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long
 End Type
 
 Private Type RGB
  Red         As Integer
  Green       As Integer
  Blue        As Integer
 End Type
  
 Private Type POINTAPI
  X           As Long
  Y           As Long
 End Type

 Private Type Msg
  hWnd        As Long
  message     As Long
  wParam      As Long
  lParam      As Long
  time        As Long
  PT          As POINTAPI
 End Type
 
 '* English: Elements of the list.
 '* Español: Elementos de la lista.
 Private Type PropertyCombo
  Color         As OLE_COLOR   '* Color of Text.
  Enabled       As Boolean     '* Item Enabled or Disabled.
  Image         As StdPicture  '* Item image.
  Index         As Long        '* Index item.
  MouseIcon     As StdPicture  '* Set MouseIcon for each item.
  SeparatorLine As Boolean     '* Set SeparatorLine for each group that you consider necessary.
  Tag           As String      '* Extra Information only if is necessary.
  Text          As String      '* Text of the item.
  TextShadow    As Boolean     '* Shadow text item.
  ToolTipText   As String      '* ToolTipText for item.
 End Type
  
 Private Type TRIVERTEX
  X             As Long
  Y             As Long
  Red           As Integer
  Green         As Integer
  Blue          As Integer
  Alpha         As Integer
 End Type
  
 '*********************************************'
 '* English: Public Enum of Control.          *'
 '* Español: Enumeración Publica del control. *'
 '*********************************************'
 
 '* English: Enum for the alignment of the text of the list.
 '* Español: Enum para la alineación del texto de la lista.
 Public Enum AlignTextCombo
  AlignLeft = 0
  AlignRight = 1
  AlignCenter = 2
 End Enum
 
 '* English: Appearance Combo.
 '* Español: Apariencias del Combo.
 Public Enum ComboAppearance
  Office = 1             '* By fred.cpp & HACKPRO TM.
  Win98 = 2              '* By fred.cpp.
  WinXp = 3              '* By fred.cpp & HACKPRO TM.
  [Soft Style] = 4       '* By fred.cpp.
  KDE = 5                '* By HACKPRO TM.
  Mac = 6                '* By fred.cpp & HACKPRO TM.
  JAVA = 7               '* By fred.cpp.
  [Explorer Bar] = 8     '* By HACKPRO TM.
  Picture = 9            '* By HACKPRO TM.
  [Special Borde] = 10   '* By HACKPRO TM.
  Circular = 11          '* By HACKPRO TM.
  [GradientV] = 12       '* By HACKPRO TM.
  [GradientH] = 13       '* By HACKPRO TM.
  [Light Blue] = 14      '* By HACKPRO TM.
  [Style Arrow] = 15     '* By HACKPRO TM.
  [NiaWBSS] = 16         '* By HACKPRO TM.
  [Rhombus] = 17         '* By HACKPRO TM.
  [Additional Xp] = 18   '* By HACKPRO TM.
  [Ardent] = 19          '* By HACKPRO TM.
  [Chocolate] = 20       '* By HACKPRO TM.
  [Button Download] = 21 '* By HACKPRO TM.
 End Enum

 '* English: Type of Combo and behavior of the list.
 '* Español: Tipo de Combo y comportamiento de la lista.
 Public Enum ComboStyle
  [Dropdown Combo] = 0
  [Dropdown List] = 1
 End Enum
 
 '* English: Appearance standard style Office.
 '* Español: Apariencias estándares del estilo Office.
 Public Enum ComboOfficeAppearance
  [Office Xp] = 0       '* By HACKPRO TM.
  [Office 2000] = 1     '* By fred.cpp.
  [Office 2003] = 2     '* By HACKPRO TM.
 End Enum
 
 '* English: Appearance standard style Xp.
 '* Español: Apariencias estándares del estilo Xp.
 Public Enum ComboXpAppearance
  [Windows Themed] = 0  '* By fred.cpp
  Aqua = 1              '* By HACKPRO TM.
  [Olive Green] = 2     '* By HACKPRO TM.
  Silver = 3            '* By HACKPRO TM.
  TasBlue = 4           '* By HACKPRO TM.
  Gold = 5              '* By HACKPRO TM.
  Blue = 6              '* By HACKPRO TM.
  CustomXP = 7          '* By HACKPRO TM.
 End Enum
  
 '* English: Direction of like the list is shown.
 '* Español: Dirección de como se muestra la lista.
 Public Enum ListDirection
  [Show Down] = 0
  [Show Up] = 1
 End Enum
  
 '* English: Enum for the type of text comparison.
 '* Español: Enum para el tipo de comparación de texto.
 Public Enum StringCompare
  None = 0
  ExactWord = 1
  CompleteWord = 2
 End Enum
  
 '********************************'
 '* English: Private variables.  *'
 '* Español: Variables privadas. *'
 '********************************'
 Private BigText                 As String
 Private ControlEnabled          As Boolean
 Private cValor                  As Long
 Private CurrentS                As Long
 Private First                   As Integer
 Private FirstView               As Integer
 Private g_Font                  As StdFont
 Private HighlightedItem         As Long
 Private iFor                    As Long
 Private IndexItemNow            As Long
 Private IsMsg                   As Msg
 Private IsPicture               As Boolean
 Private ItemFocus               As Long
 Private KeyPos                  As Integer
 Private ListContents()          As PropertyCombo
 Private ListMaxL                As Long
 Private m_btnRect               As RECT
 Private m_bOver                 As Boolean
 Private m_LeaveMouse            As Boolean
 Private m_StateG                As Integer
 Private myAlignCombo            As AlignTextCombo
 Private myAppearanceCombo       As ComboAppearance
 Private myArrowColor            As OLE_COLOR
 Private myAutoSel               As Boolean
 Private myBackColor             As OLE_COLOR
 Private myListColor             As OLE_COLOR
 Private myDisabledColor         As OLE_COLOR
 Private myDisabledPictureUser   As StdPicture
 Private myFocusPictureUser      As StdPicture
 Private myGradientColor1        As OLE_COLOR
 Private myGradientColor2        As OLE_COLOR
 Private myHighLightBorderColor  As OLE_COLOR
 Private myHighLightColorText    As OLE_COLOR
 Private myHighLightPictureUser  As StdPicture
 Private myItemsShow             As Long
 Private myListGradient          As Boolean
 Private myListShown             As ListDirection
 Private myMouseIcon             As StdPicture
 Private myMousePointer          As MousePointerConstants
 Private myNormalBorderColor     As OLE_COLOR
 Private myNormalColorText       As OLE_COLOR
 Private myNormalPictureUser     As StdPicture
 Private myOfficeAppearance      As ComboOfficeAppearance
 Private mySelectBorderColor     As OLE_COLOR
 Private mySelectListBorderColor As OLE_COLOR
 Private mySelectListColor       As OLE_COLOR
 Private myShadowColorText       As OLE_COLOR
 Private myStyleCombo            As ComboStyle
 Private myText                  As String
 Private myXpAppearance          As ComboXpAppearance
 Private OrderListContents()     As PropertyCombo
 Private RGBColor                As RGB
 Private sumItem                 As Long
 Private tempBorderColor         As OLE_COLOR
 Private tmpC1                   As Long
 Private tmpC2                   As Long
 Private tmpC3                   As Long
 Private tmpColor                As Long
 
 '***************************************'
 '* English: Constant declares.         *'
 '* Español: Declaración de Constantes. *'
 '***************************************'
 Private Const BDR_RAISEDINNER = &H4
 Private Const BDR_SUNKENOUTER = &H2
 Private Const BF_RECT = (&H1 Or &H2 Or &H4 Or &H8)
 Private Const COLOR_BTNFACE = 15
 Private Const COLOR_BTNSHADOW = 16
 Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
 Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
 Private Const COLOR_GRAYTEXT As Long = 17
 Private Const COLOR_HIGHLIGHT As Long = 13
 Private Const COLOR_HOTLIGHT As Long = 26
 Private Const COLOR_INACTIVECAPTIONTEXT As Long = 19
 Private Const COLOR_WINDOW = 5
 Private Const defAppearanceCombo = 1
 Private Const defArrowColor = &HC56A31
 Private Const defDisabledColor = &H808080
 Private Const defGradientColor1 = &HDAB278
 Private Const defGradientColor2 = &HFFDD9E
 Private Const defHighLightBorderColor = &HC56A31
 Private Const defHighLightColorText = &HFFFFFF
 Private Const defNormalBorderColor = &HDEEDEF
 Private Const defNormalColorText = &HC56A31
 Private Const defListColor = &HFFFFFF
 Private Const defListShown = 0
 Private Const defOfficeAppearance = 0
 Private Const defSelectBorderColor = &HC56A31
 Private Const defSelectListBorderColor = &H6B2408
 Private Const defSelectListColor = &HC56A31
 Private Const defShadowColorText = &H80000015
 Private Const defStyleCombo = 0
 Private Const DSS_DISABLED = &H20
 Private Const DSS_NORMAL = &H0
 Private Const DST_BITMAP = &H3
 Private Const DST_COMPLEX = &H0
 Private Const DST_ICON = &H3
 Private Const DST_TEXT = &H2
 Private Const EDGE_RAISED = (&H1 Or &H4)
 Private Const EDGE_SUNKEN = (&H2 Or &H8)
 Private Const GRADIENT_FILL_RECT_H   As Long = &H0
 Private Const GRADIENT_FILL_RECT_V   As Long = &H1
 Private Const GWL_EXSTYLE = -20
 Private Const SWP_FRAMECHANGED = &H20
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOSIZE = &H1
 Private Const WS_EX_TOOLWINDOW = &H80
 Private Const Version As String = "SComboBox 1.0.3 By HACKPRO TM"
 Private Const VK_LBUTTON = &H1
 Private Const VK_RBUTTON = &H2
 Private Const WM_MOUSELEAVE As Integer = &H2A3
 Private Const WM_MOUSEWHEEL As Integer = &H20A
 
 '*******************************'
 '* English: Private WithEvents *'
 '* Español: Private WithEvents *'
 '*******************************'
 Private WithEvents scrollI       As VScrollBar
Attribute scrollI.VB_VarHelpID = -1
 Private WithEvents picTemp       As PictureBox
Attribute picTemp.VB_VarHelpID = -1
 
 '******************************'
 '* English: Public Events.    *'
 '* Español: Eventos Públicos. *'
 '******************************'
 Public Event SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
 Public Event TotalItems(ByVal ListCount As Long)
    
 '**********************************'
 '* English: Calls to the API's.   *'
 '* Español: Llamadas a los API's. *'
 '**********************************'
 Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
 Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DispatchMessage Lib "User32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
 Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
 Private Declare Function DrawState Lib "User32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal flags As Long) As Long
 Private Declare Function DrawStateString Lib "User32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lpString As String, ByVal cbStringLen As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal fuFlags As Long) As Long
 Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
 Private Declare Function FrameRect Lib "User32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FillRect Lib "User32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
 Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
 Private Declare Function GetMessage Lib "User32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
 Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
 Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
 Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
 Private Declare Function GetWindowDC Lib "User32" (ByVal hWnd As Long) As Long
 Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
 Private Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
 Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
 Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 Private Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
 Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
 Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
   
'***********************************************************'
'* English: Events of the controls and of the Usercontrol. *'
'* Español: Eventos de los controles y del Usercontrol.    *'
'***********************************************************'
Private Sub picList_Click()
 '* English: A Element has been selected or the control has been clicked
 '* Español: Establece el elemento donde se hizo clic.
On Error Resume Next
 If (ListContents(HighlightedItem + 1).Enabled = True) Then
  If (HighlightedItem + 1 >= ListCount1) Then HighlightedItem = HighlightedItem - 1
  ItemFocus = HighlightedItem + 1
  Call ListIndex1
  Text = ListContents(ItemFocus).Text
  Call DrawAppearance(myAppearanceCombo, 1)
  tmrFocus.Enabled = True
  RaiseEvent SelectionMade(ListContents(ListIndex1).Text, ItemFocus)
 End If
End Sub

Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)
 Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '* English: The mouse has been moved over the list
 '* Español: Mueve el mouse por la lista.
 FirstView = 1
 HighlightedItem = Int(Y / 20)
 If (ListCount < 1) Or (HighlightedItem + 1 + scrollI.Value > MaxListLength) Then Exit Sub
 IndexItemNow = HighlightedItem + 1
 If (ListContents(HighlightedItem + 1 + scrollI.Value).Enabled = True) Then
  HighlightedItem = HighlightedItem + scrollI.Value
  If (HighlightedItem + 1 > scrollI.Value + MaxListLength - 1) Then HighlightedItem = scrollI.Value + MaxListLength - 1
  If (HighlightedItem + 1 > ListCount1 - 1) Then HighlightedItem = ListCount1 - 1
  If (HighlightedItem + 1 < ListCount1) Then Call DrawList(scrollI.Value, NumberItemsToShow)
  picList.Refresh
 Else
  HighlightedItem = -1
 End If
 DoEvents
End Sub

Private Sub scrollI_Change()
 FirstView = 1
 HighlightedItem = Abs(IndexItemNow - 1)
 tmrFocus.Enabled = False
 Call DrawList(scrollI.Value, NumberItemsToShow)
End Sub

Private Sub scrollI_Scroll()
 scrollI_Change
End Sub

Private Sub tmrFocus_Timer()
 If (InFocusControl(UserControl.hWnd) = True) And (picList.Visible = False) Then
  If (m_bOver = False) Then Call DrawAppearance(myAppearanceCombo, 2)
  m_bOver = True
 ElseIf (m_bOver = True) And (picList.Visible = False) Then
  Call DrawAppearance(myAppearanceCombo, 1)
  tmrFocus.Enabled = False
  m_bOver = False
 End If
 If (Enabled = False) Then Call IsEnabled(ControlEnabled)
End Sub

Private Sub txtCombo_Change()
 Dim sItem As Long, iLen As Long, iStart As Long
 
On Error Resume Next
 iStart = txtCombo.SelStart
 If (myAutoSel = False) Then
  sItem = FindItemText(txtCombo.Text, 2)
  If (sItem > 0) Then
   If (ListContents(sItem).Enabled = True) Then
    ItemFocus = sItem
    IndexItemNow = sItem
    If (IndexItemNow > NumberItemsToShow) Then
     iLen = (NumberItemsToShow + IndexItemNow) - IndexItemNow
    Else
     iLen = IndexItemNow - (NumberItemsToShow + IndexItemNow)
    End If
    If (iLen > scrollI.Max) Then
     scrollI.Value = scrollI.Max
    ElseIf (iLen < 0) Then
     scrollI.Value = 0
    Else
     scrollI.Value = scrollI.Max
    End If
    Call scrollI_Change
   End If
  Else
   ItemFocus = -1
  End If
 ElseIf (KeyPos <> 67) And (KeyPos <> 46) Then
  sItem = FindItemText(txtCombo.Text)
  If (sItem > 0) Then
   iLen = Len(txtCombo.Text)
   txtCombo.Text = txtCombo.Text & Mid$(ListContents(sItem).Text, iLen + 1, Len(ListContents(sItem).Text))
   txtCombo.SelStart = iLen
   txtCombo.SelLength = Len(txtCombo.Text)
   sItem = FindItemText(txtCombo.Text, 2)
   If (sItem > 0) Then
    If (ListContents(sItem).Enabled = True) Then
     ItemFocus = sItem
     IndexItemNow = sItem
    End If
   Else
    ItemFocus = -1
   End If
  Else
   ItemFocus = -1
  End If
  Call IsEnabled(ControlEnabled)
 Else
  ItemFocus = FindItemText(txtCombo.Text, 2)
  Call IsEnabled(ControlEnabled)
  txtCombo.SelStart = iStart
 End If
End Sub

Private Sub txtCombo_GotFocus()
 txtCombo.SelStart = 0
 txtCombo.SelLength = Len(txtCombo.Text)
End Sub

Private Sub txtCombo_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = 115) Then Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCombo_KeyUp(KeyCode As Integer, Shift As Integer)
 KeyPos = KeyCode
 If (KeyCode = 115) Then Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (picList.Visible = False) Then tmrFocus.Enabled = True
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
 If (AppearanceCombo = 18) Then Call IsEnabled(ControlEnabled)
End Sub

Private Sub UserControl_ExitFocus()
 Call IsEnabled(ControlEnabled)
 tmrFocus.Enabled = True
End Sub

Private Sub UserControl_InitProperties()
 '* English: Setup properties values.
 '* Español: Establece propiedades iniciales.
 ControlEnabled = True
 ItemFocus = -1
 IsPicture = False
 ListIndex = -1
 ListMaxL = 10
 myListShown = 0
 myAutoSel = False
 myAppearanceCombo = defAppearanceCombo
 myArrowColor = defArrowColor
 myBackColor = defListColor
 myDisabledColor = defDisabledColor
 myGradientColor1 = defGradientColor1
 myGradientColor2 = defGradientColor2
 myHighLightBorderColor = defHighLightBorderColor
 myHighLightColorText = defHighLightColorText
 myItemsShow = 7
 myListColor = defListColor
 myListGradient = False
 myNormalBorderColor = defNormalBorderColor
 myNormalColorText = defNormalColorText
 myOfficeAppearance = defOfficeAppearance
 mySelectBorderColor = defSelectBorderColor
 mySelectListBorderColor = defSelectListBorderColor
 mySelectListColor = defSelectListColor
 myShadowColorText = defShadowColorText
 myStyleCombo = defStyleCombo
 myText = Ambient.DisplayName
 Text = myText
 myXpAppearance = 1
 Set g_Font = Ambient.Font
 sumItem = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13 '* Enter.
   If (picList.Visible = True) Then Call UserControl_MouseDown(0, 0, 0, 0)
  Case 33 '* PageDown.
   If (IndexItemNow > NumberItemsToShow) Then
    IndexItemNow = IndexItemNow - NumberItemsToShow - 1
    If (IndexItemNow < 0) Then IndexItemNow = 1
    If (scrollI.Value - NumberItemsToShow - 1 > 0) Then scrollI.Value = scrollI.Value - NumberItemsToShow - 1 Else scrollI.Value = 0
   Else
    IndexItemNow = 1
    scrollI.Value = 0
   End If
   scrollI_Change
  Case 34 '* PageUp.
   If (IndexItemNow < sumItem) Then
    IndexItemNow = IndexItemNow + NumberItemsToShow - 1
    If (IndexItemNow > sumItem) Then IndexItemNow = sumItem
    If (scrollI.Value + NumberItemsToShow - 1 < scrollI.Max) Then scrollI.Value = scrollI.Value + NumberItemsToShow - 1 Else scrollI.Value = scrollI.Max
   Else
    IndexItemNow = sumItem
    scrollI.Value = scrollI.Max
   End If
   scrollI_Change
  Case 35 '* End.
   IndexItemNow = sumItem
   scrollI.Value = scrollI.Max
   scrollI_Change
  Case 36 '* Start.
   IndexItemNow = 1
   scrollI.Value = 0
   scrollI_Change
  Case 38 '* Up arrow.
   If (IndexItemNow > 0) Then
    IndexItemNow = IndexItemNow - 1
    If (scrollI.Value > 0) And (IndexItemNow - NumberItemsToShow < NumberItemsToShow) Then scrollI.Value = scrollI.Value - 1
    scrollI_Change
   End If
  Case 40 '* Down arrow.
   If (IndexItemNow < sumItem) Then
    IndexItemNow = IndexItemNow + 1
    If (scrollI.Value < scrollI.Max) And (IndexItemNow > NumberItemsToShow) Then scrollI.Value = scrollI.Value + 1
    scrollI_Change
   End If
  Case 115 '* Key F4.
   Call UserControl_MouseDown(1, 0, 0, 0)
 End Select
End Sub

Private Sub UserControl_LostFocus()
 Call UserControl_ExitFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim oRect As RECT
 
 '* English: Show or hide the list.
 '* Español: Muestra la lista ó la oculta.
 If (Button = vbLeftButton) And (picList.Visible = False) Then
  First = 1
  HighlightedItem = -1
  IndexItemNow = ListIndex
  scrollI.Max = IIf(MaxListLength - NumberItemsToShow < 0, 0, MaxListLength - NumberItemsToShow)
  If (ListCount > NumberItemsToShow) And (ItemFocus > 1) And (ItemFocus < scrollI.Max) Then
   scrollI.Value = IIf(NumberItemsToShow < ItemFocus - 1, Abs(scrollI.Max - NumberItemsToShow), 1)
  ElseIf (ItemFocus > scrollI.Max) Then
   scrollI.Value = scrollI.Max
  Else
   scrollI.Value = 0
  End If
  FirstView = 0
  tmrFocus.Enabled = False
  If (ListCount > NumberItemsToShow) Then
   picList.Height = NumberItemsToShow * 300
  ElseIf (ListCount > 0) Then
   picList.Height = ListCount * 300
  Else
   picList.Height = 240
  End If
  Call GetWindowRect(hWnd, oRect)
  If (myListShown = 1) Then
   '* The list is shown up.
   Call picList.Move(oRect.Left * Screen.TwipsPerPixelX, (oRect.Bottom * Screen.TwipsPerPixelY) - (picList.Height + UserControl.Height + 21))
  Else
   '* The list is shown down.
   Call picList.Move(oRect.Left * Screen.TwipsPerPixelX, oRect.Bottom * Screen.TwipsPerPixelY + 21)
  End If
  Call SetWindowPos(picList.hWnd, -1, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
  Call scrollI.Move(picList.ScaleWidth - 18, 1)
  If (NumberItemsToShow < MaxListLength) Then
   scrollI.Height = picList.ScaleHeight - 2
   scrollI.Visible = True
  Else
   scrollI.Visible = False
  End If
  Call SetWindowPos(scrollI.hWnd, -1, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
  Call DrawAppearance(myAppearanceCombo, 3)
  If (myAppearanceCombo = 2) Or ((myAppearanceCombo = 3) And ((myXpAppearance = 7) Or (myXpAppearance = 0))) Then
   Call Espera(0.09)
   Call DrawAppearance(myAppearanceCombo, 1)
  End If
  ItemFocus = FindItemText(myText, 2)
  Call DrawList(scrollI.Value, NumberItemsToShow)
  picList.Visible = True
 Else
  Call DrawAppearance(myAppearanceCombo, 2)
  picList.Visible = False
  First = 0
 End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (picList.Visible = False) Then tmrFocus.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (picList.Visible = False) Then tmrFocus.Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Alignment = PropBag.ReadProperty("Alignment", 0)
 AppearanceCombo = PropBag.ReadProperty("AppearanceCombo", defAppearanceCombo)
 ArrowColor = PropBag.ReadProperty("ArrowColor", defArrowColor)
 AutoCompleteWord = PropBag.ReadProperty("AutoCompleteWord", False)
 BackColor = PropBag.ReadProperty("BackColor", defListColor)
 Call ControlsSubClasing
 DisabledColor = PropBag.ReadProperty("DisabledColor", defDisabledColor)
 Set DisabledPictureUser = PropBag.ReadProperty("DisabledPictureUser", Nothing)
 Enabled = PropBag.ReadProperty("Enabled", True)
 GradientColor1 = PropBag.ReadProperty("GradientColor1", defGradientColor1)
 GradientColor2 = PropBag.ReadProperty("GradientColor2", defGradientColor2)
 Set FocusPictureUser = PropBag.ReadProperty("FocusPictureUser", Nothing)
 Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
 HighLightBorderColor = PropBag.ReadProperty("HighLightBorderColor", defHighLightBorderColor)
 HighLightColorText = PropBag.ReadProperty("HighLightColorText", defHighLightColorText)
 Set HighLightPictureUser = PropBag.ReadProperty("HighLightPictureUser", Nothing)
 ListColor = PropBag.ReadProperty("ListColor", defListColor)
 ListGradient = PropBag.ReadProperty("ListGradient", False)
 ListPositionShow = PropBag.ReadProperty("ListPositionShow", defListShown)
 MaxListLength = PropBag.ReadProperty("MaxListLength", "10")
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 MousePointer = PropBag.ReadProperty("MousePointer", 0)
 NormalBorderColor = PropBag.ReadProperty("NormalBorderColor", defNormalBorderColor)
 NormalColorText = PropBag.ReadProperty("NormalColorText", defNormalColorText)
 Set NormalPictureUser = PropBag.ReadProperty("NormalPictureUser", Nothing)
 NumberItemsToShow = PropBag.ReadProperty("NumberItemsToShow", "7")
 OfficeAppearance = PropBag.ReadProperty("OfficeAppearance", defOfficeAppearance)
 SelectBorderColor = PropBag.ReadProperty("SelectBorderColor", defSelectBorderColor)
 SelectListBorderColor = PropBag.ReadProperty("SelectListBorderColor", defSelectListBorderColor)
 SelectListColor = PropBag.ReadProperty("SelectListColor", defSelectListColor)
 ShadowColorText = PropBag.ReadProperty("ShadowColorText", defShadowColorText)
 Style = PropBag.ReadProperty("Style", defStyleCombo)
 Text = PropBag.ReadProperty("Text", Ambient.DisplayName)
 XpAppearance = PropBag.ReadProperty("XpAppearance", 1)
End Sub

Private Sub UserControl_Resize()
 Call IsEnabled(ControlEnabled)
 Call IsEnabled(ControlEnabled)
End Sub

Private Sub UserControl_Show()
 Dim lResult As Long
 
On Error Resume Next
 lResult = GetWindowLong(picList.hWnd, GWL_EXSTYLE)
 Call SetWindowLong(picList.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
 Call SetWindowPos(picList.hWnd, picList.hWnd, 0, 0, 0, 0, 39)
 Call SetWindowLong(picList.hWnd, -8, Parent.hWnd)
 Call SetParent(picList.hWnd, 0)
 lResult = GetWindowLong(scrollI.hWnd, GWL_EXSTYLE)
 Call SetWindowLong(scrollI.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
 Call SetWindowPos(scrollI.hWnd, scrollI.hWnd, 0, 0, 0, 0, 39)
 Call SetWindowLong(scrollI.hWnd, -8, Parent.hWnd)
 Call SetParent(scrollI.hWnd, picList.hWnd)
 If (IsPicture = False) Then txtCombo.Left = 8
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
 Erase ListContents
 Set picTemp = Nothing
 Set scrollI = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Alignment", myAlignCombo, 0)
 Call PropBag.WriteProperty("AppearanceCombo", myAppearanceCombo, defAppearanceCombo)
 Call PropBag.WriteProperty("ArrowColor", myArrowColor, defArrowColor)
 Call PropBag.WriteProperty("AutoCompleteWord", myAutoSel, False)
 Call PropBag.WriteProperty("BackColor", myBackColor, defListColor)
 Call PropBag.WriteProperty("DisabledColor", myDisabledColor, defDisabledColor)
 Call PropBag.WriteProperty("DisabledPictureUser", myDisabledPictureUser, Nothing)
 Call PropBag.WriteProperty("Enabled", Enabled, True)
 Call PropBag.WriteProperty("FocusPictureUser", myFocusPictureUser, Nothing)
 Call PropBag.WriteProperty("Font", g_Font, Ambient.Font)
 Call PropBag.WriteProperty("GradientColor1", myGradientColor1, defGradientColor1)
 Call PropBag.WriteProperty("GradientColor2", myGradientColor2, defGradientColor2)
 Call PropBag.WriteProperty("HighLightBorderColor", myHighLightBorderColor, defHighLightBorderColor)
 Call PropBag.WriteProperty("HighLightColorText", myHighLightColorText, defHighLightColorText)
 Call PropBag.WriteProperty("HighLightPictureUser", myHighLightPictureUser, Nothing)
 Call PropBag.WriteProperty("ListColor", myListColor, defListColor)
 Call PropBag.WriteProperty("ListGradient", myListGradient, False)
 Call PropBag.WriteProperty("ListPositionShow", myListShown, defListShown)
 Call PropBag.WriteProperty("MaxListLength", ListMaxL, "10")
 Call PropBag.WriteProperty("MouseIcon", myMouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", myMousePointer, 0)
 Call PropBag.WriteProperty("NormalBorderColor", myNormalBorderColor, defNormalBorderColor)
 Call PropBag.WriteProperty("NormalColorText", myNormalColorText, defNormalColorText)
 Call PropBag.WriteProperty("NormalPictureUser", myNormalPictureUser, Nothing)
 Call PropBag.WriteProperty("NumberItemsToShow", myItemsShow, "7")
 Call PropBag.WriteProperty("OfficeAppearance", myOfficeAppearance, defOfficeAppearance)
 Call PropBag.WriteProperty("SelectBorderColor", mySelectBorderColor, defSelectBorderColor)
 Call PropBag.WriteProperty("SelectListBorderColor", mySelectListBorderColor, defSelectListBorderColor)
 Call PropBag.WriteProperty("SelectListColor", mySelectListColor, defSelectListColor)
 Call PropBag.WriteProperty("ShadowColorText", myShadowColorText, defShadowColorText)
 Call PropBag.WriteProperty("Style", myStyleCombo, defStyleCombo)
 Call PropBag.WriteProperty("Text", myText, Ambient.DisplayName)
 Call PropBag.WriteProperty("XpAppearance", myXpAppearance, 1)
End Sub

'*******************************************'
'* English: Properties of the Usercontrol. *'
'* Español: Propiedades del Usercontrol.   *'
'*******************************************'
Public Property Get Alignment() As AlignTextCombo
Attribute Alignment.VB_Description = "Sets/Gets alignment of the text in the list."
 '* English: Sets/Gets alignment of the text in the list.
 '* Español: Devuelve o establece la alineación del texto en la lista.
 Alignment = myAlignCombo
End Property

Public Property Let Alignment(ByVal New_Align As AlignTextCombo)
 myAlignCombo = New_Align
 Call PropertyChanged("Alignment")
 Refresh
End Property

Public Property Get AppearanceCombo() As ComboAppearance
Attribute AppearanceCombo.VB_Description = "Sets/Gets the style of the Combo."
 '* English: Sets/Gets the style of the Combo.
 '* Español: Devuelve o establece el estilo del Combo.
 AppearanceCombo = myAppearanceCombo
End Property

Public Property Let AppearanceCombo(ByVal New_Style As ComboAppearance)
 myAppearanceCombo = IIf(New_Style <= 0, 1, New_Style)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("AppearanceCombo")
 Refresh
End Property

Public Property Get ArrowColor() As OLE_COLOR
Attribute ArrowColor.VB_Description = "Sets/Gets the color of the arrow."
 '* English: Sets/Gets the color of the arrow.
 '* Español: Devuelve o establece el color de la flecha.
 ArrowColor = myArrowColor
End Property

Public Property Let ArrowColor(ByVal New_Color As OLE_COLOR)
 myArrowColor = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("ArrowColor")
 Refresh
End Property

Public Property Get AutoCompleteWord() As Boolean
Attribute AutoCompleteWord.VB_Description = "Sets/Gets complete the word with a similar element of the list."
 '* English: Sets/Gets complete the word with a similar element of the list.
 '* Español: Devuelve o establece si se completa la palabra con un elemento similar de la lista.
 AutoCompleteWord = myAutoSel
End Property
'* Note: When this property this active one and the list _
         is shown, it is not tried to locate the element _
         in the list to make quicker the search of the _
         text to complete.
'* Nota: Cuando esta propiedad este activa y la lista se _
         muestre, no se intentara ubicar el elemento en la _
         lista para hacer más rápido la búsqueda del texto _
         a completar.

Public Property Let AutoCompleteWord(ByVal New_Value As Boolean)
 myAutoSel = New_Value
 Call PropertyChanged("AutoCompleteWord")
 Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets/Gets the color of the Usercontrol."
 '* English: Sets/Gets the color of the Usercontrol.
 '* Español: Devuelve o establece el color del Usercontrol.
 BackColor = myBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
 myBackColor = ConvertSystemColor(GetLngColor(New_Color))
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("BackColor")
 Refresh
End Property

Public Property Get DisabledColor() As OLE_COLOR
Attribute DisabledColor.VB_Description = "Sets/Gets the color of the disabled text."
 '* English: Sets/Gets the color of the disabled text.
 '* Español: Devuelve o establece el color del texto deshabilitado.
 DisabledColor = ShiftColorOXP(myDisabledColor, 94)
End Property

Public Property Let DisabledColor(ByVal New_Color As OLE_COLOR)
 myDisabledColor = ConvertSystemColor(GetLngColor(New_Color))
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("DisabledColor")
 Refresh
End Property

Public Property Get DisabledPictureUser() As StdPicture
Attribute DisabledPictureUser.VB_Description = "Sets/Gets an image like topic of the Combo when the Object is not enabled."
 '* English: Sets/Gets an image like topic of the Combo when the Object is not enabled.
 '* Español: Devuelve o establece una imagen como tema del combo cuando el Objeto este inactivo.
 Set DisabledPictureUser = myDisabledPictureUser
End Property

Public Property Set DisabledPictureUser(ByVal New_Picture As StdPicture)
 Set myDisabledPictureUser = New_Picture
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("DisabledPictureUser")
 Refresh
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/Gets the Enabled property of the control."
 '* English: Sets/Gets the Enabled property of the control.
 '* Español: Devuelve o establece si el Usercontrol esta habilitado ó deshabilitado.
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 UserControl.Enabled = New_Enabled
 ControlEnabled = New_Enabled
 Call IsEnabled(New_Enabled)
 Call PropertyChanged("Enabled")
End Property

Public Property Get FocusPictureUser() As StdPicture
Attribute FocusPictureUser.VB_Description = "Sets/Gets the image like topic of the Combo when It has the focus."
 '* English: Sets/Gets the image like topic of the Combo when It has the focus.
 '* Español: Devuelve o establece una imagen como tema del combo cuando se tiene el enfoque.
 Set FocusPictureUser = myFocusPictureUser
End Property

Public Property Set FocusPictureUser(ByVal New_Picture As StdPicture)
 Set myFocusPictureUser = New_Picture
 Call PropertyChanged("FocusPictureUser")
 Refresh
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Set the Font of the control."
 '* English: Sets/Gets the Font of the control.
 '* Español: Devuelve o establece el tipo de fuente del texto.
 Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
On Error Resume Next
 With g_Font
  .Name = New_Font.Name
  .Size = New_Font.Size
  .Bold = New_Font.Bold
  .Italic = New_Font.Italic
  .Underline = New_Font.Underline
  .Strikethrough = New_Font.Strikethrough
 End With
 txtCombo.Font = New_Font
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("Font")
 Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
Attribute GradientColor1.VB_Description = "Sets/Gets the color First gradient color."
 '* English: Sets/Gets the color First gradient color.
 '* Español: Devuelve o establece el color Gradient 1.
 GradientColor1 = myGradientColor1
End Property

Public Property Let GradientColor1(ByVal New_Color As OLE_COLOR)
 myGradientColor1 = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("GradientColor1")
 Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
Attribute GradientColor2.VB_Description = "Sets/Gets the Second gradient color."
 '* English: Sets/Gets the Second gradient color.
 '* Español: Devuelve o establece el color Gradient 2.
 GradientColor2 = myGradientColor2
End Property

Public Property Let GradientColor2(ByVal New_Color As OLE_COLOR)
 myGradientColor2 = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("GradientColor2")
 Refresh
End Property

Public Property Get HighLightBorderColor() As OLE_COLOR
Attribute HighLightBorderColor.VB_Description = "Sets/Gets the color of the border of the control when the the control is highlighted."
 '* English: Sets/Gets the color of the border of the control when the the control is highlighted.
 '* Español: Devuelve o establece el color del borde del control cuando el pasa sobre él.
 HighLightBorderColor = myHighLightBorderColor
End Property

Public Property Let HighLightBorderColor(ByVal New_Color As OLE_COLOR)
 myHighLightBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("HighLightBorderColor")
 Refresh
End Property

Public Property Get HighLightColorText() As OLE_COLOR
Attribute HighLightColorText.VB_Description = "Sets/Gets the color of the selection of the text."
 '* English: Sets/Gets the color of the selection of the text.
 '* Español: Devuelve o establece el color de selección del texto.
 HighLightColorText = myHighLightColorText
End Property

Public Property Let HighLightColorText(ByVal New_Color As OLE_COLOR)
 myHighLightColorText = ConvertSystemColor(New_Color)
 Call PropertyChanged("HighLightColorText")
 Refresh
End Property

Public Property Get HighLightPictureUser() As StdPicture
Attribute HighLightPictureUser.VB_Description = "Sets/Gets an image like topic of the Combo when the mouse is over the control."
 '* English: Sets/Gets an image like topic of the Combo when the mouse is over the control.
 '* Español: Devuelve o establece una imagen como tema del combo cuando el mouse pasa por el Objeto.
 Set HighLightPictureUser = myHighLightPictureUser
End Property

Public Property Set HighLightPictureUser(ByVal New_Picture As StdPicture)
 Set myHighLightPictureUser = New_Picture
 Call PropertyChanged("HighLightPictureUser")
 Refresh
End Property

Public Property Get ItemTag(ByVal ListIndex As Long) As String
Attribute ItemTag.VB_Description = "Returns the tag of a specified item."
 '* English: Returns the tag of a specified item.
 '* Español: Selecciona el tag de Item.
 ItemTag = ""
On Error GoTo myErr:
 ItemTag = ListContents(ListIndex).Tag
 Exit Property
myErr:
 ItemTag = ""
End Property

Public Property Get ListColor() As OLE_COLOR
Attribute ListColor.VB_Description = "Sets/Gets the color of the List."
 '* English: Sets/Gets the color of the List.
 '* Español: Devuelve o establece el color de la lista.
 ListColor = myListColor
End Property

Public Property Let ListColor(ByVal New_Color As OLE_COLOR)
 myListColor = ConvertSystemColor(New_Color)
 picList.BackColor = myListColor
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("ListColor")
 Refresh
End Property

Public Property Get ListCount() As Long
Attribute ListCount.VB_Description = "Returns the number of elements in the list."
 '* English: Returns the number of elements in the list.
 '* Español: Devuelve o establece el número de elementos de la lista.
 ListCount = ListCount1 - 1
End Property

Public Property Get ListGradient() As Boolean
Attribute ListGradient.VB_Description = "Sets/Gets the list in degraded form."
 '* English: Sets/Gets the list in degraded form.
 '* Español: Devuelve o establece si la lista se muestra en forma degradada.
 ListGradient = myListGradient
End Property

Public Property Let ListGradient(ByVal New_Gradient As Boolean)
 myListGradient = New_Gradient
 Call PropertyChanged("ListGradient")
 Refresh
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "Sets/Gets the selected item."
Attribute ListIndex.VB_MemberFlags = "400"
 '* English: Sets/Gets the selected item.
 '* Español: Devuelve o establece el item actual seleccionado.
 ListIndex = ListIndex1
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)
 Call ListIndex1(New_ListIndex)
End Property

Public Property Get ListPositionShow() As ListDirection
Attribute ListPositionShow.VB_Description = "Sets/Gets If the list is shown up or down."
 '* English: Sets/Gets If the list is shown up or down.
 '* Español: Devuelve o establece si la lista se muestra hacia arriba ó hacia abajo.
 ListPositionShow = myListShown
End Property

Public Property Let ListPositionShow(ByVal New_Position As ListDirection)
 myListShown = New_Position
 Call PropertyChanged("ListPositionShow")
 Refresh
End Property

Public Property Get MaxListLength() As Long
Attribute MaxListLength.VB_Description = "Sets/Gets the maximum size of the list."
Attribute MaxListLength.VB_MemberFlags = "400"
 '* English: Sets/Gets the maximum size of the list.
 '* Español: Devuelve o establece el tamaño máximo de la lista.
 MaxListLength = IIf(ListMaxL < 0, ListCount, ListMaxL)
End Property

Public Property Let MaxListLength(ByVal ListMax As Long)
 If (ListMax > 0) And (ListMax < ListCount1) Then
  ListMaxL = ListMax
 Else
  ListMaxL = ListCount
 End If
 Call PropertyChanged("MaxListLength")
 Refresh
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
 '* English: Sets a custom mouse icon.
 '* Español: Establece un icono escogido por el usuario.
 Set MouseIcon = myMouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
 Set myMouseIcon = New_MouseIcon
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Sets/Gets the type of mouse pointer displayed when over part of an object."
 '* English: Sets/Gets the type of mouse pointer displayed when over part of an object.
 '* Español: Devuelve o establece el tipo de puntero a mostrar cuando el mouse pase sobre el objeto.
 MousePointer = myMousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
 myMousePointer = New_MousePointer
End Property

Public Property Get NewIndex() As Long
Attribute NewIndex.VB_Description = "Sets/Gets the last Item added."
 '* English: Sets/Gets the last Item added.
 '* Español: Devuelve o establece el último item agregado.
 If (sumItem <= 0) Then NewIndex = -1 Else NewIndex = sumItem
End Property

Public Property Get NormalBorderColor() As OLE_COLOR
Attribute NormalBorderColor.VB_Description = "Sets/Gets the normal border color of the control."
 '* English: Sets/Gets the normal border color of the control.
 '* Español: Devuelve o establece el color normal del borde del control.
 NormalBorderColor = myNormalBorderColor
End Property

Public Property Let NormalBorderColor(ByVal New_Color As OLE_COLOR)
 myNormalBorderColor = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("NormalBorderColor")
 Refresh
End Property

Public Property Let NormalColorText(ByVal New_Color As OLE_COLOR)
Attribute NormalColorText.VB_Description = "Sets/Gets the normal text color in the control."
 myNormalColorText = ConvertSystemColor(New_Color)
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("NormalColorText")
 Refresh
End Property

Public Property Get NormalColorText() As OLE_COLOR
 '* English: Sets/Gets the normal text color in the control.
 '* Español: Devuelve o establece el color del texto normal.
 NormalColorText = myNormalColorText
End Property

Public Property Get NormalPictureUser() As StdPicture
Attribute NormalPictureUser.VB_Description = "Sets/Gets an image like topic of the Combo in normal state."
 '* English: Sets/Gets an image like topic of the Combo in normal state.
 '* Español: Devuelve o establece una imagen como tema del combo en estado normal.
 Set NormalPictureUser = myNormalPictureUser
End Property

Public Property Set NormalPictureUser(ByVal New_Picture As StdPicture)
 Set myNormalPictureUser = New_Picture
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("NormalPictureUser")
 Refresh
End Property

Public Property Get NumberItemsToShow() As Long
Attribute NumberItemsToShow.VB_Description = "Sets/Gets the number of items to show per time."
Attribute NumberItemsToShow.VB_MemberFlags = "400"
 '* English: Sets/Gets the number of items to show per time.
 '* Español: Devuelve o establece el número de items a mostrar por vez.
 If (myItemsShow < 0) Then myItemsShow = IIf(MaxListLength > 8, 7, MaxListLength)
 NumberItemsToShow = myItemsShow
End Property

Public Property Let NumberItemsToShow(ByVal ItemsShow As Long)
 If (ItemsShow <= 1) Or (ItemsShow >= MaxListLength) Then
  myItemsShow = IIf(MaxListLength > 8, MaxListLength - 8, ListCount)
 Else
  myItemsShow = ItemsShow
 End If
 Call PropertyChanged("NumberItemsToShow")
 Refresh
End Property

Public Property Get OfficeAppearance() As ComboOfficeAppearance
Attribute OfficeAppearance.VB_Description = "Sets/Gets the office apperance."
 '* English: Sets/Gets the office apperance.
 '* Español: Devuelve o establece la apariencia de Office.
 OfficeAppearance = myOfficeAppearance
End Property

Public Property Let OfficeAppearance(ByVal New_Apperance As ComboOfficeAppearance)
 myOfficeAppearance = New_Apperance
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("OfficeAppearance")
 Refresh
End Property

Public Property Get SelectBorderColor() As OLE_COLOR
Attribute SelectBorderColor.VB_Description = "Sets/Gets the color of the border of the control when It has the focus."
 '* English: Sets/Gets the color of the border of the control when It has the focus.
 '* Español: Devuelve o establece el color del borde del control cuando el tenga el enfoque.
 SelectBorderColor = mySelectBorderColor
End Property

Public Property Let SelectBorderColor(ByVal New_Color As OLE_COLOR)
 mySelectBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectBorderColor")
 Refresh
End Property

Public Property Get SelectListBorderColor() As OLE_COLOR
Attribute SelectListBorderColor.VB_Description = "Sets/Gets the border color of the item selected in the list."
 '* English: Sets/Gets the border color of the item selected in the list.
 '* Español: Devuelve o establece el color del borde del item seleccionado en la lista.
 SelectListBorderColor = mySelectListBorderColor
End Property

Public Property Let SelectListBorderColor(ByVal New_Color As OLE_COLOR)
 mySelectListBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectListBorderColor")
 Refresh
End Property

Public Property Get SelectListColor() As OLE_COLOR
Attribute SelectListColor.VB_Description = "Sets/Gets the color of the item selected in the list."
 '* English: Sets/Gets the color of the item selected in the list.
 '* Español: Devuelve o establece el color del item seleccionado en la lista.
 SelectListColor = mySelectListColor
End Property

Public Property Let SelectListColor(ByVal New_Color As OLE_COLOR)
 mySelectListColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("SelectListColor")
 Refresh
End Property

Public Property Get ShadowColorText() As OLE_COLOR
Attribute ShadowColorText.VB_Description = "Sets/Gets the text color of the shadow."
 '* English: Sets/Gets the text color of the shadow.
 '* Español: Devuelve o establece el color de la sombra del texto.
 ShadowColorText = myShadowColorText
End Property

Public Property Let ShadowColorText(ByVal New_Color As OLE_COLOR)
 myShadowColorText = ConvertSystemColor(New_Color)
 Call PropertyChanged("ShadowColorText")
 Refresh
End Property

Public Property Get Style() As ComboStyle
Attribute Style.VB_Description = "Sets/Gets the style of the Combo."
 '* English: Sets/Gets the style of the Combo.
 '* Español: Devuelve o establece el estilo del Combo.
 Style = myStyleCombo
End Property

Public Property Let Style(ByVal New_Style As ComboStyle)
 myStyleCombo = New_Style
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("Style")
 Refresh
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Sets/Gets the text of the selected item."
 '* English: Sets/Gets the text of the selected item.
 '* Español: Devuelve o establece el texto del item seleccionado.
 Text = myText
End Property

Public Property Let Text(ByVal NewText As String)
 myText = NewText
 txtCombo.Text = myText
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("Text")
End Property

Public Property Get XpAppearance() As ComboXpAppearance
Attribute XpAppearance.VB_Description = "Sets the appearance in Xp Mode."
 '* English: Sets the appearance in Xp Mode.
 '* Español: Establece la apariencia en modo Xp.
 XpAppearance = myXpAppearance
End Property

Public Property Let XpAppearance(ByVal New_Style As ComboXpAppearance)
 myXpAppearance = New_Style
 Call IsEnabled(ControlEnabled)
 Call PropertyChanged("XpAppearance")
End Property

'********************************************************'
'* English: Subs and Functions of the Usercontrol.      *'
'* Español: Procedimientos y Funciones del Usercontrol. *'
'********************************************************'
Public Sub AddItem(ByVal Item As String, Optional ByVal ColorTextItem As OLE_COLOR = &HC56A31, Optional ByVal ImageItem As StdPicture = Nothing, Optional ByVal EnabledItem As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal IndexItem As Long = -1, Optional ByVal ItemTag As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False, Optional ByVal TextShadow As Boolean = False)
Attribute AddItem.VB_Description = "Add a new item to the list."
 '* English: Add a new item to the list.
 '* Español: Agrega un nuevo item a la lista.
 If (Item = "") Then Item = " "
 sumItem = sumItem + 1
 ReDim Preserve ListContents(sumItem)
 If (IndexItem > 0) And (IndexItem < sumItem) And (NoFindIndex(IndexItem) = False) Then
  ListContents(sumItem).Index = IndexItem
 Else
  ListContents(sumItem).Index = sumItem
 End If
 ListContents(sumItem).Color = IIf(EnabledItem = True, ColorTextItem, DisabledColor)
 If (Len(Item) > Len(BigText)) Then BigText = Item
 ListContents(sumItem).Text = Item
 ListContents(sumItem).TextShadow = TextShadow
 ListContents(sumItem).Enabled = EnabledItem
 ListContents(sumItem).Index = IndexItem
 ListContents(sumItem).ToolTipText = ToolTipTextItem
 ListContents(sumItem).Tag = ItemTag
 Set ListContents(sumItem).MouseIcon = MouseIcon
 ListContents(sumItem).SeparatorLine = SeparatorLine
 Set ListContents(sumItem).Image = ImageItem
 If Not (ImageItem Is Nothing) Then IsPicture = True
 MaxListLength = sumItem
 RaiseEvent TotalItems(sumItem)
End Sub

Private Sub APIFillRect(ByVal hDC As Long, ByRef RC As RECT, ByVal Color As Long)
 Dim NewBrush As Long
 
 '* English: The FillRect function fills a rectangle by using the specified brush. _
             This function includes the left and top borders, but excludes the right _
             and bottom borders of the rectangle.
 '* Español: Pinta el rectángulo de un objeto.
 NewBrush& = CreateSolidBrush(Color&)
 Call FillRect(hDC&, RC, NewBrush&)
 Call DeleteObject(NewBrush&)
End Sub

Private Sub APILine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 '* English: Use the API LineTo for Fast Drawing.
 '* Español: Pinta líneas de forma sencilla y rápida.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, x1, y1, PT)
 Call LineTo(UserControl.hDC, x2, y2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim PT   As POINTAPI
 
 '* English: Paint a rectangle using API.
 '* Español: Pinta el rectángulo de un Objeto.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(hDC, hPen)
 Call MoveToEx(hDC, X, Y, PT)
 Call LineTo(hDC, X + W, Y)
 Call LineTo(hDC, X + W, Y + H)
 Call LineTo(hDC, X, Y + H)
 Call LineTo(hDC, X, Y)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Function

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long) As Long
 '* English: Blend two colors in a 50%.
 '* Español: Mezclar dos colores al 50%.
 BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ &H10000) And &HFF)) / 2)
End Function

Private Function CalcTextWidth(ByVal isControl As Boolean, ByVal strCtlCaption As String, Optional ByVal strCtlCaption1 As String = "", Optional ByVal isW As Integer = 0) As String
 Dim lngMaxWidth  As Long, lngX As Long
 Dim lngTextWidth As Long
  
 '* English: Establishes the width of the text of the control.
 '* Español: Establece el ancho del texto del control.
 If (isControl = True) Then
  lngMaxWidth = UserControl.ScaleWidth - Int(UserControl.TextWidth(strCtlCaption) / 2)
  lngTextWidth = UserControl.TextWidth(strCtlCaption)
 Else
  lngMaxWidth = picList.ScaleWidth - Int(picList.TextWidth(strCtlCaption) / 2)
  lngTextWidth = picList.TextWidth(strCtlCaption) + isW
 End If
 If (strCtlCaption1 = "") Then strCtlCaption1 = strCtlCaption
 lngX = (Len(strCtlCaption) / 2)
 While (lngTextWidth > lngMaxWidth) And (lngX > 3)
  strCtlCaption1 = Mid$(strCtlCaption1, 1, lngX) & IIf(Len(strCtlCaption1) = lngX, "", "...")
  If (isControl = True) Then
   lngTextWidth = UserControl.TextWidth(strCtlCaption1)
  Else
   lngTextWidth = picList.TextWidth(strCtlCaption1)
  End If
  lngX = lngX - 1
 Wend
 CalcTextWidth = strCtlCaption1
End Function
        
Public Sub ChangeItem(ByVal Index As Long, ByVal Item As String, Optional ByVal ColorTextItem As OLE_COLOR = &HC56A31, Optional ByVal ImageItem As StdPicture = Nothing, Optional ByVal EnabledItem As Boolean = True, Optional ByVal ToolTipTextItem As String = "", Optional ByVal IndexItem As Long = -1, Optional ByVal ItemTag As String = "", Optional ByVal MouseIcon As StdPicture = Nothing, Optional ByVal SeparatorLine As Boolean = False)
Attribute ChangeItem.VB_Description = "Modifies an item of the list."
 '* English: Modifies an item of the list.
 '* Español: Modifica un item de la lista.
 ListContents(Index).Color = IIf(EnabledItem = True, ColorTextItem, ShiftColorOXP(DisabledColor))
 ListContents(Index).Text = Item
 ListContents(Index).Enabled = EnabledItem
 If (IndexItem > 0) And (IndexItem < sumItem) And (NoFindIndex(IndexItem) = False) Then ListContents(Index).Index = IndexItem
 Set ListContents(Index).MouseIcon = MouseIcon
 ListContents(Index).SeparatorLine = SeparatorLine
 ListContents(Index).ToolTipText = ToolTipTextItem
 ListContents(Index).Tag = ItemTag
 Set ListContents(Index).Image = ImageItem
 If Not (ImageItem Is Nothing) Then IsPicture = True
End Sub
        
Public Sub Clear()
Attribute Clear.VB_Description = "Clear the list."
 '* English: Clear the list.
 '* Español: Borra toda la lista.
 sumItem = 0
 ReDim ListContents(0)
 Text = ""
 ItemFocus = -1
 HighlightedItem = -1
 IndexItemNow = -1
 ListIndex = -1
 IsPicture = False
 
 RaiseEvent TotalItems(sumItem)
 Refresh
End Sub

Private Sub ControlsSubClasing()
 '* English: Add controls the Usercontrol.
 '* Español: Agrega controles al Usercontrol.
 Set scrollI = UserControl.Controls.Add("VB.VScrollBar", "scrollI")
 Set picTemp = UserControl.Controls.Add("VB.PictureBox", "picTemp")
 picTemp.AutoRedraw = True
 picTemp.ScaleMode = vbPixels
 picTemp.AutoSize = True
 picTemp.TabStop = False
 scrollI.TabStop = False
End Sub
        
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 '* English: Convert Long to System Color.
 '* Español: Convierte un long en un color del sistema.
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function
        
Private Sub CreateImage(ByVal myPicture As StdPicture, ByVal ObjecthDC As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal Disabled As Boolean = False, Optional ByVal nHeight As Long = 16, Optional ByVal nWidth As Long = 16, Optional ByVal nColor As OLE_COLOR = "&HFFFFFF")
 Dim sTMPpathFName As String
 
 '* English: Draw the image in the Object.
 '* Español: Crea la imagen sobre el Objeto.
 Set picTemp.Picture = myPicture
 picTemp.BackColor = &HF0
 If (Disabled = False) Then
  Call PicDisabled(picTemp)
 Else
  sTMPpathFName = TempPathName + "\~ConvIconToBmp.tmp"
  Call SavePicture(picTemp.Image, sTMPpathFName)
  Set picTemp.Picture = LoadPicture(sTMPpathFName)
  Call Kill(sTMPpathFName)
 End If
 picTemp.Refresh
 Call CreateImageMask(picTemp, picTemp, &HF0)
 Call StretchBlt(ObjecthDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcAnd)
 Call CreateImageSprite(picTemp, picTemp, &HF0)
 Call StretchBlt(ObjecthDC, X, Y, nWidth, nHeight, picTemp.hDC, 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight, vbSrcInvert)
End Sub

'* Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6077&lngWId=1. & _
   Thanks to David Peace Author of the code.
Private Function CreateImageMask(ByRef PicSrc As PictureBox, ByRef picDest As PictureBox, ByVal bColor As OLE_COLOR)
 Dim Looper  As Integer, Looper2 As Integer
 Dim bColor2 As OLE_COLOR

 picDest.Cls
 For Looper = 0 To PicSrc.Height
  picDest.Refresh
  For Looper2 = 0 To PicSrc.Width
   If (PicSrc.Point(Looper2, Looper) = bColor) Then
    bColor2 = RGB(255, 255, 255)
   Else
    bColor2 = RGB(0, 0, 0)
   End If
   Call SetPixel(picDest.hDC, Looper2, Looper, bColor2)
  Next
 Next
 picDest.Refresh
End Function

'* Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6077&lngWId=1. & _
   Thanks to David Peace Author of the code.
Private Function CreateImageSprite(ByRef PicSrc As PictureBox, ByRef picDest As PictureBox, ByVal bColor As OLE_COLOR)
 Dim Looper  As Integer, Looper2 As Integer
 Dim bColor2 As OLE_COLOR

 picDest.Cls
 For Looper = 0 To PicSrc.Height
  picDest.Refresh
  For Looper2 = 0 To PicSrc.Width
   If (PicSrc.Point(Looper2, Looper) = bColor) Then
    bColor2 = RGB(0, 0, 0)
   Else
    bColor2 = GetPixel(PicSrc.hDC, Looper2, Looper)
   End If
   SetPixel picDest.hDC, Looper2, Looper, bColor2
  Next
 Next
 picDest.Refresh
End Function

Private Function CreateMacOSXRegion() As Long
 Dim pPoligon(8) As POINTAPI, lW As Long, lh As Long
 
 '* English: Create a nonrectangular region for the MAC OS X Style.
 '* Español: Crea el Estilo MAC OS X.
 lW = UserControl.ScaleWidth
 lh = UserControl.ScaleHeight
 pPoligon(0).X = 0:      pPoligon(0).Y = 2
 pPoligon(1).X = 2:      pPoligon(1).Y = 0
 pPoligon(2).X = lW - 2: pPoligon(2).Y = 0
 pPoligon(3).X = lW:     pPoligon(3).Y = 2
 pPoligon(4).X = lW:     pPoligon(4).Y = lh - 5
 pPoligon(5).X = lW - 6: pPoligon(5).Y = lh
 pPoligon(6).X = 3:      pPoligon(6).Y = lh
 pPoligon(7).X = 0:      pPoligon(7).Y = lh - 3
 CreateMacOSXRegion = CreatePolygonRgn(pPoligon(0), 8, 1)
End Function

Private Function CreatePicture(ByVal myIndex As Long, ByVal CurrentS As Long, Optional ByVal nColor As OLE_COLOR) As Boolean
 Dim xS As Long
 
 '* English: Set the picture of the list.
 '* Español: Crea la imagen sobre la lista.
 If (sumItem = 0) Or (myIndex > ListCount) Then Exit Function
On Error GoTo myErr
 CreatePicture = False
 If Not (ListContents(myIndex).Image Is Nothing) Then
  Call CreateImage(ListContents(myIndex).Image, picList.hDC, 4, CurrentS + xS + 3, ListContents(myIndex).Enabled, , , nColor)
  CreatePicture = True
 End If
 Exit Function
myErr:
 Debug.Print Err.Description
End Function

Private Sub CreateText(ByVal Counter As Long, Optional ByVal Left As Integer = 0)
 Dim Msg As String, isText As String, isW As Integer, isLeft As Integer
 
 '* English: Set the text of the list.
 '* Español: Crea el texto sobre el objeto.
On Error Resume Next
 With picList
  .CurrentX = 0
  Set .Font = g_Font
  If (myAlignCombo = 0) Then
   '* English: Alignment to the left.
   '* Español: Alineación a la izquierda.
   Msg = CalcTextWidth(False, ListContents(Counter + 1).Text)
   isW = 390 + Left
  ElseIf (myAlignCombo = 1) Then
   '* English: Alignment to the right.
   '* Español: Alineación a la derecha.
   isText = SameSize(ListContents(Counter + 1).Text, " ")
   Msg = CalcTextWidth(False, isText, ListContents(Counter + 1).Text, Int(.TextWidth(ListContents(Counter + 1).Text) / 2))
   If (NumberItemsToShow < MaxListLength) Then isLeft = 650 Else isLeft = 120
   isW = Abs((.ScaleWidth + IIf(NumberItemsToShow < MaxListLength, scrollI.Width, 0)) - .TextWidth(Msg) - isLeft)
  ElseIf (myAlignCombo = 2) Then
   '* English: Alignment to the Center.
   '* Español: Alineación en el centro.
   isText = SameSize(ListContents(Counter + 1).Text, " ")
   Msg = CalcTextWidth(False, isText, ListContents(Counter + 1).Text, Int(.TextWidth(ListContents(Counter + 1).Text) / 2))
   isW = Int(.ScaleWidth / 2) - Int(.TextWidth(Msg) / 2)
  End If
  .CurrentX = isW
  picList.Print Msg
 End With
End Sub

Private Sub DrawAppearance(Optional ByVal Style As ComboAppearance = 1, Optional ByVal m_State As Integer = 1)
 Dim isText    As String, isW As Integer
 Dim m_lRegion As Long, isH   As Integer
 
 '* English: Draw appearance of the control.
 '* Español: Dibuja la apariencia del control.
 Cls
 AutoRedraw = True
 FillStyle = 1
 If (Style <> 6) Then UserControl.BackColor = myBackColor
 m_StateG = m_State
 isH = 0
On Error Resume Next
 With txtCombo
  .Height = Abs(ScaleHeight / 2 - 7)
  .Top = Abs(ScaleHeight / 2 - 7) + 0.5
  .ForeColor = IIf(Enabled = True, myNormalColorText, ShiftColorOXP(DisabledColor))
 On Error Resume Next
  Set .Font = g_Font
  If (myStyleCombo = 1) Then
   .Visible = False
  Else
   .Visible = True
  End If
 End With
 If (Height < 300) And (Style <> 11) Then
  Height = 300
 ElseIf (Height < 310) And (Style = 11) Then
  Height = 310
 ElseIf (Height > 600) Then
  If (Style = 12) Then
   Height = 300
  Else
   Height = 310
  End If
 End If
 If (Width < 840) Then Width = 840
 If (m_StateG <> 3) Then picList.Visible = False
 Select Case Style
  Case 1
   Call DrawOfficeButton(myOfficeAppearance)
  Case 2
   '* English: Style Windows 98.
   '* Español: Estilo Windows 98.
   Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, EDGE_SUNKEN)
   Call APIFillRect(UserControl.hDC, m_btnRect, GetSysColor(COLOR_BTNFACE))
   tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
   Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 3, EDGE_SUNKEN, EDGE_RAISED))
   Call DrawStandardArrow(m_btnRect, ArrowColor)
  Case 3
   '* English: Style Windows Xp.
   '* Español: Estilo Windows Xp.
   If (myXpAppearance = 1) Then     '* Aqua.
    tmpColor = &HB99D7F
   ElseIf (myXpAppearance = 2) Then '* Olive Green.
    tmpColor = &H94CCBC
   ElseIf (myXpAppearance = 3) Then '* Silver.
    tmpColor = &HA29594
   ElseIf (myXpAppearance = 4) Then '* TasBlue.
    tmpColor = &HF09F5F
   ElseIf (myXpAppearance = 5) Then '* Gold.
    tmpColor = &HBFE7F0
   ElseIf (myXpAppearance = 6) Then '* Blue.
    tmpColor = ShiftColorOXP(&HA0672F, 123)
   ElseIf (myXpAppearance = 7) Then '* Custom.
    If (m_StateG = 1) Then
     tmpColor = NormalBorderColor
    ElseIf (m_StateG = 2) Then
     tmpColor = HighLightBorderColor
    ElseIf (m_StateG = 3) Then
     tmpColor = SelectBorderColor
    End If
   End If
   If (myXpAppearance <> 0) Then
    Call APIRectangle(UserControl.hDC, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, IIf(m_StateG <> -1, tmpColor, &HDEE7E7))
    Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, GetSysColor(COLOR_WINDOW))
   End If
   Call DrawWinXPButton(myXpAppearance)
   If (myXpAppearance <> 0) Then
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, UserControl.BackColor)
    Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, UserControl.BackColor)
   End If
  Case 4
   '* English: Style Soft.
   '* Español: Estilo Suavizado.
   Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER)
   Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(NormalBorderColor, 228), ShiftColorOXP(NormalBorderColor, 155)))
   tempBorderColor = GetSysColor(COLOR_BTNFACE)
   Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, GetSysColor(COLOR_BTNFACE))
   Call APILine(m_btnRect.Left - 1, m_btnRect.Top, m_btnRect.Left - 1, m_btnRect.Bottom, GetSysColor(COLOR_BTNFACE))
   Call DrawCtlEdgeByRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 3, BDR_SUNKENOUTER, BDR_RAISEDINNER))
   Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(ArrowColor, 106), ArrowColor))
  Case 5
   Call DrawKDEButton
  Case 6
   '* English: Style MAC.
   '* Español: Estilo MAC.
   isH = 2
   Call DrawMacOSXCombo
  Case 7
   '* English: Style JAVA.
   '* Español: Estilo JAVA.
   tmpColor = ShiftColorOXP(NormalBorderColor, 52)
   tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
   Call DrawJavaBorder(0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
   Call APIFillRect(UserControl.hDC, m_btnRect, IIf(m_StateG = 2, tmpColor, IIf(m_StateG <> -1, NormalBorderColor, ShiftColorOXP(NormalBorderColor, 192))))
   Call DrawJavaBorder(m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_WINDOW))
   Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(ArrowColor, 166), ArrowColor))
  Case 8
   Call DrawExplorerBarButton(m_StateG)
  Case 9
   Dim tempPict As StdPicture
   
   '* English: Style User Picture.
   '* Español: Estilo Imagen de Usuario.
   Set tempPict = Nothing
   If (m_StateG = 1) Then
    Set tempPict = myNormalPictureUser
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    Set tempPict = myHighLightPictureUser
    tmpColor = HighLightBorderColor
   ElseIf (m_StateG = 3) Then
    Set tempPict = myFocusPictureUser
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
   Else
    Set tempPict = myDisabledPictureUser
    tmpColor = ShiftColorOXP(NormalBorderColor, 43)
   End If
   Call DrawRectangleBorder(ScaleWidth - 19, 0, ScaleWidth, ScaleHeight, GetLngColor(Parent.BackColor), False)
   Call DrawRectangleBorder(0, 0, ScaleWidth - 19, ScaleHeight, IIf(m_StateG <> -1, tmpColor, ShiftColorOXP(DisabledColor, 145)), True)
   If Not (tempPict Is Nothing) Then Call CreateImage(tempPict, UserControl.hDC, ScaleWidth - 18, Abs(Int(ScaleHeight / 2) - 11), True, 18, 17, GetLngColor(Parent.BackColor))
  Case 10
   '* English: Special Style.
   '* Español: Estilo especial con borde recortado.
   If (m_StateG = 1) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 75)
   ElseIf (m_StateG = 2) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 45)
   ElseIf (m_StateG = 3) Then
    tmpColor = ShiftColorOXP(&HDCC6B4, 15)
   Else
    tmpColor = ShiftColorOXP(&H0&, 237)
   End If
   Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 1, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 25), False)
   If (m_StateG = 1) Then
    tmpColor = ShiftColorOXP(&HC56A31, 143)
   ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
    tmpColor = ShiftColorOXP(&HC56A31, 113)
    tempBorderColor = tmpColor
   Else
    tmpColor = ShiftColorOXP(&H0&)
   End If
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5), True)
   tmpC2 = 12
   For tmpC1 = 2 To 5
    tmpC2 = tmpC2 + 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, BackColor)
   Next
   tmpC2 = 17
   tmpC3 = -2
   For tmpC1 = 5 To 2 Step -1
    tmpC2 = tmpC2 - 1
    tmpC3 = tmpC3 + 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
   Next
   Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
  Case 11
   '* English: Rounded Style.
   '* Español: Estilo Circular.
   isH = 2
   tempBorderColor = ShiftColorOXP(&H9F3000, 45)
   If (m_StateG = 1) Then
    iFor = &HCF989F
    tmpColor = &HA07F7F
    cValor = &HFFFFFF
   ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
    iFor = &H9F3000
    tmpColor = &HAF572F
    cValor = &HFFFFFF
   Else
    tmpColor = ShiftColorOXP(&H404040, 166)
    iFor = &HFFF8FF
    cValor = ShiftColorOXP(&H404040, 16)
   End If
   FillStyle = 0
   FillColor = iFor
   UserControl.Circle (m_btnRect.Left + 7, CInt(UserControl.ScaleHeight / 2)), 8, tmpColor
   UserControl.Circle (m_btnRect.Left + 7, CInt(UserControl.ScaleHeight / 2)), 7, &HFFFFFF
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   m_btnRect.Left = m_btnRect.Left - 5
   m_btnRect.Top = CInt(UserControl.ScaleHeight / 2) - 11
   UserControl.Line (m_btnRect.Left + 9, m_btnRect.Top + 8)-(m_btnRect.Left + 13, m_btnRect.Top + 12), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 10, m_btnRect.Top + 8)-(m_btnRect.Left + 13, m_btnRect.Top + 11), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 15, m_btnRect.Top + 8)-(m_btnRect.Left + 11, m_btnRect.Top + 12), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 14, m_btnRect.Top + 8)-(m_btnRect.Left + 11, m_btnRect.Top + 11), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 9, m_btnRect.Top + 12)-(m_btnRect.Left + 13, m_btnRect.Top + 16), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 10, m_btnRect.Top + 12)-(m_btnRect.Left + 13, m_btnRect.Top + 15), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 15, m_btnRect.Top + 12)-(m_btnRect.Left + 11, m_btnRect.Top + 16), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
   UserControl.Line (m_btnRect.Left + 14, m_btnRect.Top + 12)-(m_btnRect.Left + 11, m_btnRect.Top + 15), IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), cValor)
  Case 12
   Call DrawGradientButton(1)
  Case 13
   Call DrawGradientButton(2)
  Case 14
   Call DrawLightBlueButton
  Case 15
   '* English: Arrow Style.
   '* Español: Estilo Flecha.
   If (m_StateG = 1) Then
    cValor = GetLngColor(GradientColor1)
    iFor = GetLngColor(GradientColor2)
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
    iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
    tmpColor = HighLightBorderColor
   ElseIf (m_StateG = 3) Then
    cValor = GetLngColor(GradientColor1)
    iFor = GetLngColor(GradientColor2)
    tmpColor = SelectBorderColor
   Else
    cValor = GetLngColor(ShiftColorOXP(GradientColor1))
    iFor = GetLngColor(GradientColor2)
    tmpColor = ShiftColorOXP(&H0&)
   End If
   tempBorderColor = tmpColor
   Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1, iFor, cValor, 1)
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
   cValor = UserControl.ScaleHeight / 2 + 1
   iFor = IIf(m_StateG = -1, &HC0C0C0, ArrowColor)
   For tmpColor = 7 To -2 Step -1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, cValor - (tmpColor / 2), UserControl.ScaleWidth - 7, cValor - (tmpColor / 2), IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
   Next
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 5, cValor, UserControl.ScaleWidth - 6, cValor, IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, cValor + 1, UserControl.ScaleWidth - 7, cValor + 1, IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 7, cValor + 2, UserControl.ScaleWidth - 8, cValor + 2, IIf(m_StateG = -1, ShiftColorOXP(iFor, 26), iFor))
  Case 16
   Call DrawNiaWBSSButton
  Case 17
   Call DrawRhombusButton
  Case 18
   Call DrawXpButton
  Case 19
   '* English: Ardent Style.
   '* Español: Estilo Ardent.
   If (m_StateG = 1) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 24)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 65)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = HighLightBorderColor
   ElseIf (m_StateG = 3) Then
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2), 14)
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
   Else
    tmpC1 = ShiftColorOXP(BlendColors(GradientColor1, GradientColor2))
    cValor = tmpC1
    iFor = tmpC1
    tmpColor = &HC0C0C0
   End If
   Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
   Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, tmpColor)
   Call DrawRectangleBorder(1, 1, UserControl.ScaleWidth - 19, UserControl.ScaleHeight - 2, ShiftColorOXP(cValor, 85))
   tmpC1 = 7
   tmpC2 = 4
   tmpC3 = ScaleHeight / 2 + 1
   For tmpColor = 6 To 2 Step -1
    tmpC1 = tmpC1 - 1
    tmpC2 = tmpC2 - 1
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC1, tmpC3 + tmpC2, UserControl.ScaleWidth - (tmpColor + 2), tmpC3 + tmpC2, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))
    Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC1, tmpC3 - 2 + tmpC2 - 1, UserControl.ScaleWidth - (tmpColor + 2), tmpC3 - 2 + tmpC2 - 1, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))
   Next
  Case 20
   Call DrawChocolateButton
  Case 21
   Call DrawButtonDownload
 End Select
 Call SetRect(m_btnRect, UserControl.ScaleWidth - 18, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)
 If (Style = 6) Then
  If (m_lRegion <> 0) Then Call DeleteObject(m_lRegion)
  m_lRegion = CreateMacOSXRegion
  Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
 Else
  m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
  Call SetWindowRgn(UserControl.hWnd, m_lRegion, True)
 End If
 If (ItemFocus > 0) Then
  '* English: Sets the image of the current item.
  '* Español: Establece la imagen del item actual.
  picTemp.BackColor = ListColor
  If Not (ListContents(ItemFocus).Image Is Nothing) Then
   Call CreateImage(ListContents(ItemFocus).Image, UserControl.hDC, 5, Abs(Int(ScaleHeight / 2) - 7) - IIf(Style = 7, 1, 0), Enabled, , , myBackColor)
   cValor = 27
  Else
   cValor = 8
  End If
  isText = ListContents(ItemFocus).Text
  isW = 47 + isH
 Else
  isText = Text
  cValor = 8
  isW = 26 + isH
 End If
 txtCombo.Left = cValor
 If (myStyleCombo = 1) Then
  With UserControl
   .CurrentX = cValor
   .CurrentY = Int(UserControl.ScaleHeight / 2) - 7
   Set .Font = g_Font
   If (Enabled = False) Then
    Call SetTextColor(.hDC, DisabledColor)
   Else
    Call SetTextColor(.hDC, NormalColorText)
   End If
   isText = CalcTextWidth(True, myText)
   Call DrawStateString(.hDC, 0, 0, isText, Len(isText), .CurrentX, .CurrentY, 0, 0, DST_TEXT Or DSS_NORMAL)
  End With
 End If
 txtCombo.Width = Abs(ScaleWidth - isW)
 picList.Width = Width
 txtCombo.BackColor = myBackColor
End Sub

Private Sub DrawButtonDownload()
 '* English: Draw Button Download appearance.
 '* Español: Crea la apariencia de un Botón de Descarga.
 cValor = IIf(m_StateG = -1, ShiftColorOXP(&H92603C), &H92603C)
 tempBorderColor = cValor
 tmpC3 = IIf(m_StateG = -1, ShiftColorOXP(&HE0C6AE), &HE0C6AE)
 If (m_StateG = 1) Or (m_StateG = 3) Then
  tmpC1 = &HBE8F63
  tmpC2 = &HE8DBCB
  tmpColor = ArrowColor
 ElseIf (m_StateG = 2) Then
  tmpC1 = ShiftColorOXP(&HBE8F63, 49)
  tmpC2 = ShiftColorOXP(&HE8DBCB, 49)
  tmpColor = ShiftColorOXP(ArrowColor, 89)
 Else
  tmpC1 = ShiftColorOXP(&HBE8F63)
  tmpC2 = ShiftColorOXP(&HE8DBCB)
  tmpColor = ShiftColorOXP(&HC0C0C0, 85)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 2, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tmpC1, tmpC2, 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(cValor, 5))
 Call DrawXpArrow(tmpColor)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, Int(UserControl.ScaleHeight / 2) + 4, UserControl.ScaleWidth - 6, Int(UserControl.ScaleHeight / 2) + 4, tmpColor)
 Call DrawShadow(tmpC3, tmpC3, False)
End Sub

Private Sub DrawChocolateButton()
 '* English: Chocolate Style.
 '* Español: Estilo Chocolate.
 cValor = IIf(m_StateG = -1, ShiftColorOXP(&H4A464B), &H4A464B)
 tempBorderColor = cValor
 tmpC3 = &HFFFFFF
 If (m_StateG = 1) Or (m_StateG = 3) Then
  tmpC1 = &H686567
  tmpC2 = ShiftColorOXP(&H292929, 89)
  tmpColor = &H0
 ElseIf (m_StateG = 2) Then
  tmpC1 = ShiftColorOXP(&H686567, 89)
  tmpC2 = ShiftColorOXP(&H292929, 178)
  tmpColor = ShiftColorOXP(&H0, 89)
 Else
  tmpC1 = ShiftColorOXP(&H838181)
  tmpC2 = ShiftColorOXP(&H292929)
  tmpColor = ShiftColorOXP(&H0)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 2, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tmpC1, tmpC2, 2)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(cValor, 5))
 Call DrawShadow(tmpC3, tmpColor, False)
 m_btnRect.Bottom = m_btnRect.Bottom / 2 + 4
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, m_btnRect.Bottom + 2, UserControl.ScaleWidth - 5, m_btnRect.Bottom + 2, tmpColor)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, m_btnRect.Bottom + 3, UserControl.ScaleWidth - 5, m_btnRect.Bottom + 3, tmpColor)
 For iFor = 4 To 7
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor - 1, m_btnRect.Bottom - iFor + 5, UserControl.ScaleWidth - iFor - 1, m_btnRect.Bottom - iFor + 5, tmpColor)
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor - 1, m_btnRect.Bottom + iFor, UserControl.ScaleWidth - (iFor + 1), m_btnRect.Bottom + iFor, tmpColor)
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor, m_btnRect.Bottom - iFor + 5, UserControl.ScaleWidth - (iFor + 2), m_btnRect.Bottom - iFor + 5, &HFFFFFF)
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + iFor, m_btnRect.Bottom + iFor, UserControl.ScaleWidth - (iFor + 2), m_btnRect.Bottom + iFor, &HFFFFFF)
 Next
End Sub

Private Sub DrawCtlEdge(ByVal hDC As Long, ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, Optional ByVal Style As Long = EDGE_RAISED, Optional ByVal flags As Long = BF_RECT)
 Dim R As RECT
 
 '* English: The DrawEdge function draws one or more edges of rectangle. _
             using the specified coords.
 '* Español: Dibuja uno ó más bordes del rectángulo.
 With R
  .Left = X
  .Top = Y
  .Right = X + W
  .Bottom = Y + H
 End With
 Call DrawEdge(hDC, R, Style, flags)
End Sub

Private Sub DrawCtlEdgeByRect(ByVal hDC As Long, ByRef RT As RECT, Optional ByVal Style As Long = EDGE_RAISED, Optional ByVal flags As Long = BF_RECT)
 '* English: Draws the edge in a rect.
 '* Español: Colorea uno ó más bordes del rectángulo del Control.
 Call DrawEdge(hDC, RT, Style, flags)
End Sub

Private Sub DrawExplorerBarButton(ByVal m_StateG As Long)
 '* English: Style ExplorerBar.
 '* Español: Estilo ExplorerBar.
 myBackColor = ShiftColorOXP(&HDEEAF0, 184)
 txtCombo.BackColor = myBackColor
 UserControl.BackColor = myBackColor
 Call DrawRectangleBorder(1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, &HEAF3F7)
 If (m_StateG = 1) Then
  cValor = ShiftColorOXP(&HB6BFC3, 91)
  iFor = &HEAF3F7
  tmpColor = ShiftColorOXP(&HB6BFC3, 162)
 ElseIf (m_StateG = 2) Then
  cValor = ShiftColorOXP(&HB6BFC3, 31)
  iFor = &HDCEBF1
  tmpColor = ShiftColorOXP(&HB6BFC3, 132)
 ElseIf (m_StateG = 3) Then
  cValor = ShiftColorOXP(&HB6BFC3, 21)
  iFor = &HCEE3EC
  tmpColor = ShiftColorOXP(&HB6BFC3, 112)
  tempBorderColor = ShiftColorOXP(&HB6BFC3, 21)
 Else
  UserControl.BackColor = ShiftColorOXP(&HEAF3F7, 124)
  txtCombo.BackColor = UserControl.BackColor
  cValor = ShiftColorOXP(&HB6BFC3, 84)
  tmpC1 = ShiftColorOXP(&HEAF3F7, 124)
  iFor = ShiftColorOXP(&HEAF3F7, 123)
  tmpColor = ShiftColorOXP(&HB6BFC3, 132)
 End If
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, cValor)
 If (m_StateG = -1) Then Call DrawRectangleBorder(1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpC1)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, iFor, False)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, 2, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 2, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 18, m_btnRect.Bottom - 1, UserControl.BackColor)
 Call SetPixel(UserControl.hDC, m_btnRect.Right - 1, UserControl.ScaleHeight - 3, UserControl.BackColor)
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 196), ArrowColor))
End Sub

Private Sub DrawGradient(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Direction As Integer)
 Dim Vert(1) As TRIVERTEX, gRect As GRADIENT_RECT

 '* English: Draw a gradient in the selected coords and hDC.
 '* Español: Dibuja el objeto en forma degradada.
 Call LongToRGB(Color1)
 With Vert(0)
  .X = X
  .Y = Y
  .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
  .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
  .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
  .Alpha = 1
 End With
 Call LongToRGB(Color2)
 With Vert(1)
  .X = x1
  .Y = y1
  .Red = Val("&H" & Hex$(RGBColor.Red) & "00")
  .Green = Val("&H" & Hex$(RGBColor.Green) & "00")
  .Blue = Val("&H" & Hex$(RGBColor.Blue) & "00")
  .Alpha = 0
 End With
 gRect.UpperLeft = 0
 gRect.LowerRight = 1
 If (Direction = 1) Then
  Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)
 Else
  Call GradientFillRect(hDC, Vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H)
 End If
End Sub

Private Sub DrawGradientButton(ByVal WhatGradient As Long)
 '* English: Draw a Vertical or Horizontal Gradient style appearance.
 '* Español: Dibuja la apariencia degradada bien sea vertical ó horizontal.
 If (m_StateG = 1) Then
  tmpColor = ShiftColorOXP(&HC56A31, 133)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HD8CEC5)
 ElseIf (m_StateG = 2) Then
  tmpColor = ShiftColorOXP(&HC56A31, 113)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HD6BEB5)
 ElseIf (m_StateG = 3) Then
  tmpColor = ShiftColorOXP(&HC56A31, 93)
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HB3A29B)
  tempBorderColor = tmpColor
 Else
  tmpColor = CLng(ShiftColorOXP(&H0&))
  cValor = GetLngColor(&HC0C0C0)
  iFor = GetLngColor(&HFFFFFF)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, cValor, iFor, WhatGradient)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor, True)
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawJavaBorder(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal lColorShadow As Long, ByVal lColorLight As Long, ByVal lColorBack As Long)
 '* English: Draw the edge with a JAVA style.
 '* Español: Dibuja el borde estilo JAVA.
 Call APIRectangle(UserControl.hDC, X, Y, W - 1, H - 1, lColorShadow)
 Call APIRectangle(UserControl.hDC, X + 1, Y + 1, W - 1, H - 1, lColorLight)
 Call SetPixel(UserControl.hDC, X, Y + H, lColorBack)
 Call SetPixel(UserControl.hDC, X + W, Y, lColorBack)
 Call SetPixel(UserControl.hDC, X + 1, Y + H - 1, BlendColors(lColorLight, lColorShadow))
 Call SetPixel(UserControl.hDC, X + W - 1, Y + 1, BlendColors(lColorLight, lColorShadow))
End Sub

Private Sub DrawKDEButton()
 '* English: Style KDE.
 '* Español: Estilo KDE.
 If (m_StateG = 1) Then
  tmpColor = NormalBorderColor
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 63))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 63))
 ElseIf (m_StateG = 2) Then
  tmpColor = HighLightBorderColor
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 127))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 127))
 ElseIf (m_StateG = 3) Then
  tmpColor = SelectBorderColor
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
 Else
  tmpColor = &HC0C0C0
  cValor = GetLngColor(&HFFFFFF)
  iFor = ShiftColorOXP(GetLngColor(&HC0C0C0), 45)
 End If
 tempBorderColor = tmpColor
 '* Español: Top Left.
 '* Español: Parte Superior Izquierda.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - 8, m_btnRect.Bottom - 8, iFor, cValor, 1)
 '* Español: Top Right.
 '* Español: Parte Inferior Derecha.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 8, m_btnRect.Top + 8, m_btnRect.Right, m_btnRect.Bottom, cValor, iFor, 1)
 '* Español: Bottom Right.
 '* Español: Parte Inferior Derecha.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 8, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom - 8, iFor, cValor, 1)
 '* Español: Bottom Left.
 '* Español: Parte Inferior Izquierda.
 Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top + 8, m_btnRect.Right - 8, m_btnRect.Bottom, cValor, iFor, 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 2, 16, UserControl.ScaleHeight - 4, tmpColor, True)
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawLightBlueButton()
 Dim PT      As POINTAPI, cX As Long, cY As Long
 Dim hPenOld As Long, hPen   As Long
   
 '* English: Style LightBlue.
 '* Español: Estilo LightBlue.
 If (m_StateG = 1) Or (m_StateG = 3) Then
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HA87057)
  tmpColor = &HA69182
  tempBorderColor = tmpColor
 ElseIf (m_StateG = 2) Then
  cValor = GetLngColor(&HFFFFFF)
  iFor = GetLngColor(&HCFA090)
  tmpColor = &HAF9080
 Else
  cValor = GetLngColor(&HFFFFFF)
  iFor = ShiftColorOXP(GetLngColor(&HA87057))
  tmpColor = ShiftColorOXP(&HA69182, 146)
 End If
 Call DrawGradient(UserControl.hDC, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1, cValor, iFor, 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 If (m_StateG = 2) Then
  Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, &H53969F)
  Call DrawRectangleBorder(UserControl.ScaleWidth - 16, 1, 15, UserControl.ScaleHeight - 2, &H92C4D8)
  tmpColor = &H3EB4DE
 ElseIf (m_StateG <> -1) Then
  Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
  tmpColor = ArrowColor
 End If
 cX = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
 cY = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 + 2
 hPen = CreatePen(0, 1, IIf(m_StateG <> -1, tmpColor, ShiftColorOXP(&HC0C0C0, 97)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, cX - 3, cY - 1, PT)
 Call LineTo(UserControl.hDC, cX + 1, cY - 1)
 Call LineTo(UserControl.hDC, cX, cY)
 Call LineTo(UserControl.hDC, cX - 2, cY)
 Call LineTo(UserControl.hDC, cX, cY + 2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
 hPen = CreatePen(0, 1, IIf(m_StateG <> -1, tmpColor, ShiftColorOXP(&HC0C0C0, 97)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 cX = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 3
 Call MoveToEx(UserControl.hDC, cX - 4, cY - 3, PT)
 Call LineTo(UserControl.hDC, cX, cY - 3)
 Call LineTo(UserControl.hDC, cX - 2, cY - 5)
 Call LineTo(UserControl.hDC, cX - 3, cY - 4)
 Call LineTo(UserControl.hDC, cX - 1, cY - 3)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub DrawList(ByVal TopItem As Integer, ByVal NumberOfItems As Integer)
 Dim Counter As Long, d     As Long, R As RECT
 Dim isFocus As Long, Valor As Long
 
 '* English: Draw the list with the elements.
 '* Español: Crea la lista con los elementos guardados.
 picList.Cls
 If (myListGradient = True) Then Call DrawGradient(picList.hDC, 0, 0, picList.ScaleWidth, picList.ScaleHeight, MSSoftColor(myGradientColor1), MSSoftColor(myGradientColor2), 2)
 picList.Line (0, 0)-(picList.ScaleWidth - 1, picList.ScaleHeight - 1), tempBorderColor, B
 If (ListCount < 1) Then Exit Sub
 With picList
  .AutoRedraw = True
  .ScaleMode = vbTwips
  Counter = TopItem - 1
  CurrentS = -20
  isFocus = -1
 On Error Resume Next
  Do Until (Counter = TopItem + NumberOfItems)
   CurrentS = CurrentS + 20
   d = (Counter - TopItem) * 255 + CurrentS * 2 + 27
  On Error Resume Next
   Valor = IIf((HighlightedItem + 1) < ListCount, HighlightedItem + 1, ListCount)
   If ((ListContents(Counter + 1).Enabled = True) And (Len(txtCombo.Text) > 0) And (ItemFocus = Counter + 1) And (ListContents(Counter + 1).Text = txtCombo.Text) And (FirstView = 0)) Or ((Counter = HighlightedItem) And (ListContents(Valor).Enabled = True) And (FirstView = 1)) Then
    If (scrollI.Visible = True) Then
     picList.Line (3, d - 45)-(picList.Width - 290, d + 245), SelectListColor, BF
    ElseIf (Abs(Counter - scrollI.Value + 1) = NumberItemsToShow) And (ListCount > 1) Then
     picList.Line (3, d - 45)-(picList.Width, d + Abs(picList.Height - (d + 20))), SelectListColor, BF
    Else
     picList.Line (3, d - 45)-(picList.Width, d + 255), SelectListColor, BF
    End If
    If (Abs(Counter - scrollI.Value + 1) = 1) And (ListCount > 1) Then
     picList.Line (3, d - 65)-(picList.Width - 15, d + 255), SelectListBorderColor, B
    ElseIf (Abs(Counter - scrollI.Value + 1) = NumberItemsToShow) And (ListCount > 1) Then
     picList.Line (3, d - 45)-(picList.Width - 15, d + Abs(picList.Height - (d + 20))), SelectListBorderColor, B
    ElseIf (ListCount = Counter) Then
     picList.Line (3, d - 45)-(picList.Width, d - 55), SelectListBorderColor, B
    ElseIf (ListCount = 1) Then
     picList.Line (0, 0)-(picList.ScaleWidth - IIf(scrollI.Visible = False, 8, 88), picList.ScaleHeight - 8), SelectListBorderColor, B
    Else
     picList.Line (3, d - 45)-(picList.Width - IIf(scrollI.Visible = False, 15, 95), d + IIf(Abs(picList.Height - (d + 21)) > 255, 255, Abs(picList.Height - (d + 21)))), SelectListBorderColor, B
    End If
    picList.ForeColor = HighLightColorText
    picList.CurrentY = d + 20 - IIf(ListCount = 1, 9, 0)
    Call CreatePicture(Counter + 1, Abs(CurrentS - 20), SelectBorderColor)
    Call CreateText(Counter)
    picList.ToolTipText = ListContents(Counter + 1).ToolTipText
    picList.MousePointer = vbCustom
    Set picList.MouseIcon = ListContents(Counter + 1).MouseIcon
    isFocus = Counter
   Else
    If (ListContents(Counter + 1).Enabled = False) Then
     picList.ForeColor = DisabledColor
    Else
     picList.ForeColor = ListContents(Counter + 1).Color
    End If
    If (scrollI.Visible = True) Then tmpC1 = 240 Else tmpC1 = 30
    If (Counter < sumItem) And (sumItem > 1) And (ListContents(Counter).SeparatorLine = True) And (Counter <> isFocus + 1) Then picList.Line (20, picList.CurrentY + 32)-(picList.ScaleWidth - tmpC1, picList.CurrentY + 32), vbButtonShadow, B
   End If
   picList.CurrentY = d + 20 - IIf(ListCount = 1, 9, 0)
   Call CreatePicture(Counter + 2, CurrentS, ListColor)
   If (ListContents(Counter + 1).TextShadow = True) Then
    picList.ForeColor = myShadowColorText
    Call CreateText(Counter)
    picList.ForeColor = ListContents(Counter + 1).Color
    picList.CurrentY = d + 35
    Call CreateText(Counter, 15)
   Else
    Call CreateText(Counter)
   End If
   Counter = Counter + 1
  Loop
  .ScaleMode = vbPixels
 End With
 FirstView = 1
End Sub

Private Sub DrawMacOSXCombo()
 Dim PT      As POINTAPI, cY  As Long, cX     As Long, Color1 As Long, ColorG As Long
 Dim hPen    As Long, hPenOld As Long, Color2 As Long, Color3 As Long, ColorH As Long
 Dim Color4  As Long, Color5  As Long, Color6 As Long, Color7 As Long, ColorI As Long
 Dim Color8  As Long, Color9  As Long, ColorA As Long, ColorB As Long
 Dim ColorC  As Long, ColorD  As Long, ColorE As Long, ColorF As Long
 
 '* English: Draw the Mac OS X combo (this is a cool style!).
 '* Español: Dibujar el combo estilo Mac OS X (este es un estilo chevere).
 m_btnRect.Left = m_btnRect.Left - 4
 tempBorderColor = GetSysColor(COLOR_BTNSHADOW)
 '* English: Button gradient top.
 ColorA = &HA0A0A0
 UserControl.BackColor = BackColor
 If (m_StateG = 1) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 9)
  Color2 = ShiftColorOXP(&HDE8B45, 9)
  Color3 = ShiftColorOXP(&HDD873E, 9)
  Color4 = ShiftColorOXP(&HB33A01, 9)
  Color5 = ShiftColorOXP(&HE9BD96, 9)
  Color6 = ShiftColorOXP(&HB9B2AD, 9)
  Color7 = ShiftColorOXP(&H968A82, 9)
  Color8 = ShiftColorOXP(&HA25022, 9)
  Color9 = ShiftColorOXP(&HB8865E, 9)
  ColorB = ShiftColorOXP(&HDFBC86, 9)
  ColorC = ShiftColorOXP(&HFFBA77, 9)
  ColorD = ShiftColorOXP(&HE3D499, 9)
  ColorE = ShiftColorOXP(&HFFD996, 9)
  ColorF = ShiftColorOXP(&HE1A46D, 9)
  ColorG = ShiftColorOXP(&HCBA47B, 9)
  ColorH = ShiftColorOXP(&HDFDFDF, 9)
  ColorI = ShiftColorOXP(&HD0D0D0, 9)
 ElseIf (m_StateG = 2) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 89)
  Color2 = ShiftColorOXP(&HDE8B45, 89)
  Color3 = ShiftColorOXP(&HDD873E, 89)
  Color4 = ShiftColorOXP(&HB33A01, 99)
  Color5 = ShiftColorOXP(&HE9BD96, 109)
  Color6 = ShiftColorOXP(&HB9B2AD, 109)
  Color7 = ShiftColorOXP(&H968A82, 109)
  Color8 = ShiftColorOXP(&HA25022, 109)
  Color9 = ShiftColorOXP(&HB8865E, 109)
  ColorB = ShiftColorOXP(&HDFBC86, 109)
  ColorC = ShiftColorOXP(&HFFBA77, 109)
  ColorD = ShiftColorOXP(&HE3D499, 109)
  ColorE = ShiftColorOXP(&HFFD996, 109)
  ColorF = ShiftColorOXP(&HE1A46D, 109)
  ColorG = ShiftColorOXP(&HCBA47B, 109)
  ColorH = ShiftColorOXP(&HDFDFDF, 109)
  ColorI = ShiftColorOXP(&HD0D0D0, 109)
 ElseIf (m_StateG = 3) Then
  Color1 = ShiftColorOXP(&HFDF2C3, 15)
  Color2 = ShiftColorOXP(&HDE8B45, 15)
  Color3 = ShiftColorOXP(&HDD873E, 15)
  Color4 = ShiftColorOXP(&HB33A01, 15)
  Color5 = ShiftColorOXP(&HE9BD96, 15)
  Color6 = ShiftColorOXP(&HB9B2AD, 15)
  Color7 = ShiftColorOXP(&H968A82, 15)
  Color8 = ShiftColorOXP(&HA25022, 15)
  Color9 = ShiftColorOXP(&HB8865E, 15)
  ColorB = ShiftColorOXP(&HDFBC86, 15)
  ColorC = ShiftColorOXP(&HFFBA77, 15)
  ColorD = ShiftColorOXP(&HE3D499, 15)
  ColorE = ShiftColorOXP(&HFFD996, 15)
  ColorF = ShiftColorOXP(&HE1A46D, 15)
  ColorG = ShiftColorOXP(&HCBA47B, 15)
  ColorH = ShiftColorOXP(&HDFDFDF, 15)
  ColorI = ShiftColorOXP(&HD0D0D0, 15)
 Else
  Color1 = ShiftColorOXP(&H808080, 195)
  Color2 = ShiftColorOXP(&H808080, 135)
  Color3 = ShiftColorOXP(&H808080, 135)
  Color4 = ShiftColorOXP(&H808080, 5)
  Color5 = Color1
  Color6 = GetLngColor(Parent.BackColor)
  Color7 = Color6
  Color8 = ShiftColorOXP(&H808080, 65)
  Color9 = Color6
  ColorA = Color6
  ColorB = Color4
  ColorC = Color4
  ColorD = Color4
  ColorE = Color4
  ColorF = Color4
  ColorG = Color4
  ColorH = Color6
  ColorI = Color6
 End If
 Call DrawVGradient(Color1, Color2, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 3)
 '* English: Button gradient bottom.
 Call DrawVGradient(Color3, Color1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, UserControl.ScaleHeight / 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight * 2 / 3 - 4)
 '* English: Lines for the text area.
 Call APILine(2, 0, UserControl.ScaleWidth - 3, 0, &HA1A1A1)
 Call APILine(1, 0, 1, UserControl.ScaleHeight - 3, &HA1A1A1)
 '* English: Left shadow.
 If (m_StateG <> -1) Then
  Call DrawVGradient(ColorH, &HBBBBBB, 0, 0, 1, 3)
  Call DrawVGradient(&HBBBBBB, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
  Call DrawVGradient(ColorA, &HBBBBBB, 0, UserControl.ScaleHeight / 2, 1, UserControl.ScaleHeight / 2 - 5)
  Call DrawVGradient(&HBBBBBB, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
 Else
  Call DrawVGradient(ColorH, ColorH, 0, 0, 1, 3)
  Call DrawVGradient(ColorA, ColorA, 0, 4, 1, UserControl.ScaleHeight / 2 - 4)
  Call DrawVGradient(ColorA, ColorA, 0, UserControl.ScaleHeight / 2, 1, UserControl.ScaleHeight / 2 - 5)
  Call DrawVGradient(ColorH, ColorH, 0, UserControl.ScaleHeight - 5, 1, 2)
 End If
 '* English: Bottom shadows.
 Call APILine(1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, &H747474)
 Call APILine(1, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 2, &HA1A1A1)
 Call APILine(2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, &HDDDDDD)
 '* English: Lines for the button area.
 Call DrawVGradient(ColorB, Color3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 1, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight / 3)
 Call DrawVGradient(Color3, ColorB, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, UserControl.ScaleHeight / 3, UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, UserControl.ScaleHeight * 2 / 3 - 4)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left, 0, UserControl.ScaleWidth - 3, 0, Color4)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 1, 1, UserControl.ScaleWidth - 4, 1, Color5)
 '* English: Right shadow.
 Call DrawVGradient(ColorH, ColorI, UserControl.ScaleWidth - 1, 2, UserControl.ScaleWidth, 3)
 Call DrawVGradient(ColorI, ColorA, UserControl.ScaleWidth - 1, 3, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
 Call DrawVGradient(ColorA, ColorI, UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 2 - 2, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 6)
 Call DrawVGradient(ColorI, ColorH, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 8, UserControl.ScaleWidth, 3)
 '* English: Layer1.
 Call DrawVGradient(Color4, Color3, UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 7)
 '* English: Layer2.
 Call DrawVGradient(Color4, ColorC, UserControl.ScaleWidth - 3, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6)
 '* English: Doted Area / 1-Bottom.
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4, ColorG)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 4, Color7)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5, ColorF)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 5, Color7)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 6, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 4, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 2, &HCACACA)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 2, &HBFBFBF)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 1, &HE4E4E4)
 '* English: Doted Area / 2-Botom
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 4, ColorD)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 5, ColorE)
 '* English: Doted Area / 3-Top.
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 0, IIf(m_StateG <> -1, &HA76E4A, ShiftColorOXP(&H808080, 55)))
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 0, Color6)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 1, Color8)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 2, 1, IIf(m_StateG <> -1, &HB3A49D, GetLngColor(Parent.BackColor)))
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 5, 1, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 1, Color8)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 4, 2, Color9)
 Call SetPixel(UserControl.hDC, UserControl.ScaleWidth - 3, 3, Color8)
 '* English: Draw Twin Arrows.
 cX = m_btnRect.Left + (m_btnRect.Right - m_btnRect.Left) / 2 + 2
 cY = m_btnRect.Top + (m_btnRect.Bottom - m_btnRect.Top) / 2 - 1
 hPen = CreatePen(0, 1, IIf(m_StateG <> -1, &H0&, ShiftColorOXP(&H0&)))
 hPenOld = SelectObject(UserControl.hDC, hPen)
 '* English: Down Arrow.
 Call MoveToEx(UserControl.hDC, cX - 3, cY + 1, PT)
 Call LineTo(UserControl.hDC, cX + 1, cY + 1)
 Call LineTo(UserControl.hDC, cX, cY + 2)
 Call LineTo(UserControl.hDC, cX - 2, cY + 2)
 Call LineTo(UserControl.hDC, cX - 2, cY + 3)
 Call LineTo(UserControl.hDC, cX, cY + 3)
 Call LineTo(UserControl.hDC, cX - 1, cY + 4)
 Call LineTo(UserControl.hDC, cX - 1, cY + 6)
 '* English: Up Arrow.
 Call MoveToEx(UserControl.hDC, cX - 3, cY - 2, PT)
 Call LineTo(UserControl.hDC, cX + 1, cY - 2)
 Call LineTo(UserControl.hDC, cX, cY - 3)
 Call LineTo(UserControl.hDC, cX - 2, cY - 3)
 Call LineTo(UserControl.hDC, cX - 2, cY - 4)
 Call LineTo(UserControl.hDC, cX, cY - 4)
 Call LineTo(UserControl.hDC, cX - 1, cY - 5)
 Call LineTo(UserControl.hDC, cX - 1, cY - 7)
 '* English: Destroy PEN.
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
 '* English: Undo the offset.
 m_btnRect.Left = m_btnRect.Left + 4
End Sub

Private Sub DrawNiaWBSSButton()
 '* English: NiaWBSS Style.
 '* Español: Estilo NiaWBSS.
 If (m_StateG = 1) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = NormalBorderColor
 ElseIf (m_StateG = 2) Then
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
  tmpColor = HighLightBorderColor
 ElseIf (m_StateG = 3) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = SelectBorderColor
  tempBorderColor = tmpColor
 Else
  cValor = GetLngColor(ShiftColorOXP(GradientColor1))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2))
  tmpColor = ShiftColorOXP(DisabledColor, 156)
 End If
 Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
 tmpC1 = UserControl.ScaleHeight / 2 - 2
 tmpC2 = IIf(m_StateG = -1, DisabledColor, ArrowColor)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 2, tmpC1 - 1, UserControl.ScaleWidth - 12, tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 10, tmpC1 - 1, UserControl.ScaleWidth - 4, tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 3, tmpC1, UserControl.ScaleWidth - 10, tmpC1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 8, tmpC1, UserControl.ScaleWidth - 5, tmpC1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 4, tmpC1 + 1, UserControl.ScaleWidth - 6, tmpC1 + 1, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 4, tmpC1 + 2, UserControl.ScaleWidth - 6, tmpC1 + 2, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
 For tmpC3 = 3 To 6
  If (tmpC3 = 3) Or (tmpC3 = 4) Then
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 5, tmpC1 + tmpC3, UserControl.ScaleWidth - 7, tmpC1 + tmpC3, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
  Else
   Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 6, tmpC1 + tmpC3, UserControl.ScaleWidth - 8, tmpC1 + tmpC3, IIf(m_StateG = -1, ShiftColorOXP(tmpC2, 166), tmpC2))
  End If
 Next
End Sub

Private Sub DrawOfficeButton(ByVal WhatOffice As ComboOfficeAppearance)
 Dim tmpRect As RECT
 
 '* English: Draw Office Style appearance.
 '* Español: Dibuja la apariencia de Office.
 tmpRect = m_btnRect
 Select Case WhatOffice
  Case 0
   '* English: Style Office Xp, appearance default.
   '* Español: Estilo Office Xp, apariencia por defecto.
   If (m_StateG = 1) Then
    '* English: Normal Color.
    '* Español: Color Normal.
    tmpColor = NormalBorderColor
   ElseIf (m_StateG = 2) Then
    '* English: Highlight Color.
    '* Español: Color de Selección MouseMove.
    tmpColor = HighLightBorderColor
    cValor = 185
   ElseIf (m_StateG = 3) Then
    '* English: Down Color.
    '* Español: Color de Selección MouseDown.
    tmpColor = SelectBorderColor
    tempBorderColor = tmpColor
    cValor = 125
   Else
    '* English: Disabled Color.
    '* Español: Color deshabilitado.
    tmpColor = ConvertSystemColor(ShiftColorOXP(NormalBorderColor, 41))
   End If
   If (m_StateG > 1) Then
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tmpColor, B
    UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 14, UserControl.ScaleHeight - 2), ShiftColorOXP(tmpColor, cValor), BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), tmpColor, B
   Else
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tmpColor, B
    UserControl.Line (UserControl.ScaleWidth - 3, 2)-(UserControl.ScaleWidth - 13, UserControl.ScaleHeight - 3), tmpColor, BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 1), tmpColor, B
   End If
   Call DrawStandardArrow(m_btnRect, ArrowColor)
  Case 1
   '* English: Style Office 2000.
   '* Español: Estilo Office 2000.
   If (m_StateG = 1) Then
    '* English: Flat.
    '* Español: Normal.
    tmpColor = NormalBorderColor
    Call DrawRectangleBorder(UserControl.ScaleWidth - 13, 1, 12, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 175), False)
   ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
    '* English: Mouse Hover or Mouse Pushed.
    '* Español: Mouse presionado o MouseMove.
    If (m_StateG = 2) Then
     tmpColor = ShiftColorOXP(HighLightBorderColor)
    Else
     tmpColor = ShiftColorOXP(SelectBorderColor)
     tempBorderColor = tmpColor
    End If
    Call DrawCtlEdge(UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER)
    tmpRect.Left = tmpRect.Left + 4
    Call APIFillRect(UserControl.hDC, tmpRect, tmpColor)
    tmpRect.Left = tmpRect.Left - 1
    Call APIRectangle(UserControl.hDC, 1, 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, tmpColor)
    Call APILine(tmpRect.Left, tmpRect.Top, tmpRect.Left, tmpRect.Bottom, tmpColor)
    If (m_StateG = 2) Then
     Call DrawCtlEdgeByRect(UserControl.hDC, tmpRect, BDR_RAISEDINNER)
    Else
     Call DrawCtlEdgeByRect(UserControl.hDC, tmpRect, BDR_SUNKENOUTER)
    End If
    m_btnRect.Left = m_btnRect.Left - 3
   Else
    '* English: Disabled control.
    '* Español: Control deshabilitado.
    Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, ShiftColorOXP(&HC0C0C0, 36))
    Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 1, 17, UserControl.ScaleHeight - 2, myBackColor, False)
   End If
   tmpRect.Left = tmpRect.Left + 4
   Call DrawStandardArrow(tmpRect, IIf(m_StateG = -1, ShiftColorOXP(&HC0C0C0, 36), ArrowColor))
  Case 2
   '* English: Style Office 2003.
   '* Español: Estilo Office 2003.
   If (m_StateG <> -1) Then
    tmpC2 = GetSysColor(COLOR_WINDOW)
   Else
    tmpC2 = ShiftColorOXP(GetSysColor(COLOR_BTNFACE))
   End If
   tmpC1 = ArrowColor
   UserControl.BackColor = tmpC2
   txtCombo.BackColor = tmpC2
   tmpColor = GetSysColor(COLOR_HOTLIGHT)
   If (m_StateG = 1) Then
    cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), GetSysColor(29)), 109)
    iFor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_INACTIVECAPTIONTEXT), GetSysColor(COLOR_GRADIENTINACTIVECAPTION)))
   ElseIf (m_StateG = 2) Then
    cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), GetSysColor(29)), 170)
    iFor = cValor
   ElseIf (m_StateG = 3) Then
    cValor = ShiftColorOXP(BlendColors(GetSysColor(COLOR_GRADIENTACTIVECAPTION), GetSysColor(29)), 140)
    iFor = cValor
   Else
    tmpC1 = GetSysColor(COLOR_GRAYTEXT)
    Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpC1)
    txtCombo.ForeColor = tmpC1
    GoTo DrawNowArrow
   End If
   Call DrawGradient(UserControl.hDC, m_btnRect.Left + 4, tmpRect.Top - 1, tmpRect.Right + 1, tmpRect.Bottom + 1, iFor, cValor, 1)
   If (m_StateG = 2) Or (m_StateG = 3) Then
    Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
    Call DrawRectangleBorder(UserControl.ScaleWidth - 15, 0, 17, UserControl.ScaleHeight, tmpColor, True)
    tempBorderColor = tmpColor
   End If
DrawNowArrow:
   Call DrawStandardArrow(tmpRect, tmpC1)
   myBackColor = tmpC2
 End Select
End Sub

Private Sub DrawRectangleBorder(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional ByVal SetBorder As Boolean = True)
 Dim hBrush As Long, tempRect As RECT

 '* English: Draw a rectangle.
 '* Español: Crea el rectángulo.
 tempRect = m_btnRect
 m_btnRect.Left = X
 m_btnRect.Top = Y
 m_btnRect.Right = X + Width
 m_btnRect.Bottom = Y + Height
 hBrush = CreateSolidBrush(Color)
 If (SetBorder = True) Then
  Call FrameRect(UserControl.hDC, m_btnRect, hBrush)
 Else
  Call FillRect(UserControl.hDC, m_btnRect, hBrush)
 End If
 Call DeleteObject(hBrush)
 m_btnRect = tempRect
End Sub

Private Sub DrawRhombusButton()
 '* English: Rhombus Style.
 '* Español: Estilo Rombo.
 If (m_StateG = 1) Then
  tmpColor = ShiftColorOXP(NormalBorderColor, 25)
 ElseIf (m_StateG = 2) Then
  tmpColor = ShiftColorOXP(HighLightBorderColor, 25)
 ElseIf (m_StateG = 3) Then
  tmpColor = ShiftColorOXP(SelectBorderColor, 25)
 Else
  tmpColor = ShiftColorOXP(&H0&, 237)
 End If
 Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 1, 16, UserControl.ScaleHeight - 2, ShiftColorOXP(tmpColor, 25), False)
 If (m_StateG = 1) Then
  tmpColor = ShiftColorOXP(ArrowColor, 143)
 ElseIf (m_StateG = 2) Or (m_StateG = 3) Then
  tmpColor = ShiftColorOXP(ArrowColor, 113)
  tempBorderColor = tmpColor
 Else
  tmpColor = ShiftColorOXP(&H0&)
 End If
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 17, 0, 17, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5), True)
 '* English: Left top border.
 '* Español: Borde Superior Izquierdo.
 tmpC2 = 12
 For tmpC1 = 2 To 5
  tmpC2 = tmpC2 + 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, BackColor)
 Next
 tmpC2 = 17
 tmpC3 = -2
 For tmpC1 = 5 To 2 Step -1
  tmpC2 = tmpC2 - 1
  tmpC3 = tmpC3 + 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
 Next
 '* English: Left bottom border.
 '* Español: Borde Inferior Izquierdo.
 tmpC2 = 17
 For tmpC1 = 3 To 1 Step -1
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 1, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, BackColor)
 Next
 tmpC2 = 12
 tmpC3 = 3
 For tmpC1 = 1 To 3
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, tmpColor)
 Next
 '* English: Right top border.
 '* Español: Borde Superior Derecho.
 tmpC2 = 0
 tmpC3 = 23
 For tmpC1 = 6 To 1 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, GetLngColor(Parent.BackColor))
 Next
 tmpC2 = 0
 tmpC3 = 17
 For tmpC1 = 6 To 1 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, tmpC1 - 1, UserControl.ScaleWidth - tmpC2, tmpC1 - 1, tmpColor)
 Next
 '* English: Right bottom border.
 '* Español: Borde Inferior Derecho.
 tmpC2 = 6
 For tmpC1 = 0 To 3
  tmpC2 = tmpC2 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + 19, UserControl.ScaleHeight - tmpC1 - 1, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 1, GetLngColor(Parent.BackColor))
 Next
 tmpC2 = 1
 tmpC3 = 16
 For tmpC1 = 3 To 0 Step -1
  tmpC2 = tmpC2 + 1
  tmpC3 = tmpC3 - 1
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC3, UserControl.ScaleHeight - tmpC1 - 2, UserControl.ScaleWidth - tmpC2, UserControl.ScaleHeight - tmpC1 - 2, tmpColor)
 Next
 m_btnRect.Left = m_btnRect.Left + 1
 Call DrawStandardArrow(m_btnRect, IIf(m_StateG = -1, ShiftColorOXP(&H404040, 166), ArrowColor))
End Sub

Private Sub DrawShadow(ByVal iColor1 As Long, ByVal iColor2 As Long, Optional ByVal SoftColor As Boolean = True)
 '* English: Set a Shadow Border.
 '* Español: Coloca un borde con sombra.
 tmpC2 = 15
 If (SoftColor = True) Then
  tmpC3 = 178
  iFor = 10
 Else
  tmpC3 = 0
  iFor = 0
 End If
 For tmpC1 = 1 To 16
  tmpC2 = tmpC2 - 1
  '* Horizontal Top Border.
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, 1, UserControl.ScaleWidth - tmpC1, 1, ShiftColorOXP(iColor1, tmpC3))
  '* Horizontal Bottom Border.
  Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left + tmpC2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - tmpC1, UserControl.ScaleHeight - 2, IIf(m_StateG = -1, ShiftColorOXP(iColor2, tmpC3), ShiftColorOXP(iColor2, iFor)))
  If (SoftColor = True) Then
   tmpC3 = tmpC3 - 5
   iFor = iFor + 5
  End If
 Next
 m_btnRect.Bottom = m_btnRect.Bottom - 11
 If (SoftColor = True) Then
  tmpC3 = 128
  iFor = 70
 End If
 For tmpC1 = 0 To 12
  '* Vertical Left Border.
  Call APILine(m_btnRect.Left + 1, m_btnRect.Top + tmpC1 - 1, m_btnRect.Left + 1, m_btnRect.Bottom + tmpC1 - 1, ShiftColorOXP(iColor1, tmpC3))
  '* Vertical Right Border.
  Call APILine(UserControl.ScaleWidth - 2, m_btnRect.Top + tmpC1 - 1, UserControl.ScaleWidth - 2, m_btnRect.Bottom + tmpC1 - 1, IIf(m_StateG = -1, ShiftColorOXP(iColor2, tmpC3), ShiftColorOXP(iColor2, iFor)))
  If (SoftColor = True) Then
   tmpC3 = tmpC3 + 5
   iFor = iFor - 5
  End If
 Next
End Sub

Private Sub DrawStandardArrow(ByRef RT As RECT, ByVal lColor As Long)
 Dim PT   As POINTAPI, hPenOld As Long, cX As Long
 Dim hPen As Long, cY          As Long
 
 '* English: Draw the standard arrow in a Rect.
 '* Español: Dibuje la flecha normal en un Rect.
 If (AppearanceCombo = 1) And (OfficeAppearance = 1) Or (AppearanceCombo = 10) Or (AppearanceCombo = 17) Then
  hPen = 1
 ElseIf ((OfficeAppearance = 2) Or (OfficeAppearance = 0)) And (AppearanceCombo = 1) Then
  hPen = 2
 End If
 cX = RT.Left + (RT.Right - RT.Left) - (7 - hPen)
 cY = RT.Top + (RT.Bottom - RT.Top) / 2
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(UserControl.hDC, hPen)
 Call MoveToEx(UserControl.hDC, cX - 3, cY - 1, PT)
 Call LineTo(UserControl.hDC, cX + 1, cY - 1)
 Call LineTo(UserControl.hDC, cX, cY)
 Call LineTo(UserControl.hDC, cX - 2, cY)
 Call LineTo(UserControl.hDC, cX, cY + 2)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT) As Boolean
 Dim hTheme  As Long '* hTheme Handle.
 Dim lResult As Long '* Temp Variable.
 
 '* If a error occurs then or we are not running XP or the visual style is Windows Classic.
On Error GoTo NoXP
 '* Get out hTheme Handle.
 hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
 '* Did we get a theme handle?.
 If (hTheme) Then
  '* Yes! Draw the control Background.
  lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
  '* If drawing was successful, return true, or false If not.
  DrawTheme = IIf(lResult, False, True)
 Else
  '* No, we couldn't get a hTheme, drawing failed.
  DrawTheme = False
 End If
 '* Close theme.
 Call CloseThemeData(hTheme)
 '* Exit the function now.
 Exit Function
NoXP:
 '* An Error was detected, drawing Failed.
 DrawTheme = False
End Function

Private Sub DrawVGradient(ByVal lEndColor As Long, ByVal lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal x2 As Long, ByVal y2 As Long)
 Dim dR As Single, dG As Single, dB As Single, Ni As Long
 Dim sR As Single, sG As Single, sB As Single
 Dim eR As Single, eG As Single, eB As Single
 
 '* English: Draw a Vertical Gradient in the current hDC.
 '* Español: Dibuja un degradado en forma vertical.
 sR = (lStartcolor And &HFF)
 sG = (lStartcolor \ &H100) And &HFF
 sB = (lStartcolor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / y2
 dG = (sG - eG) / y2
 dB = (sB - eB) / y2
 For Ni = 0 To y2
  Call APILine(X, Y + Ni, x2, Y + Ni, RGB(eR + (Ni * dR), eG + (Ni * dG), eB + (Ni * dB)))
 Next
End Sub

Private Sub DrawWinXPButton(ByVal XpAppearance As ComboXpAppearance)
 Dim tmpXPAppearance   As ComboXpAppearance, isState  As Integer
 Dim bDrawThemeSuccess As Boolean, tmpRect            As RECT
 
 '* English: This Sub Draws the XpAppearance Button.
 '* Español: Este procedimiento dibuja el Botón estilo XP.
 If (XpAppearance <> 7) Then myBackColor = GetSysColor(COLOR_WINDOW)
 If (XpAppearance = 0) Then
  '* Draw the XP Themed Style.
  isState = IIf(m_StateG < 0, 4, m_StateG)
  Call SetRect(tmpRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
  bDrawThemeSuccess = DrawTheme("Edit", 2, isState, tmpRect)
  Call SetRect(tmpRect, m_btnRect.Left - 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
  bDrawThemeSuccess = DrawTheme("ComboBox", 1, isState, tmpRect)
  If (bDrawThemeSuccess = True) Then
   Exit Sub
  Else '* If themed failed, then use the Next Style.
   myBackColor = BackColor
   tmpXPAppearance = 7 '* If failed, use custom colors.
   GoTo noUxThemed
  End If
 Else
  tmpXPAppearance = XpAppearance
 End If
noUxThemed:
 Select Case tmpXPAppearance
  Case 1
   '* English: Style WinXp Aqua.
   '* Español: Estilo WinXp Aqua.
   cValor = &H85614D
   tempBorderColor = &HC56A31
   tmpC2 = &HB99D7F
   If (m_StateG = 1) Then
    tmpC3 = &HF5C8B3
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = ShiftColorOXP(&HF5C8B3, 58)
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HF9A477
    tmpColor = &HFFFFFF
   End If
  Case 2
   '* English: Style WinXp Olive Green.
   '* Español: Estilo WinXp Olive Green.
   cValor = &HFFFFFF
   tempBorderColor = &H668C7D
   tmpC2 = &H94CCBC
   If (m_StateG = 1) Then
    tmpC3 = &H8BB4A4
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = &HA7D7CA
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = &H80AA98
    tmpColor = &HFFFFFF
   End If
  Case 3
   '* English: Style WinXp Silver.
   '* Español: Estilo WinXp Silver.
   tempBorderColor = &HA29594
   cValor = &H48483E
   tmpC2 = &HA29594
   If (m_StateG = 1) Then
    tmpC3 = &HDACCCB
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = ShiftColorOXP(&HDACCCB, 58)
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HE5D1CF
    tmpColor = &HFFFFFF
   End If
  Case 4
   '* English: Style WinXp TasBlue.
   '* Español: Estilo WinXp TasBlue.
   tempBorderColor = &HF09F5F
   cValor = ShiftColorOXP(&H703F00, 58)
   tmpC2 = &HF09F5F
   If (m_StateG = 1) Then
    tmpC3 = &HF0AF70
    tmpColor = &HFFE7CF
   ElseIf (m_StateG = 2) Then
    tmpC3 = ShiftColorOXP(&HF0BF80, 58)
    tmpColor = &HFFEFD0
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HF09F5F
    tmpColor = &HFFEFD0
   End If
  Case 5
   '* English: Style WinXp Gold.
   '* Español: Estilo WinXp Gold.
   tempBorderColor = &HBFE7F0
   cValor = ShiftColorOXP(&H6F5820, 45)
   tmpC2 = &HBFE7F0
   If (m_StateG = 1) Then
    tmpC3 = ShiftColorOXP(&HCFFFFF, 54)
    tmpColor = &HBFF0FF
   ElseIf (m_StateG = 2) Then
    tmpC3 = &HBFEFFF
    tmpColor = ShiftColorOXP(&HCFFFFF, 58)
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HCFFFFF
    tmpColor = &HBFE8FF
   End If
  Case 6
   '* English: Style WinXp Blue.
   '* Español: Estilo WinXp Blue.
   tempBorderColor = ShiftColorOXP(&HA0672F, 123)
   cValor = &H6F5820
   tmpC2 = ShiftColorOXP(&HA0672F, 123)
   If (m_StateG = 1) Then
    tmpC3 = &HEFF0F0
    tmpColor = &HF0F7F0
   ElseIf (m_StateG = 2) Then
    tmpC3 = &HF0F8FF
    tmpColor = &HF0F7F0
   ElseIf (m_StateG = 3) Then
    tmpC3 = &HF1946E
    tmpColor = &HEEC2B4
   End If
  Case 7
   '* English: Style WinXp Custom.
   '* Español: Estilo WinXp Custom.
   tempBorderColor = SelectBorderColor
   cValor = ArrowColor
   If (m_StateG = 1) Then
    tmpC3 = NormalBorderColor
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 2) Then
    tmpC3 = HighLightBorderColor
    tmpColor = &HFFFFFF
   ElseIf (m_StateG = 3) Then
    tmpC3 = SelectBorderColor
    tmpColor = &HFFFFFF
   End If
   tmpC2 = tmpC3
 End Select
 If (m_StateG = -1) Then
  tmpColor = &HE5ECEC
  tmpC3 = m_btnRect.Bottom - m_btnRect.Top
  For iFor = 3 To tmpC1
   Call APILine(m_btnRect.Left + 1, tmpC3 - iFor + 3, m_btnRect.Right - 1, tmpC3 - iFor + 3, tmpColor)
  Next
  tmpC1 = &HE5ECEC
  tmpC2 = &HEED2C1
 Else
  tmpC1 = tmpC2
  Call DrawGradient(UserControl.hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom, tmpColor, tmpC3, 1)
 End If
 Call APIRectangle(hDC, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, tmpC1)
 Call DrawXpArrow(IIf(m_StateG = -1, &HC2C9C9, cValor))
End Sub

Private Sub DrawXpArrow(Optional ByVal iColor3 As OLE_COLOR = &H0)
 '* English: Draw The XP Style Arrow.
 '* Español: Dibuja la flecha estilo Xp.
 tmpC1 = m_btnRect.Right - m_btnRect.Left
 tmpC2 = m_btnRect.Bottom - m_btnRect.Top + 1
 tmpC1 = m_btnRect.Left + tmpC1 / 2 + 1
 tmpC2 = m_btnRect.Top + tmpC2 / 2
 If (iColor3 = &H0) Then iColor3 = ArrowColor
 Call APILine(tmpC1 - 5, tmpC2 - 2, tmpC1, tmpC2 + 3, iColor3)
 Call APILine(tmpC1 - 4, tmpC2 - 2, tmpC1, tmpC2 + 2, iColor3)
 Call APILine(tmpC1 - 4, tmpC2 - 3, tmpC1, tmpC2 + 1, iColor3)
 Call APILine(tmpC1 + 3, tmpC2 - 2, tmpC1 - 2, tmpC2 + 3, iColor3)
 Call APILine(tmpC1 + 2, tmpC2 - 2, tmpC1 - 2, tmpC2 + 2, iColor3)
 Call APILine(tmpC1 + 2, tmpC2 - 3, tmpC1 - 2, tmpC2 + 1, iColor3)
End Sub

Private Sub DrawXpButton()
 '* English: Additional Xp Style.
 '* Español: Estilo Xp Adicional.
 If (m_StateG = 1) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = NormalBorderColor
 ElseIf (m_StateG = 2) Then
  cValor = GetLngColor(ShiftColorOXP(GradientColor1, 65))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2, 65))
  tmpColor = HighLightBorderColor
 ElseIf (m_StateG = 3) Then
  cValor = GetLngColor(GradientColor1)
  iFor = GetLngColor(GradientColor2)
  tmpColor = SelectBorderColor
  tempBorderColor = tmpColor
 Else
  cValor = GetLngColor(ShiftColorOXP(GradientColor1))
  iFor = GetLngColor(ShiftColorOXP(GradientColor2))
  tmpColor = DisabledColor
 End If
 Call DrawVGradient(cValor, iFor, m_btnRect.Left + 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1)
 Call DrawRectangleBorder(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpColor)
 Call DrawRectangleBorder(UserControl.ScaleWidth - 18, 0, 19, UserControl.ScaleHeight, ShiftColorOXP(tmpColor, 5))
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 3, 1, UserControl.ScaleWidth - 144, 1, IIf(m_StateG = -1, ShiftColorOXP(DisabledColor, 198), ShiftColorOXP(tmpColor, 168)))
 Call DrawXpArrow(IIf(m_StateG <> -1, ArrowColor, tmpColor))
 Call DrawShadow(GradientColor1, &H646464)
 Call APILine(UserControl.ScaleWidth - m_btnRect.Right + m_btnRect.Left - 3, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 144, UserControl.ScaleHeight - 2, IIf(m_StateG = -1, ShiftColorOXP(DisabledColor, 198), ShiftColorOXP(tmpColor, 168)))
End Sub

Private Sub Espera(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, FinSeg As Single
 
 '* English: Wait a certain time.
 '* Español: Esperar un determinado tiempo.
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 Do While FinSeg > Timer
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
 Loop
End Sub

Public Function FindItemText(ByVal Text As String, Optional ByVal Compare As StringCompare = 0) As Long
Attribute FindItemText.VB_Description = "Search Text in the list and return the position."
 Dim i As Long
 
 '* English: Search Text in the list and return the position.
 '* Español: Busca una cadena dentro de la lista y devuelve su posición en la misma.
 FindItemText = -1
 If (Text = "") Or (Compare < 0) Or (Compare > 2) Then Exit Function
 For i = 1 To sumItem
  If (Compare = 0) Then
   If (InStr(1, UCase$(ListContents(i).Text), UCase$(Text), vbBinaryCompare) <> 0) Then
    FindItemText = i
    Exit For
   End If
  ElseIf (Compare = 1) Then
   If (UCase$(Text) = UCase$(ListContents(i).Text)) Then
    FindItemText = i
    Exit For
   End If
  Else
   If (Text = ListContents(i).Text) Then
    FindItemText = i
    Exit For
   End If
  End If
 Next
End Function

Public Function GetControlVersion() As String
Attribute GetControlVersion.VB_Description = "Control Version."
 '* English: Control Version.
 '* Español: Version del Control.
 GetControlVersion = Version & " © " & Year(Now)
End Function

Private Function GetLngColor(ByVal Color As Long) As Long
 '* English: The GetSysColor function retrieves the current color of the specified display element. Display elements are the parts of a window and the Windows display that appear on the system display screen.
 '* Español: Recupera el color actual del elemento de despliegue especificado.
 If (Color And &H80000000) Then
  GetLngColor = GetSysColor(Color And &H7FFFFFFF)
 Else
  GetLngColor = Color
 End If
End Function

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean
 Dim mPos  As POINTAPI, KeyLeft As Boolean
 Dim oRect As RECT, KeyRight    As Boolean
 
 '* English: Verifies if the mouse is on the object or if one makes clic outside of him.
 '* Español: Verifica si el mouse se encuentra sobre el objeto ó si se hace clic fuera de él.
 Call GetCursorPos(mPos)
 Call GetWindowRect(ObjecthWnd, oRect)
 KeyLeft = GetAsyncKeyState(VK_LBUTTON)
 KeyRight = GetAsyncKeyState(VK_RBUTTON)
 UserControl.MousePointer = myMousePointer
 '* English: Set MouseIcon only drop down list.
 '* Español: Coloca el icono del mouse únicamente donde se expande ó retrae la lista.
 If (mPos.X > oRect.Left + (UserControl.ScaleWidth - 18)) And (mPos.X < oRect.Right) Then
  Set UserControl.MouseIcon = myMouseIcon
 Else
  Set UserControl.MouseIcon = Nothing
 End If
 If (mPos.X >= oRect.Left) And (mPos.X <= oRect.Right) And (mPos.Y >= oRect.Top) And (mPos.Y <= oRect.Bottom) Then
  InFocusControl = True
  First = 0
 ElseIf (KeyLeft = True) Or (KeyRight = True) And (First = 0) Then
  If (HighlightedItem > -1) And (FirstView <> 1) Then
   If (mPos.X < oRect.Left) Or (mPos.X > oRect.Right) Or (mPos.Y < oRect.Top) Or (mPos.Y > oRect.Bottom) Then
    m_LeaveMouse = False
    First = 1
    picList.Visible = False
   End If
  End If
 End If
End Function

Private Sub IsEnabled(ByVal isTrue As Boolean)
 '* English: Shows the state of Enabled or Disabled of the Control.
 '* Español: Muestra el estado de Habilitado ó Deshabilitado del Control.
 If (isTrue = True) Then
  Call DrawAppearance(myAppearanceCombo, 1)
 Else
  Call DrawAppearance(myAppearanceCombo, -1)
  tmrFocus.Enabled = False
 End If
End Sub

Public Sub ItemEnabled(ByVal ListIndex As Long, ByVal ValueItem As Boolean)
Attribute ItemEnabled.VB_Description = "Sets the Enabled/disabled property in an Item."
 '* English: Sets the Enabled/disabled property in an Item.
 '* Español: Habilita o Deshabilita un Item.
On Error GoTo myErr:
 ListContents(ListIndex).Enabled = ValueItem
 Exit Sub
myErr:
End Sub

Public Function List(ByVal ListIndex As Long) As String
Attribute List.VB_Description = "Show one item of the list."
 '* English: Show one item of the list.
 '* Español: Muestra un elemento de la lista.
 HighlightedItem = ListIndex
 ItemFocus = ListIndex
 List = ListContents(ListIndex).Text
 Call IsEnabled(ControlEnabled)
End Function

Private Function ListCount1() As Long
On Error Resume Next
 '* English: Total of elements of the list.
 '* Español: Total de elementos de la lista.
 If (sumItem = 0) Then Exit Function
 ListCount1 = UBound(ListContents) + 1
End Function

Private Function ListIndex1(Optional ByVal Item As Long = -1) As Long
 '* English: Function to know the position of the selected index of the list.
 '* Español: Función para saber la posición del index seleccionado de la lista.
On Error Resume Next
 If (Item = -1) And (ListCount1 > 0) Then
  ListIndex1 = IIf(ItemFocus = 0, -1, ItemFocus)
 Else
  ListIndex1 = Item
 End If
 If (ListIndex1 > 0) Then
  HighlightedItem = ListIndex1
  ItemFocus = ListIndex1
  Text = ListContents(ListIndex1).Text
 End If
End Function

Private Sub LongToRGB(ByVal lColor As Long)
 '* English: Convert a Long to RGB format.
 '* Español: Convierte un Long en formato RGB.
 RGBColor.Red = lColor And &HFF
 RGBColor.Green = (lColor \ &H100) And &HFF
 RGBColor.Blue = (lColor \ &H10000) And &HFF
End Sub

Private Function MSSoftColor(ByVal lColor As Long) As Long
 Dim lRed  As Long, lGreen As Long, lb As Long
 Dim lBlue As Long, lr     As Long, lg As Long
 
 '* English: Set a soft color.
 '* Español: Devuelve un color suave.
 lr = (lColor And &HFF)
 lg = ((lColor And 65280) \ 256)
 lb = ((lColor) And 16711680) \ 65536
 lRed = (76 - Int(((lColor And &HFF) + 32) \ 64) * 19)
 lGreen = (76 - Int((((lColor And 65280) \ 256) + 32) \ 64) * 19)
 lBlue = (76 - Int((((lColor And &HFF0000) \ &H10000) + 32) / 64) * 19)
 MSSoftColor = RGB(lr + lRed, lg + lGreen, lb + lBlue)
End Function

Private Function NoFindIndex(ByVal Index As Long) As Boolean
 Dim i As Long
 
 '* English: Search if the Index has not been assigned.
 '* Español: Busca si ya no se ha asignado este Index.
 NoFindIndex = False
 For i = 1 To sumItem
  If (ListContents(i).Index = Index) Then NoFindIndex = True: Exit For
 Next
End Function

Public Sub OrderList(Optional ByVal Order As Integer = 1)
Attribute OrderList.VB_Description = "Order the list with the search method (I Exchange)."
 Dim N As Long, i As Long, j As Long
 
 '* English: Order the list with the search method (I Exchange).
 '* Español: Ordena la lista con el método de búsqueda (Intercambio).
 If (Order <> 1) And (Order <> 2) Then Exit Sub
 ReDim OrderListContents(0)
 N = UBound(ListContents)
 For i = 1 To N
  ReDim Preserve OrderListContents(i)
  OrderListContents(i).Color = ListContents(i).Color
  OrderListContents(i).Enabled = ListContents(i).Enabled
  Set OrderListContents(i).Image = ListContents(i).Image
  OrderListContents(i).Index = ListContents(i).Index
  Set OrderListContents(i).MouseIcon = ListContents(i).MouseIcon
  OrderListContents(i).SeparatorLine = ListContents(i).SeparatorLine
  OrderListContents(i).Tag = ListContents(i).Tag
  OrderListContents(i).Text = ListContents(i).Text
  OrderListContents(i).ToolTipText = ListContents(i).ToolTipText
 Next
 i = 1
 For i = 1 To N
  For j = (i + 1) To N
   Select Case Order
    Case 1: If (OrderListContents(j).Text < OrderListContents(i).Text) Then Call SetInfo(i, j)
    Case 2: If (OrderListContents(j).Text > OrderListContents(i).Text) Then Call SetInfo(i, j)
   End Select
  Next
 Next
 ReDim ListContents(0)
 For i = 1 To N
  ReDim Preserve ListContents(i)
  ListContents(i).Color = OrderListContents(i).Color
  ListContents(i).Enabled = OrderListContents(i).Enabled
  Set ListContents(i).Image = OrderListContents(i).Image
  ListContents(i).Index = OrderListContents(i).Index
  Set ListContents(i).MouseIcon = OrderListContents(i).MouseIcon
  ListContents(i).SeparatorLine = OrderListContents(i).SeparatorLine
  ListContents(i).Tag = OrderListContents(i).Tag
  ListContents(i).Text = OrderListContents(i).Text
  ListContents(i).ToolTipText = OrderListContents(i).ToolTipText
 Next
 ReDim OrderListContents(0)
End Sub

Private Sub PicDisabled(ByRef picTo As PictureBox)
 Dim sTMPpathFName As String, lFlags As Long
 
 '* English: Disables a image.
 '* Español: Deshabilita la imagen.
 Select Case picTo.Picture.Type
  Case vbPicTypeBitmap
   lFlags = DST_BITMAP
  Case vbPicTypeIcon
   lFlags = DST_ICON
  Case Else
   lFlags = DST_COMPLEX
 End Select
 If Not (picTo.Picture Is Nothing) Then
  Call DrawState(picTo.hDC, 0, 0, picTo.Picture, 0, 0, 0, picTo.ScaleWidth, picTo.ScaleHeight, lFlags Or DSS_DISABLED)
  sTMPpathFName = TempPathName + "\~ConvIconToBmp.tmp"
  Call SavePicture(picTo.Image, sTMPpathFName)
  Set picTo.Picture = LoadPicture(sTMPpathFName)
  Call Kill(sTMPpathFName)
  picTo.Refresh
 End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "Delete a Item from the list."
 Dim TempList() As PropertyCombo, sCount As Long
 Dim Count      As Long, TempCount       As Long
 
 '* English: Delete a Item from the list.
 '* Español: Elimina un elemento de la lista.
On Error GoTo myErr
 If (ListCount = 0) Then Exit Sub
 If (sumItem > 0) Then sumItem = Abs(sumItem - 1)
 For Count = 1 To ListCount1 - 1
  If (Index <> Count) Then
   sCount = sCount + 1
   ReDim Preserve TempList(sCount)
   TempList(sCount).Color = ListContents(Count).Color
   TempList(sCount).Enabled = ListContents(Count).Enabled
   Set TempList(sCount).Image = ListContents(Count).Image
   TempList(sCount).Index = sCount
   TempList(sCount).Tag = ListContents(Count).Tag
   TempList(sCount).Text = ListContents(Count).Text
   TempList(sCount).ToolTipText = ListContents(Count).ToolTipText
   Set TempList(sCount).MouseIcon = ListContents(Count).MouseIcon
   TempList(sCount).SeparatorLine = ListContents(Count).SeparatorLine
  End If
 Next
 TempCount = Abs(Count - 2)
 sCount = 0
 ReDim ListContents(0)
 For Count = 1 To TempCount
  sCount = sCount + 1
  ReDim Preserve ListContents(sCount)
  ListContents(sCount).Color = TempList(Count).Color
  ListContents(sCount).Enabled = TempList(Count).Enabled
  Set ListContents(sCount).Image = TempList(Count).Image
  ListContents(sCount).Index = TempList(Count).Index
  ListContents(sCount).Tag = TempList(Count).Tag
  ListContents(sCount).Text = TempList(Count).Text
  ListContents(sCount).ToolTipText = TempList(Count).ToolTipText
  Set ListContents(sCount).MouseIcon = TempList(Count).MouseIcon
  ListContents(sCount).SeparatorLine = TempList(Count).SeparatorLine
 Next
 ReDim Preserve ListContents(TempCount)
 Refresh
 MaxListLength = Abs(MaxListLength - 1)
 If (myText = ListContents(MaxListLength + 1).Text) Then
  ListIndex = -1
  ItemFocus = -1
  Text = ""
  Call IsEnabled(ControlEnabled)
 End If
 RaiseEvent TotalItems(sumItem)
 Exit Sub
myErr:
End Sub

Private Function SameSize(ByVal isText As String, Optional ByVal isChar As String = " ") As String
 Dim isH As Long, isFor As Long
 
 '* English: It places the chains of same size, with empty spaces.
 '* Español: Coloca las cadenas de igual tamaño, con espacios vacíos.
 If (Len(BigText) > Len(isText)) Then
  isH = Len(BigText) - Len(isText)
  For isFor = 1 To isH
   SameSize = SameSize & isChar
  Next
  SameSize = SameSize & isText
 Else
  SameSize = isText
 End If
End Function

Private Sub SetInfo(ByVal i As Long, ByVal j As Long)
 Dim Temp As Variant
 
 '* English: Reorders the values.
 '* Español: Reordena los valores.
 Temp = OrderListContents(i).Color
 OrderListContents(i).Color = OrderListContents(j).Color
 OrderListContents(j).Color = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Enabled
 OrderListContents(i).Enabled = OrderListContents(j).Enabled
 OrderListContents(j).Enabled = Temp
 '*******************************************************************'
 Set Temp = OrderListContents(i).Image
 Set OrderListContents(i).Image = OrderListContents(j).Image
 Set OrderListContents(j).Image = Temp
 '*******************************************************************'
 Set Temp = OrderListContents(i).MouseIcon
 Set OrderListContents(i).MouseIcon = OrderListContents(j).MouseIcon
 Set OrderListContents(j).MouseIcon = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).SeparatorLine
 OrderListContents(i).SeparatorLine = OrderListContents(j).SeparatorLine
 OrderListContents(j).SeparatorLine = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Tag
 OrderListContents(i).Tag = OrderListContents(j).Tag
 OrderListContents(j).Tag = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).Text
 OrderListContents(i).Text = OrderListContents(j).Text
 OrderListContents(j).Text = Temp
 '*******************************************************************'
 Temp = OrderListContents(i).ToolTipText
 OrderListContents(i).ToolTipText = OrderListContents(j).ToolTipText
 OrderListContents(j).ToolTipText = Temp
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim cRed   As Long, cBlue  As Long
 Dim Delta  As Long, cGreen As Long

 '* English: Shift a color.
 '* Español: Devuelve un Color con menos intensidad.
 cBlue = ((theColor \ &H10000) Mod &H100)
 cGreen = ((theColor \ &H100) Mod &H100)
 cRed = (theColor And &HFF)
 Delta = &HFF - Base
 cBlue = Base + cBlue * Delta \ &HFF
 cGreen = Base + cGreen * Delta \ &HFF
 cRed = Base + cRed * Delta \ &HFF
 If (cRed > 255) Then cRed = 255
 If (cGreen > 255) Then cGreen = 255
 If (cBlue > 255) Then cBlue = 255
 ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
End Function

   '********************************'
   '*  Extracted of KPD-Team 1998  *'
   '*  URL: http://www.allapi.net  *'
   '*  E-Mail: KPDTeam@Allapi.net  *'
   '********************************'
Private Function TempPathName() As String
 Dim strTemp As String
 
 '* English: Returns the name of the temporary directory of Windows.
 '* Español: Devuelve el nombre del directorio temporal de Windows.
 strTemp = String$(100, Chr$(0)) '* Create a buffer.
 Call GetTempPath(100, strTemp)  '* Get the temporary path.
 '* Strip the rest of the buffer.
 TempPathName = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
End Function
