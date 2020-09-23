VERSION 5.00
Begin VB.UserControl SOptionButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ToolboxBitmap   =   "SOptionButton.ctx":0000
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "SOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2004        *'
'*******************************************************'
'*                   Version 1.0.0                     *'
'*******************************************************'
'* Control:       SOptionButton                        *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Description:   This usercontrol simulates a Option  *'
'*                Button.                              *'
'*                                                     *'
'*                Also many thanks to Paul Caton for   *'
'*                it's spectacular self-subclassing    *'
'*                usercontrol template, please see     *'
'*                the [CodeId = 54117].                *'
'*                                                     *'
'*                Richard Mewett, for the Unicode      *'
'*                support routines.                    *'
'*******************************************************'
'* Started on:    Friday, 15-oct-2004.                 *'
'*******************************************************'
'*                   Version 1.0.0                     *'
'*                                                     *'
'* Enhancements:  - Only Style.            (15/10/04)  *'
'*******************************************************'
'* Release date:  Sunday, 17-oct-2004.                 *'
'*******************************************************'
'*                                                     *'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*        Please rate my work on this control.         *'
'*    That lives the Soccer and the América of Cali    *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2004        *'
'*******************************************************'
Option Explicit
 
 '* Declares for Unicode support.
 Private Const VER_PLATFORM_WIN32_NT = 2
 
 Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
 End Type
 
 Private mWindowsNT   As Boolean
 
 '*******************************************************'
 '* Subclasser Declarations Paul Caton                  *'
  
 Private Const ALL_MESSAGES          As Long = -1
 Private Const GMEM_FIXED            As Long = 0
 Private Const GWL_WNDPROC           As Long = -4
 Private Const PATCH_04              As Long = 88
 Private Const PATCH_05              As Long = 93
 Private Const PATCH_08              As Long = 132
 Private Const PATCH_09              As Long = 137
 Private Const WM_MOUSEMOVE          As Long = &H200
 Private Const WM_MOUSELEAVE         As Long = &H2A3
 Private Const WM_MOVING             As Long = &H216
 Private Const WM_SIZING             As Long = &H214
 Private Const WM_EXITSIZEMOVE       As Long = &H232
 
 Private Type tSubData
  hWnd                               As Long
  nAddrSub                           As Long
  nAddrOrig                          As Long
  nMsgCntA                           As Long
  nMsgCntB                           As Long
  aMsgTblA()                         As Long
  aMsgTblB()                         As Long
 End Type

 Private Enum eMsgWhen
  MSG_AFTER = 1
  MSG_BEFORE = 2
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
 End Enum
 
 Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
 End Enum

 Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hWndTrack                          As Long
  dwHoverTime                        As Long
 End Type

 Private bTrack                      As Boolean
 Private bTrackUser32                As Boolean
 Private bInCtrl                     As Boolean
 Private sc_aSubData()               As tSubData

 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
 Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
 Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
 Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
 Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
 Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 
 Public Event MouseEnter()
 Public Event MouseLeave()
'*******************************************************'

 Public Enum AlignTextOption
  AlignLeft = &H0
  AlignRight = &H1
 End Enum
 
 '****************************'
 '* English: Private Type.   *'
 '* Español: Tipos Privados. *'
 '****************************'
 
 Private Type POINTAPI
  X            As Long
  Y            As Long
 End Type
 
 Private Type RECT
  Left         As Long
  Top          As Long
  Right        As Long
  Bottom       As Long
 End Type
 
 Private Type RGBQUAD
  rgbBlue      As Byte
  rgbGreen     As Byte
  rgbRed       As Byte
  rgbReserved  As Byte
 End Type
  
 '***************************************'
 '* English: Constant declares.         *'
 '* Español: Declaración de Constantes. *'
 '***************************************'
 Private Const COLOR_GRAYTEXT = 17
 Private Const defBackColor = &H8000000F
 Private Const defBorderColor = vbHighlight
 Private Const DC_TEXT = &H8
 Private Const Version As String = "SOptionButton 1.0.0 By HACKPRO TM"

 '********************************'
 '* English: Private variables.  *'
 '* Español: Variables privadas. *'
 '********************************'
 Private bChecked          As Boolean
 Private ControlEnabled    As Boolean
 Private FirstColor        As OLE_COLOR
 Private hasFocus          As Boolean
 Private g_Font            As StdFont
 Private LastColor         As OLE_COLOR
 Private m_btnRect         As RECT
 Private m_lCaption        As String
 Private m_StateG          As Integer
 Private myAlignOption     As AlignTextOption
 Private myBackColor       As OLE_COLOR
 Private myBorderColor     As OLE_COLOR
 Private myObj             As Object
 Private TheFocusColor     As OLE_COLOR
 Private TheForeColor      As OLE_COLOR
 Private TheRoundColor     As OLE_COLOR
 
 '******************************'
 '* English: Public Events.    *'
 '* Español: Eventos Públicos. *'
 '******************************'
 Public Event Click()
 
 '**********************************'
 '* English: Calls to the API's.   *'
 '* Español: Llamadas a los API's. *'
 '**********************************'
 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
 'Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'* ========================================================================================================
'*  Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'* ========================================================================================================
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
 '* Parameters:
 '*  bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
 '*  bHandled - Set this variable to True in a before callback to prevent the message being subsequently processed by the default handler... and if set, an after callback
 '*  lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
 '*  hWnd     - The window handle
 '*  uMsg     - The message number
 '*  wParam   - Message related data
 '*  lParam   - Message related data
 '* Notes:
 '*  If you really know what youre doing, its possible to change the values of the _
     hWnd, uMsg, wParam and lParam parameters in a before callback so that different _
     values get passed to the default handler.. and optionaly, the after callback.
 Select Case uMsg
  Case WM_MOUSEMOVE
   If Not (bInCtrl = True) Then
    bInCtrl = True
    Call TrackMouseLeave(lng_hWnd)
    Call DrawAppearance(2, bChecked)
    RaiseEvent MouseEnter
   End If
  Case WM_MOUSELEAVE
   bInCtrl = False
   Call DrawAppearance(1, bChecked)
   RaiseEvent MouseLeave
 End Select
End Sub

'*******************************************'
'* English: Properties of the Usercontrol. *'
'* Español: Propiedades del Usercontrol.   *'
'*******************************************'
Public Property Get Alignment() As AlignTextOption
 Alignment = myAlignOption
End Property

Public Property Let Alignment(ByVal New_Align As AlignTextOption)
 Dim isAlign As Integer
 
 myAlignOption = New_Align
 Call PropertyChanged("Alignment")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = myBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
 myBackColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("BackColor")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get BorderColor() As OLE_COLOR
 BorderColor = myBorderColor
End Property

Public Property Let BorderColor(ByVal New_Color As OLE_COLOR)
 myBorderColor = ConvertSystemColor(New_Color)
 Call PropertyChanged("BorderColor")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get Caption() As String
 Caption = m_lCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
 m_lCaption = New_Caption
 Call PropertyChanged("Caption")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get Enabled() As Boolean
 Enabled = ControlEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 UserControl.Enabled = New_Enabled
 ControlEnabled = New_Enabled
 Call PropertyChanged("Enabled")
 If (ControlEnabled = False) Then
  Call DrawAppearance(-1, bChecked)
 Else
  Call DrawAppearance(1, bChecked)
 End If
End Property

Public Property Get FocusColor() As OLE_COLOR
 FocusColor = TheFocusColor
End Property

Public Property Let FocusColor(ByVal NewColor As OLE_COLOR)
 TheFocusColor = ConvertSystemColor(NewColor)
 Call PropertyChanged("FocusColor")
End Property

Public Property Get Font() As StdFont
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
 Call PropertyChanged("Font")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get ForeColor() As OLE_COLOR
 ForeColor = TheForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
 TheForeColor = ConvertSystemColor(NewColor)
 Call PropertyChanged("ForeColor")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get GetControlVersion() As String
 GetControlVersion = Version & " © " & Year(Now)
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_MemberFlags = "400"
 hWnd = UserControl.hWnd
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal MouseIcon As StdPicture)
 Set UserControl.MouseIcon = MouseIcon
 Call PropertyChanged("MouseIcon")
End Property

Public Property Get MousePointer() As MousePointerConstants
 MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal MousePointer As MousePointerConstants)
 UserControl.MousePointer = MousePointer
 Call PropertyChanged("MousePointer")
End Property

Public Property Get RoundColor() As OLE_COLOR
 RoundColor = TheRoundColor
End Property

Public Property Let RoundColor(ByVal NewColor As OLE_COLOR)
 TheRoundColor = ConvertSystemColor(NewColor)
 Call PropertyChanged("RoundColor")
 Call DrawAppearance(m_StateG, bChecked)
End Property

Public Property Get Value() As Boolean
 Value = bChecked
End Property

Public Property Let Value(ByVal lChecked As Boolean)
 Call CheckAllValue(False)
 bChecked = lChecked
 Call PropertyChanged("Value")
 Call DrawAppearance(m_StateG, bChecked)
End Property

'********************************************************'
'* English: Subs and Functions of the Usercontrol.      *'
'* Español: Procedimientos y Funciones del Usercontrol. *'
'********************************************************'
Private Sub APILine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT   As POINTAPI, hPenOld As Long
 Dim hPen As Long
 
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(picButton.hDC, hPen)
 Call MoveToEx(picButton.hDC, x1, y1, PT)
 Call LineTo(picButton.hDC, x2, y2)
 Call SelectObject(picButton.hDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Sub CheckAllValue(ByVal isValue As Boolean)
 For Each myObj In Parent.Controls
  If (TypeOf myObj Is SOptionButton) Then
   If Not (myObj.Container Is UserControl.Parent) Then
    If (myObj.hWnd = UserControl.hWnd) Then
     Call CheckContainerControls(myObj.Container, isValue)
     Exit Sub
    End If
   End If
  End If
 Next
 Call CheckContainerControls(UserControl.Parent, False)
End Sub

Private Sub CheckContainerControls(ByVal cContainer As Object, ByVal ctlValue As Boolean)
 For Each myObj In Parent.Controls
  If (TypeOf myObj Is SOptionButton) Then
   If (myObj.Container Is cContainer) Then
    If Not (myObj.hWnd = UserControl.hWnd) Then
     If (myObj.Value = True) Then myObj.Value = ctlValue
    End If
   End If
  End If
 Next
End Sub

Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function

Private Sub DrawAppearance(Optional ByVal m_State As Integer = 1, Optional ByVal bChecked As Boolean = False)
 Dim isColor As OLE_COLOR, i As Integer, TheForeCol As OLE_COLOR
 
On Error Resume Next
 UserControl.Cls
 Set UserControl.Font = g_Font
 m_StateG = IIf(m_State = 0, 1, m_State)
 UserControl.BackColor = myBackColor
 'UserControl.Height = 240
 UserControl.Height = UserControl.TextHeight("Qr") * Screen.TwipsPerPixelY + 35
 With picButton
  .Cls
  .Top = (UserControl.ScaleHeight / 2) - 8
  If (myAlignOption = 0) Then
   .Left = 0
  Else
   .Left = UserControl.ScaleWidth - 16
  End If
  .BackColor = myBackColor
 End With
 If (m_StateG = 1) Then
  isColor = ConvertSystemColor(myBorderColor)
  TheForeCol = TheForeColor
 ElseIf (m_StateG = 2) Then
  isColor = ShiftColorOXP(myBorderColor, 38)
  TheForeCol = TheFocusColor
 ElseIf (m_StateG = 3) Then
  isColor = ShiftColorOXP(myBorderColor, -53)
  TheForeCol = TheForeColor
 Else
  isColor = &HE5ECEC
  TheForeCol = TheForeColor
 End If
 Call DrawGradient(ShiftColorOXP(isColor, 43), ShiftColorOXP(isColor, 143), 4, 2, 17)
 For i = 2 To picButton.ScaleHeight
  Call APILine(14, i, 25, i, myBackColor)
 Next
 For i = 13 To picButton.ScaleHeight
  Call APILine(1, i, 19, i, myBackColor)
 Next
 Call APILine(5, 4, 4, 5, &HFFFFFF)
 Call APILine(3, 3, 5, 6, &HFFFFFF)
 Call APILine(3, 2, 5, 2, myBackColor)
 Call APILine(12, 2, 16, 2, myBackColor)
 Call APILine(3, 3, 4, 3, myBackColor)
 Call APILine(13, 3, 14, 3, myBackColor)
 Call APILine(7, 1, 10, 1, ShiftColorOXP(FirstColor, &H70))
 Call APILine(6, 2, 11, 2, isColor)
 Call APILine(5, 3, 4, 3, isColor)
 Call APILine(11, 3, 10, 3, isColor)
 Call APILine(4, 4, 5, 4, isColor)
 Call APILine(12, 4, 11, 4, isColor)
 Call APILine(3, 5, 4, 5, isColor)
 Call APILine(13, 5, 12, 5, isColor)
 Call APILine(3, 6, 3, 9, isColor)
 Call APILine(13, 6, 13, 9, isColor)
 Call APILine(3, 9, 4, 9, isColor)
 Call APILine(13, 9, 12, 9, isColor)
 Call APILine(4, 10, 5, 10, isColor)
 Call APILine(12, 10, 11, 10, isColor)
 Call APILine(5, 11, 4, 11, isColor)
 Call APILine(3, 11, 4, 11, myBackColor)
 Call APILine(11, 11, 10, 11, isColor)
 Call APILine(13, 11, 12, 11, myBackColor)
 Call APILine(6, 12, 11, 12, isColor)
 Call APILine(7, 13, 10, 13, ShiftColorOXP(LastColor, &H70))
 Call APILine(3, 12, 5, 12, myBackColor)
 Call APILine(12, 12, 14, 12, myBackColor)
 If (bChecked = True) Then
  isColor = IIf(m_State = -1, &HE5ECEC, ConvertSystemColor(TheRoundColor))
  Call APILine(8, 4, 7, 4, isColor)
  Call APILine(7, 4, 6, 4, ShiftColorOXP(isColor, 56))
  Call APILine(9, 4, 8, 4, ShiftColorOXP(isColor, 56))
  Call APILine(6, 5, 11, 5, isColor)
  Call APILine(6, 6, 11, 6, isColor)
  Call APILine(5, 6, 4, 6, ShiftColorOXP(isColor, 36))
  Call APILine(11, 6, 12, 6, ShiftColorOXP(isColor, 36))
  Call APILine(5, 7, 12, 7, isColor)
  Call APILine(6, 8, 11, 8, isColor)
  Call APILine(5, 8, 4, 8, ShiftColorOXP(isColor, 66))
  Call APILine(11, 8, 12, 8, ShiftColorOXP(isColor, 66))
  Call APILine(6, 9, 11, 9, isColor)
  Call APILine(8, 10, 7, 10, isColor)
  Call APILine(7, 10, 6, 10, ShiftColorOXP(isColor, 76))
  Call APILine(9, 10, 8, 10, ShiftColorOXP(isColor, 76))
 End If
 Call DrawCaption(m_lCaption, TheForeCol)
 If (hasFocus = True) Then
  isColor = OffSetColor(myBorderColor, -&H45)
  m_btnRect.Left = 19
  m_btnRect.Right = UserControl.ScaleWidth - 1
  If (myAlignOption = 0) Then
   For i = m_btnRect.Left - 2 To m_btnRect.Right Step 2
    Call SetPixel(UserControl.hDC, i, m_btnRect.Top, isColor)
   Next
   For i = m_btnRect.Left - 2 To m_btnRect.Right Step 2
    Call SetPixel(UserControl.hDC, i, m_btnRect.Top + UserControl.TextHeight(m_lCaption), isColor)
   Next
   For i = m_btnRect.Top To m_btnRect.Top + UserControl.TextHeight(m_lCaption) Step 2
    Call SetPixel(UserControl.hDC, m_btnRect.Left - 2, i, isColor)
   Next
   For i = m_btnRect.Top To m_btnRect.Top + UserControl.TextHeight(m_lCaption) Step 2
    Call SetPixel(UserControl.hDC, m_btnRect.Right, i, isColor)
   Next
  Else
   For i = 0 To m_btnRect.Right - 17 Step 2
    Call SetPixel(UserControl.hDC, i, m_btnRect.Top, isColor)
   Next
   For i = 0 To m_btnRect.Right - 17 Step 2
    Call SetPixel(UserControl.hDC, i, m_btnRect.Top + UserControl.TextHeight(m_lCaption), isColor)
   Next
   For i = m_btnRect.Top To m_btnRect.Top + UserControl.TextHeight(m_lCaption) Step 2
    Call SetPixel(UserControl.hDC, 0, i, isColor)
   Next
   For i = m_btnRect.Top To m_btnRect.Top + UserControl.TextHeight(m_lCaption) Step 2
    Call SetPixel(UserControl.hDC, m_btnRect.Right - 16, i, isColor)
   Next
  End If
 End If
End Sub

Private Sub DrawCaption(ByVal lCaption As String, Optional ByVal lColor As OLE_COLOR = &HF0)
 If (ControlEnabled = False) Then lColor = GetSysColor(COLOR_GRAYTEXT)
 Call SetTextColor(UserControl.hDC, lColor)
 m_btnRect.Bottom = UserControl.ScaleHeight
 m_btnRect.Top = 1
 If (myAlignOption = 0) Then
  m_btnRect.Left = 18
 Else
  m_btnRect.Left = 0
 End If
 m_btnRect.Right = UserControl.ScaleWidth
 '*************************************************************************
 '* Draws the text with Unicode support based on OS version.              *
 '* Thanks to Richard Mewett.                                             *
 '*************************************************************************
 If (mWindowsNT = True) Then
  Call DrawTextW(UserControl.hDC, StrPtr(lCaption), Len(lCaption), m_btnRect, DC_TEXT)
 Else
  Call DrawTextA(UserControl.hDC, lCaption, Len(lCaption), m_btnRect, DC_TEXT)
 End If
End Sub

Private Sub DrawGradient(ByVal LngColor1 As Long, ByVal LngColor2 As Long, ByVal Y As Long, ByVal x2 As Long, ByVal y2 As Long)
 Dim RgbColor1 As RGBQUAD, RgbColor2 As RGBQUAD, isColor  As OLE_COLOR
 Dim ColorRojo As Double, ColorVerde As Double, ColorAzul As Double
 Dim CDiffRed  As Double, CDiffGreen As Double, CDiffBlue As Double
 Dim CFadeRed  As Double, CFadeGreen As Double, CFadeBlue As Double
 Dim Fade      As Double
  
 Call Long2RGB(LngColor1, RgbColor1)
 Call Long2RGB(LngColor2, RgbColor2)
 CDiffRed = -(CLng(RgbColor1.rgbRed) - CLng(RgbColor2.rgbRed))
 CDiffGreen = -(CLng(RgbColor1.rgbGreen) - CLng(RgbColor2.rgbGreen))
 CDiffBlue = -(CLng(RgbColor1.rgbBlue) - CLng(RgbColor2.rgbBlue))
 ColorRojo = RgbColor1.rgbRed
 ColorVerde = RgbColor1.rgbGreen
 ColorAzul = RgbColor1.rgbBlue
 CFadeRed = CDiffRed / y2
 CFadeGreen = CDiffGreen / y2
 CFadeBlue = CDiffBlue / y2
 For Fade = 0 To y2
  Call APILine(Fade + Y, x2, x2, Fade + Y, isColor)
  ColorRojo = ColorRojo + CFadeRed
  ColorVerde = ColorVerde + CFadeGreen
  ColorAzul = ColorAzul + CFadeBlue + Fade
  isColor = ShiftColorOXP(RGB(ColorRojo, ColorVerde, ColorAzul), Fade * 2)
  If (Fade = 10) Then FirstColor = isColor
 Next
 LastColor = isColor
End Sub

Private Sub Long2RGB(ByVal LngColor As Long, ByRef RGBColor As RGBQUAD)
 Dim Aux As Byte
 
 Call CopyMemory(RGBColor, LngColor, 4)
 Aux = RGBColor.rgbBlue
 RGBColor.rgbBlue = RGBColor.rgbRed
 RGBColor.rgbRed = Aux
End Sub

Private Function OffSetColor(ByVal lColor As OLE_COLOR, ByVal lOffset As Long) As OLE_COLOR
 Dim lRed  As OLE_COLOR, lGreen As OLE_COLOR
 Dim lBlue As OLE_COLOR, lR     As OLE_COLOR
 Dim lG    As OLE_COLOR, lB     As OLE_COLOR
   
 lR = (lColor And &HFF)
 lG = ((lColor And 65280) \ 256)
 lB = ((lColor) And 16711680) \ 65536
 lRed = (lOffset + lR)
 lGreen = (lOffset + lG)
 lBlue = (lOffset + lB)
 If (lRed > 255) Then lRed = 255
 If (lRed < 0) Then lRed = 0
 If (lGreen > 255) Then lGreen = 255
 If (lGreen < 0) Then lGreen = 0
 If (lBlue > 255) Then lBlue = 255
 If (lBlue < 0) Then lBlue = 0
 OffSetColor = RGB(lRed, lGreen, lBlue)
End Function

Private Sub SetAccessKeys()
 Dim AmperSandPos As Long

 UserControl.AccessKeys = ""
 If (Len(Caption) > 1) Then
  AmperSandPos = InStr(1, Caption, "&", vbTextCompare)
  If (AmperSandPos < Len(Caption)) And (AmperSandPos > 0) Then
   If (Mid$(Caption, AmperSandPos + 1, 1) <> "&") Then
    UserControl.AccessKeys = LCase$(Mid$(Caption, AmperSandPos + 1, 1))
   Else
    AmperSandPos = InStr(AmperSandPos + 2, Caption, "&", vbTextCompare)
    If (Mid$(Caption, AmperSandPos + 1, 1) <> "&") Then
     UserControl.AccessKeys = LCase$(Mid$(Caption, AmperSandPos + 1, 1))
    End If
   End If
  End If
 End If
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim cRed   As Long, cBlue  As Long
 Dim Delta  As Long, cGreen As Long

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

Private Sub picButton_GotFocus()
 Call UserControl_GotFocus
End Sub

Private Sub picButton_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = vbKeySpace) Then
  Call CheckAllValue(False)
  Call UserControl_MouseUp(1, 0, 0, 0)
 End If
End Sub

Private Sub picButton_LostFocus()
 Call UserControl_LostFocus
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picButton_Paint()
 Call DrawAppearance(m_StateG, bChecked)
End Sub

Private Sub UserControl_GotFocus()
 hasFocus = True
 Call DrawAppearance(m_StateG, bChecked)
End Sub

Private Sub UserControl_Initialize()
 Dim OS As OSVERSIONINFO

 '* Get the operating system version for text drawing purposes.
 OS.dwOSVersionInfoSize = Len(OS)
 Call GetVersionEx(OS)
 mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub

Private Sub UserControl_InitProperties()
 myAlignOption = 0
 ControlEnabled = True
 ForeColor = &H80000012
 FocusColor = &H80000012
 m_lCaption = Ambient.DisplayName
 m_StateG = 1
 myBackColor = ConvertSystemColor(defBackColor)
 myBorderColor = ConvertSystemColor(defBorderColor)
 Set g_Font = Ambient.Font
 TheRoundColor = vb3DHighlight
 Value = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 If (KeyCode = vbKeySpace) Then Call UserControl_MouseUp(1, 0, 0, 0)
End Sub

Private Sub UserControl_LostFocus()
 hasFocus = False
 Call DrawAppearance(m_StateG, bChecked)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (bChecked = True) Or (ControlEnabled = False) Or Not (Button = vbLeftButton) Then Exit Sub
 Call CheckAllValue(False)
 hasFocus = True
 bChecked = True
 Call DrawAppearance(3, bChecked)
 RaiseEvent Click
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (Button <> vbLeftButton) Then Exit Sub
 hasFocus = True
 bChecked = True
 Call DrawAppearance(2, bChecked)
End Sub

Private Sub UserControl_Paint()
 Call DrawAppearance(m_StateG, bChecked)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Alignment = PropBag.ReadProperty("Alignment", 0)
 BackColor = PropBag.ReadProperty("BackColor", ConvertSystemColor(defBackColor))
 BorderColor = PropBag.ReadProperty("BorderColor", ConvertSystemColor(defBorderColor))
 m_lCaption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
 Enabled = PropBag.ReadProperty("Enabled", True)
 FocusColor = PropBag.ReadProperty("FocusColor", &H80000012)
 Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
 ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
 UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
 Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 RoundColor = PropBag.ReadProperty("RoundColor", vb3DHighlight)
 Value = PropBag.ReadProperty("Value", False)
 Call SetAccessKeys
 If (Ambient.UserMode = True) Then
  bTrack = True
  bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  If Not (bTrackUser32 = True) Then
   If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
    bTrack = False
   End If
  End If
  If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
   '* Start subclassing the UserControl.
   Call Subclass_Start(hWnd)
   Call Subclass_AddMsg(hWnd, WM_MOUSEMOVE, MSG_AFTER)
   Call Subclass_AddMsg(hWnd, WM_MOUSELEAVE, MSG_AFTER)
   Call Subclass_Start(picButton.hWnd)
   Call Subclass_AddMsg(picButton.hWnd, WM_MOUSEMOVE, MSG_AFTER)
   Call Subclass_AddMsg(picButton.hWnd, WM_MOUSELEAVE, MSG_AFTER)
  End If
 End If
End Sub

Private Sub UserControl_Resize()
 If (Ambient.UserMode = False) Then Call DrawAppearance(m_StateG, bChecked)
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
 Call Subclass_StopAll '* Stop all subclassing.
 Exit Sub
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Alignment", myAlignOption, 0)
 Call PropBag.WriteProperty("BackColor", myBackColor, ConvertSystemColor(defBackColor))
 Call PropBag.WriteProperty("BorderColor", myBorderColor, ConvertSystemColor(defBorderColor))
 Call PropBag.WriteProperty("Caption", m_lCaption, Ambient.DisplayName)
 Call PropBag.WriteProperty("Enabled", ControlEnabled, True)
 Call PropBag.WriteProperty("FocusColor", TheFocusColor, &H80000012)
 Call PropBag.WriteProperty("Font", g_Font, Ambient.Font)
 Call PropBag.WriteProperty("ForeColor", TheForeColor, &H80000012)
 Call PropBag.WriteProperty("MousePointer", MousePointer, vbDefault)
 Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
 Call PropBag.WriteProperty("RoundColor", TheRoundColor, vb3DHighlight)
 Call PropBag.WriteProperty("Value", bChecked, False)
End Sub

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
 Dim hMod As Long, bLibLoaded As Boolean
 
 hMod = GetModuleHandleA(sModule)
 If (hMod = 0) Then
  hMod = LoadLibraryA(sModule)
  If (hMod) Then bLibLoaded = True
 End If
 If (hMod) Then
  If (GetProcAddress(hMod, sFunction)) Then IsFunctionExported = True
 End If
 If (bLibLoaded = True) Then Call FreeLibrary(hMod)
End Function

'* Track the mouse leaving the indicated window.
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
 Dim tme As TRACKMOUSEEVENT_STRUCT
 
 If (bTrack = True) Then
  With tme
   .cbSize = Len(tme)
   .dwFlags = TME_LEAVE
   .hWndTrack = lng_hWnd
  End With
  If (bTrackUser32 = True) Then
   Call TrackMouseEvent(tme)
  Else
   Call TrackMouseEventComCtl(tme)
  End If
 End If
End Sub

'* ============================================================================================================================
'*  Subclass code - The programmer may call any of the following Subclass_??? routines
'*  Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'* ============================================================================================================================
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
 '* Parameters:
 '*  lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
 '*  uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
 '*  When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
 With sc_aSubData(zIdx(lng_hWnd))
  If (When) And (eMsgWhen.MSG_BEFORE) Then
   Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
  End If
  If (When) And (eMsgWhen.MSG_AFTER) Then
   Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
  End If
 End With
End Sub

'* Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
 '* Parameters:
 '*  lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
 '*  uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
 '*  When      - Whether the msg is to be removed from the before, after or both callback tables
 With sc_aSubData(zIdx(lng_hWnd))
  If (When) And (eMsgWhen.MSG_BEFORE) Then
   Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
  End If
  If (When) And (eMsgWhen.MSG_AFTER) Then
   Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
  End If
 End With
End Sub

'* Return whether were running in the IDE.
Private Function Subclass_InIDE() As Boolean
 Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'* Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
 '* Parameters:
 '*  lng_hWnd  - The handle of the window to be subclassed.
 '*  Returns;
 '*  The sc_aSubData() index.
 Const CODE_LEN              As Long = 200
 Const FUNC_CWP              As String = "CallWindowProcA"
 Const FUNC_EBM              As String = "EbMode"
 Const FUNC_SWL              As String = "SetWindowLongA"
 Const MOD_USER              As String = "user32"
 Const MOD_VBA5              As String = "vba5"
 Const MOD_VBA6              As String = "vba6"
 Const PATCH_01              As Long = 18
 Const PATCH_02              As Long = 68
 Const PATCH_03              As Long = 78
 Const PATCH_06              As Long = 116
 Const PATCH_07              As Long = 121
 Const PATCH_0A              As Long = 186
 Static aBuf(1 To CODE_LEN)  As Byte
 Static pCWP                 As Long
 Static pEbMode              As Long
 Static pSWL                 As Long
 Dim i                       As Long
 Dim j                       As Long
 Dim nSubIdx                 As Long
 Dim sHex                    As String
 
 '* If its the first time through here..
 If (aBuf(1) = 0) Then
  '* The hex pair machine code representation.
  sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
  '* Convert the string from hex pairs to bytes and store in the static machine code buffer
  i = 1
  Do While (j < CODE_LEN)
   j = j + 1
   aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
   i = i + 2
  Loop
  '* Get API function addresses.
  If (Subclass_InIDE = True) Then
   aBuf(16) = &H90
   aBuf(17) = &H90
   pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
   If (pEbMode = 0) Then pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
  End If
  pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
  pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
  ReDim sc_aSubData(0 To 0) As tSubData
 Else
  nSubIdx = zIdx(lng_hWnd, True)
  If (nSubIdx = -1) Then
   nSubIdx = UBound(sc_aSubData()) + 1
   ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
  End If
  Subclass_Start = nSubIdx
 End If
 With sc_aSubData(nSubIdx)
  .hWnd = lng_hWnd
  .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
  .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
  Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
  Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
  Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
  Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
  Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
  Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
  Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
 End With
End Function

'* Stop all subclassing
Private Sub Subclass_StopAll()
 Dim i As Long
 
On Error GoTo myErr
 i = UBound(sc_aSubData())
 Do While (i >= 0)
  With sc_aSubData(i)
   If (.hWnd <> 0) Then Call Subclass_Stop(.hWnd)
  End With
  i = i - 1
 Loop
 Exit Sub
myErr:
End Sub

'* Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
 '* Parameters:
 '*  lng_hWnd  - The handle of the window to stop being subclassed
 With sc_aSubData(zIdx(lng_hWnd))
  Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
  Call zPatchVal(.nAddrSub, PATCH_05, 0)
  Call zPatchVal(.nAddrSub, PATCH_09, 0)
  Call GlobalFree(.nAddrSub)
  .hWnd = 0
  .nMsgCntB = 0
  .nMsgCntA = 0
  Erase .aMsgTblB
  Erase .aMsgTblA
 End With
End Sub

'* ======================================================================================================
'*  These z??? routines are exclusively called by the Subclass_??? routines.
'*  Worker sub for Subclass_AddMsg
'* ======================================================================================================
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
 Dim nEntry As Long, nOff1 As Long, nOff2 As Long
 
 If (uMsg = ALL_MESSAGES) Then
  nMsgCnt = ALL_MESSAGES
 Else
  Do While (nEntry < nMsgCnt)
   nEntry = nEntry + 1
   If (aMsgTbl(nEntry) = 0) Then
    aMsgTbl(nEntry) = uMsg
    Exit Sub
   ElseIf (aMsgTbl(nEntry) = uMsg) Then
    Exit Sub
   End If
  Loop
  nMsgCnt = nMsgCnt + 1
  ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
  aMsgTbl(nMsgCnt) = uMsg
 End If
 If (When = eMsgWhen.MSG_BEFORE) Then
  nOff1 = PATCH_04
  nOff2 = PATCH_05
 Else
  nOff1 = PATCH_08
  nOff2 = PATCH_09
 End If
 If (uMsg <> ALL_MESSAGES) Then Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
 Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

'* Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
 zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
 Debug.Assert zAddrFunc
End Function

'* Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
 Dim nEntry As Long
 
 If (uMsg = ALL_MESSAGES) Then
  nMsgCnt = 0
  If (When = eMsgWhen.MSG_BEFORE) Then
   nEntry = PATCH_05
  Else
   nEntry = PATCH_09
  End If
  Call zPatchVal(nAddr, nEntry, 0)
 Else
  Do While (nEntry < nMsgCnt)
   nEntry = nEntry + 1
   If aMsgTbl(nEntry) = uMsg Then
    aMsgTbl(nEntry) = 0
    Exit Do
   End If
  Loop
 End If
End Sub

'* Get the sc_aSubData() array index of the passed hWnd.
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
 '* Get the upper bound of sc_aSubData() - If you get an error here, youre probably Subclass_AddMsg-ing before Subclass_Start.
 zIdx = UBound(sc_aSubData)
 Do While (zIdx >= 0)
  With sc_aSubData(zIdx)
   If (.hWnd = lng_hWnd) And Not (bAdd = True) Then
    Exit Function
   ElseIf (.hWnd = 0) And (bAdd = True) Then
    Exit Function
   End If
  End With
  zIdx = zIdx - 1
 Loop
 If Not (bAdd = True) Then Debug.Assert False
 '* If we exit here, were returning -1, no freed elements were found.
End Function

'* Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
 Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'* Patch the machine code buffer at the indicated offset with the passed value.
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
 Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'* Worker function for Subclass_InIDE.
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
 zSetTrue = True
 bValue = True
End Function
