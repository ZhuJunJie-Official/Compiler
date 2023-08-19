VERSION 5.00
Begin VB.UserControl GpTabStrip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   ClipBehavior    =   0  '��
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   PropertyPages   =   "GpTabs.ctx":0000
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   149
   ToolboxBitmap   =   "GpTabs.ctx":0014
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "GpTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Option Explicit
'

Private Const MODULE_NAME = "GpTabs"
Private Const XPBorderColor = &H808080       '&H733C00  ' XP���߿����ɫ
Private Const XPFlatBorderColor = &HE0A193
Private Const XPFlatTabColor = &HFFC3B3
'Private Const XPFlatTabColorActive = vbWhite
Private Const XPFlatTabColorActive = &HEAEAEA
Private Const XPFlatTabColorHover = &HFFA49B
Private Const TabsInterval = 2          ' ÿ��Tab֮��ļ������
Private Const RoundRectSize = 1         ' Բ�ǵĴ�С
Private Const DiscrepancyHeight = 2     ' ѡ�е�Tab��û��ѡ�е�Tab�ĸ߶Ȳ�
Private Const InflateFontHeight = 6     ' ��Tab��Caption�ڵ�ǰ�����ʵ�ʸ߶���ӵĵ�Tab��Ĭ�ϸ߶�
Private Const InflateFontWidth = 4      ' ��Tab��Caption�ڵ�ǰ�����ʵ�ʿ����ӵĵ�Tab��Ĭ�Ͽ��
Private Const InflateIconHeight = 2     ' ��Tab��Icon��ʵ�ʸ߶���ӵĵ�Tab��Ĭ�ϸ߶�
Private Const InflateIconWidth = 0      ' ��Tab��Icon��ʵ�ʿ����ӵĵ�Tab��Ĭ�Ͽ��

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

#Const DEBUGMODE = 0

'Set of bit flags that indicate which common control classes will be loaded
'from the DLL. The dwICC value of tagINITCOMMONCONTROLSEX can
'be a combination of the following:
 Const ICC_LISTVIEW_CLASSES = &H1          '/* listview, header
 Const ICC_TREEVIEW_CLASSES = &H2          '/* treeview, tooltips
 Const ICC_BAR_CLASSES = &H4               '/* toolbar, statusbar, trackbar, tooltips
 Const ICC_TAB_CLASSES = &H8               '/* tab, tooltips
 Const ICC_UPDOWN_CLASS = &H10             '/* updown
 Const ICC_PROGRESS_CLASS = &H20           '/* progress
 Const ICC_HOTKEY_CLASS = &H40             '/* hotkey
 Const ICC_ANIMATE_CLASS = &H80            '/* animate
 Const ICC_WIN95_CLASSES = &HFF            '/* loads everything above
 Const ICC_DATE_CLASSES = &H100            '/* month picker, date picker, time picker, updown
 Const ICC_USEREX_CLASSES = &H200          '/* ComboEx
 Const ICC_COOL_CLASSES = &H400            '/* Rebar (coolbar) control


' ָ�����ڵĽṹ��ȡ����Ϣ������GetWindowLong��SetWindowLong����
 Const GWL_EXSTYLE = (-20)                 '/* ��չ������ʽ */
 Const GWL_HINSTANCE = (-6)                '/* ӵ�д��ڵ�ʵ���ľ�� */
 Const GWL_HWNDPARENT = (-8)               '/* �ô���֮���ľ������Ҫ��SetWindowWord���ı����ֵ */
 Const GWL_ID = (-12)                      '/* �Ի�����һ���Ӵ��ڵı�ʶ�� */
 Const GWL_STYLE = (-16)                   '/* ������ʽ */
 Const GWL_USERDATA = (-21)                '/* ������Ӧ�ó���涨 */
 Const GWL_WNDPROC = (-4)                  '/* �ô��ڵĴ��ں����ĵ�ַ */
 Const DWL_DLGPROC = 4                     '/* ������ڵĶԻ�������ַ */
 Const DWL_MSGRESULT = 0                   '/* �ڶԻ������д����һ����Ϣ���ص�ֵ */
 Const DWL_USER = 8                        '/* ������Ӧ�ó���涨 */


' GetDeviceCaps����������GetDeviceCaps����
 Const DRIVERVERSION = 0                   '/* ����������汾
 Const BITSPIXEL = 12                      '/*
 Const LOGPIXELSX = 88                     '/*  Logical pixels/inch in X
 Const LOGPIXELSY = 90                     '/*  Logical pixels/inch in Y

' Windows������������GetSysColor
 Const COLOR_ACTIVEBORDER = 10             '/* ����ڵı߿�
 Const COLOR_ACTIVECAPTION = 2             '/* ����ڵı���
 Const COLOR_ADJ_MAX = 100                 '/*
 Const COLOR_ADJ_MIN = -100                '/*
 Const COLOR_APPWORKSPACE = 12             '/* MDI����ı���
 Const COLOR_BACKGROUND = 1                '/*
 Const COLOR_BTNDKSHADOW = 21              '/*
 Const COLOR_BTNLIGHT = 22                 '/*
 Const COLOR_BTNFACE = 15                  '/* ��ť
 Const COLOR_BTNHIGHLIGHT = 20             '/* ��ť��3D������
 Const COLOR_BTNSHADOW = 16                '/* ��ť��3D��Ӱ
 Const COLOR_BTNTEXT = 18                  '/* ��ť����
 Const COLOR_CAPTIONTEXT = 9               '/* ���ڱ����е�����
 Const COLOR_GRAYTEXT = 17                 '/* ��ɫ���֣���ʹ���˶���������Ϊ��
 Const COLOR_HIGHLIGHT = 13                '/* ѡ������Ŀ����
 Const COLOR_HIGHLIGHTTEXT = 14            '/* ѡ������Ŀ����
 Const COLOR_INACTIVEBORDER = 11           '/* ������ڵı߿�
 Const COLOR_INACTIVECAPTION = 3           '/* ������ڵı���
 Const COLOR_INACTIVECAPTIONTEXT = 19      '/* ������ڵ�����
 Const COLOR_MENU = 4                      '/* �˵�
 Const COLOR_MENUTEXT = 7                  '/* �˵�����
 Const COLOR_SCROLLBAR = 0                 '/* ������
 Const COLOR_WINDOW = 5                    '/* ���ڱ���
 Const COLOR_WINDOWFRAME = 6               '/* ����
 Const COLOR_WINDOWTEXT = 8                '/* ��������
Const COLORONCOLOR = 3

' ����CombineRgn�ķ���ֵ������Long
 Const COMPLEXREGION = 3                   '/* �����л��ཻ���ı߽� */
 Const SIMPLEREGION = 2                    '/* ����߽�û�л��ཻ�� */
 Const NULLREGION = 1                      '/* ����Ϊ�� */
 Const ERRORAPI = 0                        '/* ���ܴ���������� */

' ���������ķ���������CombineRgn�ĵĲ���nCombineMode��ʹ�õĳ���
 Const RGN_AND = 1                         '/* hDestRgn������Ϊ����Դ����Ľ��� */
 Const RGN_COPY = 5                        '/* hDestRgn������ΪhSrcRgn1�Ŀ��� */
 Const RGN_DIFF = 4                        '/* hDestRgn������ΪhSrcRgn1����hSrcRgn2���ཻ�Ĳ��� */
 Const RGN_OR = 2                          '/* hDestRgn������Ϊ��������Ĳ��� */
 Const RGN_XOR = 3                         '/* hDestRgn������Ϊ������Դ����OR֮��Ĳ��� */

' Missing Draw State constants declarations���ο�DrawState����
'/* Image type */
 Const DST_COMPLEX = &H0                   '/* ��ͼ����lpDrawStateProc����ָ���Ļص������ڼ�ִ�С�lParam��wParam�ᴫ�ݸ��ص��¼�
 Const DST_TEXT = &H1                      '/* lParam�������ֵĵ�ַ����ʹ��һ���ִ���������wParam�����ִ��ĳ���
 Const DST_PREFIXTEXT = &H2                '/* ��DST_TEXT���ƣ�ֻ�� & �ַ�ָ��Ϊ�¸��ַ������»���
 Const DST_ICON = &H3                      '/* lParam����ͼ����
 Const DST_BITMAP = &H4                    '/* lParam�еľ��
' /* State type */
 Const DSS_NORMAL = &H0                    '/* ��ͨͼ��
 Const DSS_UNION = &H10                    '/* ͼ����ж�������
 Const DSS_DISABLED = &H20                 '/* ͼ����и���Ч��
 Const DSS_MONO = &H80                     '/* ��hBrush���ͼ��
 Const DSS_RIGHT = &H8000                  '/*

' Built in ImageList drawing methods:
 Const ILD_NORMAL = 0&
 Const ILD_TRANSPARENT = 1&
 Const ILD_BLEND25 = 2&
 Const ILD_SELECTED = 4&
 Const ILD_FOCUS = 4&
 Const ILD_MASK = &H10&
 Const ILD_IMAGE = &H20&
 Const ILD_ROP = &H40&
 Const ILD_OVERLAYMASK = 3840&
 Const ILC_MASK = &H1&
 Const ILCF_MOVE = &H0&
 Const ILCF_SWAP = &H1&

 Const CLR_DEFAULT = -16777216
 Const CLR_HILIGHT = -16777216
 Const CLR_NONE = -1

' General windows messages:
 Const WM_COMMAND = &H111
 Const WM_KEYDOWN = &H100
 Const WM_KEYUP = &H101
 Const WM_CHAR = &H102
 Const WM_SETFOCUS = &H7
 Const WM_KILLFOCUS = &H8
 Const WM_SETFONT = &H30
 Const WM_GETTEXT = &HD
 Const WM_GETTEXTLENGTH = &HE
 Const WM_SETTEXT = &HC
 Const WM_NOTIFY = &H4E&

' Show window styles
 Const SW_SHOWNORMAL = 1
 Const SW_ERASE = &H4
 Const SW_HIDE = 0
 Const SW_INVALIDATE = &H2
 Const SW_MAX = 10
 Const SW_MAXIMIZE = 3
 Const SW_MINIMIZE = 6
 Const SW_NORMAL = 1
 Const SW_OTHERUNZOOM = 4
 Const SW_OTHERZOOM = 2
 Const SW_PARENTCLOSING = 1
 Const SW_RESTORE = 9
 Const SW_PARENTOPENING = 3
 Const SW_SHOW = 5
 Const SW_SCROLLCHILDREN = &H1
 Const SW_SHOWDEFAULT = 10
 Const SW_SHOWMAXIMIZED = 3
 Const SW_SHOWMINIMIZED = 2
 Const SW_SHOWMINNOACTIVE = 7
 Const SW_SHOWNA = 8
 Const SW_SHOWNOACTIVATE = 4

' �����Ĺ�դ��������
 Const BLACKNESS = &H42                    '/* ��ʾʹ���������ɫ�������0��ص�ɫ�������Ŀ��������򣬣���ȱʡ�������ɫ����ԣ�����ɫΪ��ɫ����
 Const DSTINVERT = &H550009                '/* ��ʾʹĿ�����������ɫȡ����
 Const MERGECOPY = &HC000CA                '/* ��ʾʹ�ò����͵�AND���룩��������Դ�����������ɫ���ض�ģʽ���һ��
 Const MERGEPAINT = &HBB0226               '/* ͨ��ʹ�ò����͵�OR���򣩲������������Դ�����������ɫ��Ŀ������������ɫ�ϲ���
 Const NOTSRCCOPY = &H330008               '/* ��Դ����������ɫȡ�����ڿ�����Ŀ���������
 Const NOTSRCERASE = &H1100A6              '/* ʹ�ò������͵�OR���򣩲��������Դ��Ŀ������������ɫֵ��Ȼ�󽫺ϳɵ���ɫȡ����
 Const PATCOPY = &HF00021                  '/* ���ض���ģʽ������Ŀ��λͼ�ϡ�
 Const PATINVERT = &H5A0049                '/* ͨ��ʹ�ò���OR���򣩲�������Դ��������ȡ�������ɫֵ���ض�ģʽ����ɫ�ϲ���Ȼ��ʹ��OR���򣩲��������ò����Ľ����Ŀ����������ڵ���ɫ�ϲ���
 Const PATPAINT = &HFB0A09                 '/* ͨ��ʹ��XOR����򣩲�������Դ��Ŀ����������ڵ���ɫ�ϲ���
 Const SRCAND = &H8800C6                   '/* ͨ��ʹ��AND���룩����������Դ��Ŀ����������ڵ���ɫ�ϲ�
 Const SRCCOPY = &HCC0020                  '/* ��Դ��������ֱ�ӿ�����Ŀ���������
 Const SRCERASE = &H440328                 '/* ͨ��ʹ��AND���룩��������Ŀ�����������ɫȡ������Դ�����������ɫֵ�ϲ���
 Const SRCINVERT = &H660046                '/* ͨ��ʹ�ò����͵�XOR����򣩲�������Դ��Ŀ������������ɫ�ϲ���
 Const SRCPAINT = &HEE0086                 '/* ͨ��ʹ�ò����͵�OR���򣩲�������Դ��Ŀ������������ɫ�ϲ���
 Const WHITENESS = &HFF0062                '/* ʹ���������ɫ��������1�йص���ɫ���Ŀ��������򡣣�����ȱʡ�����ɫ����˵�������ɫ���ǰ�ɫ����

'--- for mouse_event
 Const MOUSE_MOVED = &H1
 Const MOUSEEVENTF_ABSOLUTE = &H8000       '/*
 Const MOUSEEVENTF_LEFTDOWN = &H2          '/* ģ������������
Const MOUSEEVENTF_LEFTUP = &H4            '/* ģ��������̧��
 Const MOUSEEVENTF_MIDDLEDOWN = &H20       '/* ģ������м�����
 Const MOUSEEVENTF_MIDDLEUP = &H40         '/* ģ������м�����
 Const MOUSEEVENTF_MOVE = &H1              '/* �ƶ���� */
 Const MOUSEEVENTF_RIGHTDOWN = &H8         '/* ģ������Ҽ�����
 Const MOUSEEVENTF_RIGHTUP = &H10          '/* ģ������Ҽ�����
Const MOUSETRAILS = 39                    '/*

 Const BMP_MAGIC_COOKIE = 19778            '/* this is equivalent to ascii string "BM" */
' constants for the biCompression field
 Const BI_RGB = 0&
 Const BI_RLE4 = 2&
 Const BI_RLE8 = 1&
 Const BI_BITFIELDS = 3&
' Const BITSPIXEL = 12                     '/* Number of bits per pixel
' DIB color table identifiers
 Const DIB_PAL_COLORS = 1                  '/* ����ɫ����װ��һ��16λ�������飬�����뵱ǰѡ���ĵ�ɫ���й� color table in palette indices
 Const DIB_PAL_INDICES = 2                 '/* No color table indices into surf palette
 Const DIB_PAL_LOGINDICES = 4              '/* No color table indices into DC palette
 Const DIB_PAL_PHYSINDICES = 2             '/* No color table indices into surf palette
 Const DIB_RGB_COLORS = 0                  '/* ����ɫ����װ��RGB��ɫ

' BLENDFUNCTION AlphaFormat-Konstante
 Const AC_SRC_ALPHA = &H1
' BLENDFUNCTION BlendOp-Konstante
 Const AC_SRC_OVER = &H0

' ======================================================================================
' Methods
' ======================================================================================
' ����SetBkModen����BkMode
 Enum KhanBackStyles
    TRANSPARENT = 1                              '/* ͸������������������� */
    OPAQUE = 2                                   '/* �õ�ǰ�ı���ɫ������߻��ʡ���Ӱˢ���Լ��ַ��Ŀ�϶ */
    NEWTRANSPARENT = 3                           '/* NT4: Uses chroma-keying upon BitBlt. Undocumented feature that is not working on Windows 2000/XP.
End Enum

' ����ε����ģʽ
 Enum KhanPolyFillModeFalgs
    ALTERNATE = 1                                '/* �������
    WINDING = 2                                  '/* ���ݻ�ͼ�������
End Enum

' DrawIconEx
 Enum KhanDrawIconExFlags
    DI_MASK = &H1                                '/* ��ͼʱʹ��ͼ���MASK���֣��絥��ʹ�ã��ɻ��ͼ�����ģ��
    DI_IMAGE = &H2                               '/* ��ͼʱʹ��ͼ���XOR���֣���ͼ��û��͸������
    DI_NORMAL = &H3                              '/* �ó��淽ʽ��ͼ���ϲ� DI_IMAGE �� DI_MASK��
    DI_COMPAT = &H4                              '/* ����׼��ϵͳָ�룬������ָ����ͼ��
    DI_DEFAULTSIZE = &H8                         '/* ����cxWidth��cyWidth���ã�������ԭʼ��ͼ���С
End Enum

'ָ����װ��ͼ������,LoadImage,CopyImage
 Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

 Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/* ��ʾ����һ��ͼ��ľ�ȷ�����������Բ���cxDesired��cyDesired
    LR_COPYDELETEORG = &H8                       '/* ��ʾ����һ��������ɾ��ԭʼͼ��
    LR_CREATEDIBSECTION = &H2000                 '/* ������uTypeָ��ΪIMAGE_BITMAPʱ��ʹ�ú�������һ��DIB����λͼ��������һ�����ݵ�λͼ�������־��װ��һ��λͼ��������ӳ��������ɫ����ʾ�豸ʱ�ǳ����á�
    LR_DEFAULTCOLOR = &H0                        '/* �Գ��淽ʽ����ͼ��
    LR_DEFAULTSIZE = &H40                        '/* �� cxDesired��cyDesiredδ����Ϊ�㣬ʹ��ϵͳָ���Ĺ���ֵ��ʶ����ͼ��Ŀ�͸ߡ���������������������cxDesired��cyDesired����Ϊ�㣬����ʹ��ʵ����Դ�ߴ硣�����Դ�������ͼ����ʹ�õ�һ��ͼ��Ĵ�С��
    LR_LOADFROMFILE = &H10                       '/* ���ݲ���lpszName��ֵװ��ͼ�������δ��������lpszName��ֵΪ��Դ���ơ�
    LR_LOADMAP3DCOLORS = &H1000                  '/* ��ͼ���е����(Dk Gray RGB��128��128��128��)����(Gray RGB��192��192��192��)���Լ�ǳ��(Gray RGB��223��223��223��)���ض��滻��COLOR_3DSHADOW��COLOR_3DFACE�Լ�COLOR_3DLIGHT�ĵ�ǰ����
    LR_LOADTRANSPARENT = &H20                    '/* ��fuLoad����LR_LOADTRANSPARENT��LR_LOADMAP3DCOLORS����ֵ����LRLOADTRANSPARENT���ȡ����ǣ���ɫ��ӿ���COLOR_3DFACE�����������COLOR_WINDOW��
    LR_MONOCHROME = &H1                          '/* ��ͼ��ת���ɵ�ɫ
    LR_SHARED = &H8000                           '/* ��ͼ�񽫱����װ���������LR_SHAREDδ�����ã�������ͬһ����Դ�ڶ��ε������ͼ���Ǿͻ���װ���Ա����ͼ���ҷ��ز�ͬ�ľ����
    LR_COPYFROMRESOURCE = &H4000                 '/*
End Enum

 Enum KhanDrawTextStyles
    DT_BOTTOM = &H8&                             '/* ����ͬʱָ��DT_SINGLE��ָʾ�ı������ʽ�����εĵױ�
    DT_CALCRECT = &H400&                         '/* ���������������ʽ�����Σ����л�ͼʱ���εĵױ߸�����Ҫ������չ���Ա������������֣����л�ͼʱ����չ���ε��Ҳࡣ��������֡���lpRect����ָ���ľ��λ�������������ֵ
    DT_CENTER = &H1&                             '/* �ı���ֱ����
    DT_EXPANDTABS = &H40&                        '/* ������ֵ�ʱ�򣬶��Ʊ�վ������չ��Ĭ�ϵ��Ʊ�վ�����8���ַ������ǣ�����DT_TABSTOP��־�ı������趨
    DT_EXTERNALLEADING = &H200&                  '/* �����ı��и߶ȵ�ʱ��ʹ�õ�ǰ������ⲿ������ԣ�the external leading attribute��
    DT_INTERNAL = &H1000&                        '/* Uses the system font to calculate text metrics
    DT_LEFT = &H0&                               '/* �ı������
    DT_NOCLIP = &H100&                           '/* �������ʱ�����е�ָ���ľ��Σ�DrawTextEx is somewhat faster when DT_NOCLIP is used.
    DT_NOPREFIX = &H800&                         '/* ͨ����������Ϊ & �ַ���ʾӦΪ��һ���ַ������»��ߡ��ñ�־��ֹ������Ϊ
    DT_RIGHT = &H2&                              '/* �ı��Ҷ���
    DT_SINGLELINE = &H20&                        '/* ֻ������
    DT_TABSTOP = &H80&                           '/* ָ���µ��Ʊ�վ��࣬������������ĸ�8λ
    DT_TOP = &H0&                                '/* ����ͬʱָ��DT_SINGLE��ָʾ�ı������ʽ�����εĵױ�
    DT_VCENTER = &H4&                            '/* ����ͬʱָ��DT_SINGLE��ָʾ�ı������ʽ�����ε��в�
    DT_WORDBREAK = &H10&                         '/* �����Զ����С�����SetTextAlign����������TA_UPDATECP��־���������������Ч
' #if(WINVER >= =&H0400)
    DT_EDITCONTROL = &H2000&                     '/* ��һ�����б༭�ؼ�����ģ�⡣����ʾ���ֿɼ�����
    DT_END_ELLIPSIS = &H8000&                    '/* �����ִ������ھ�����ȫ�����£�����ĩβ��ʾʡ�Ժ�
    DT_PATH_ELLIPSIS = &H4000&                   '/* ���ִ������� \ �ַ�������ʡ�Ժ��滻�ִ����ݣ�ʹ�����ھ�����ȫ�����¡����磬һ���ܳ���·�������ܻ���������ʾ����c:\windows\...\doc\readme.txt
    DT_MODIFYSTRING = &H10000                    '/* ��ָ����DT_ENDELLIPSES �� DT_PATHELLIPSES���ͻ���ִ������޸ģ�ʹ����ʵ����ʾ���ִ����
    DT_RTLREADING = &H20000                      '/* ��ѡ���豸��������������ϣ������������ϵ���ʹ��ҵ����������
    DT_WORD_ELLIPSIS = &H40000                   '/* Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

 Enum KhanDrawFrameControlType
    DFC_CAPTION = 1                              '/* Title bar.
    DFC_MENU = 2                                 '/* Menu bar.
    DFC_SCROLL = 3                               '/* Scroll bar.
    DFC_BUTTON = 4                               '/* Standard button.
    DFC_POPUPMENU = 5                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Popup menu item.
End Enum

 Enum KhanDrawFrameControlStyle
    DFCS_BUTTONCHECK = &H0                       '/* Check box.
    DFCS_BUTTONRADIOIMAGE = &H1                  '/* Image for radio button (nonsquare needs image).
    DFCS_BUTTONRADIOMASK = &H2                   '/* Mask for radio button (nonsquare needs mask).
    DFCS_BUTTONRADIO = &H4                       '/* Radio button.
    DFCS_BUTTON3STATE = &H8                      '/* Three-state button.
    DFCS_BUTTONPUSH = &H10                       '/* Push button.
    DFCS_CAPTIONCLOSE = &H0                      '/* <b>Close</b> button.
    DFCS_CAPTIONMIN = &H1                        '/* <b>Minimize</b> button.
    DFCS_CAPTIONMAX = &H2                        '/* <b>Maximize</b> button.
    DFCS_CAPTIONRESTORE = &H3                    '/* <b>Restore</b> button.
    DFCS_CAPTIONHELP = &H4                       '/* <b>Help</b> button.
    DFCS_MENUARROW = &H0                         '/* Submenu arrow.
    DFCS_MENUCHECK = &H1                         '/* Check mark.
    DFCS_MENUBULLET = &H2                        '/* Bullet.
    DFCS_MENUARROWRIGHT = &H4                    '/* Submenu arrow pointing left. This is used for the right-to-left cascading menus used with right-to-left languages such as Arabic or Hebrew.
    DFCS_SCROLLUP = &H0                          '/* Up arrow of scroll bar.
    DFCS_SCROLLDOWN = &H1                        '/* Down arrow of scroll bar.
    DFCS_SCROLLLEFT = &H2                        '/* Left arrow of scroll bar.
    DFCS_SCROLLRIGHT = &H3                       '/* Right arrow of scroll bar.
    DFCS_SCROLLCOMBOBOX = &H5                    '/* Combo box scroll bar.
    DFCS_SCROLLSIZEGRIP = &H8                    '/* Size grip in bottom-right corner of window.
    DFCS_SCROLLSIZEGRIPRIGHT = &H10              '/* Size grip in bottom-left corner of window. This is used with right-to-left languages such as Arabic or Hebrew.
    DFCS_INACTIVE = &H100                        '/* Button is inactive (grayed).
    DFCS_PUSHED = &H200                          '/* Button is pushed.
    DFCS_CHECKED = &H400                         '/* Button is checked.
    DFCS_TRANSPARENT = &H800                     '/* <b>Windows 98/Me, Windows 2000 or later:</b> The background remains untouched.
    DFCS_HOT = &H1000                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Button is hot-tracked.
    DFCS_ADJUSTRECT = &H2000                     '/* Bounding rectangle is adjusted to exclude the surrounding edge of the push button.
    DFCS_FLAT = &H4000                           '/* Button has a flat border.
    DFCS_MONO = &H8000                           '/* Button has a monochrome border.
End Enum

' ָ��������ʽ������CreatePen�Ĳ���CreatePen��ʹ�õĳ���
 Enum KhanPenStyles
    ' CreatePen��ExtCreatePen
    ' ���ʵ���ʽ
    PS_SOLID = 0                                 '/* ���ʻ�������ʵ�� */
    PS_DASH = 1                                  '/* ���ʻ����������ߣ�nWidth������1�� */
    PS_DOT = 2                                   '/* ���ʻ������ǵ��ߣ�nWidth������1�� */
    PS_DASHDOT = 3                               '/* ���ʻ������ǵ㻮�ߣ�nWidth������1�� */
    PS_DASHDOTDOT = 4                            '/* ���ʻ������ǵ�-��-���ߣ�nWidth������1�� */
    PS_NULL = 5                                  '/* ���ʲ��ܻ�ͼ */
    PS_INSIDEFRAME = 6                           '/* ����������Բ�����Ρ�Բ�Ǿ��Ρ���ͼ�Լ��ҵ����ɵķ�ն�����л�ͼ����ָ����׼ȷRGB��ɫ�����ڣ��ͽ��ж������� */
    ' ExtCreatePen
    ' ���ʵ���ʽ
    PS_USERSTYLE = 7                             '/* <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
    PS_ALTERNATE = 8                             '/* <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
    ' ���ʵıʼ�
    PS_ENDCAP_ROUND = &H0                        '/* End caps are round.
    PS_ENDCAP_SQUARE = &H100                     '/* End caps are square.
    PS_ENDCAP_FLAT = &H200                       '/* End caps are flat.
    PS_ENDCAP_MASK = &HF00                       '/* Mask for previous PS_ENDCAP_XXX values.
    ' ��ͼ���������߶λ���·��������ֱ�ߵķ�ʽ
    PS_JOIN_ROUND = &H0                          '/* Joins are beveled.
    PS_JOIN_BEVEL = &H1000                       '/* Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
    PS_JOIN_MITER = &H2000                       '/* Joins are round.
    PS_JOIN_MASK = &HF000                        '/* Mask for previous PS_JOIN_XXX values.
    ' ���ʵ�����
    PS_COSMETIC = &H0                            '/* The pen is cosmetic.
    PS_GEOMETRIC = &H10000                       '/* The pen is geometric.
    '
    PS_STYLE_MASK = &HF                          '/* Mask for previous PS_XXX values.
    PS_TYPE_MASK = &HF0000                       '/* Mask for previous PS_XXX (pen type).
End Enum

 Enum KhanBrushStyle
    BS_SOLID = 0                                 '/* Solid brush.
    BS_HOLLOW = 1                                '/* Hollow brush.
    BS_NULL = 1                                  '/* Same as BS_HOLLOW.
    BS_HATCHED = 2                               '/* Hatched brush.
    BS_PATTERN = 3                               '/* Pattern brush defined by a memory bitmap.
    BS_INDEXED = 4                               '/*
    BS_DIBPATTERN = 5                            '/* A pattern brush defined by a device-independent bitmap (DIB) specification.
    BS_DIBPATTERNPT = 6                          '/* A pattern brush defined by a device-independent bitmap (DIB) specification. If <b>lbStyle</b> is BS_DIBPATTERNPT, the <b>lbHatch</b> member contains a pointer to a packed DIB.
    BS_PATTERN8X8 = 7                            '/* Same as BS_PATTERN.
    BS_DIBPATTERN8X8 = 8                         '/* Same as BS_DIBPATTERN.
    BS_MONOPATTERN = 9                           '/* The brush is a monochrome (black & white) bitmap.
End Enum

 Enum KhanHatchStyles
    HS_HORIZONTAL = 0                            '/* Horizontal hatch.
    HS_VERTICAL = 1                              '/* Vertical hatch.
    HS_FDIAGONAL = 2                             '/* A 45-degree downward, left-to-right hatch.
    HS_BDIAGONAL = 3                             '/* A 45-degree upward, left-to-right hatch.
    HS_CROSS = 4                                 '/* Horizontal and vertical cross-hatch.
    HS_DIAGCROSS = 5                             '/* A 45-degree crosshatch.
End Enum

' DrawEdge
 Enum KhanBorderStyles
    BDR_RAISEDOUTER = &H1                        '/* Raised outer edge.
    BDR_SUNKENOUTER = &H2                        '/* Sunken outer edge.
    BDR_RAISEDINNER = &H4                        '/* Raised inner edge.
    BDR_SUNKENINNER = &H8                        '/* Sunken inner edge.
    BDR_OUTER = &H3                              '/* (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = &HC                              '/* (BDR_RAISEDINNER Or BDR_SUNKENINNER)
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum

 Enum KhanBorderFlags
    BF_LEFT = &H1                                '/* Left side of border rectangle.
    BF_TOP = &H2                                 '/* Top of border rectangle.
    BF_RIGHT = &H4                               '/* Right side of border rectangle.
    BF_BOTTOM = &H8                              '/* Bottom of border rectangle.
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10                           '/* Diagonal border.
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_MIDDLE = &H800                            '/* Fill in the middle.
    BF_SOFT = &H1000                             '/* Use for softer buttons.
    BF_ADJUST = &H2000                           '/* Calculate the space left over.
    BF_FLAT = &H4000                             '/* For flat rather than 3-D borders.
    BF_MONO = &H8000&                            '/* For monochrome borders
End Enum

' ����ָ��һ����λ�ú�״̬������SetWindowPos����
 Enum KhanSetWindowPosStyles
    HWND_BOTTOM = 1                              '/* ���������ڴ����б�ײ� */
    HWND_NOTOPMOST = -2                          '/* �����������б�������λ���κ�������ڵĺ��� */
    HWND_TOP = 0                                 '/* ����������Z���еĶ�����Z���д����ڷּ��ṹ�У��������һ����������Ĵ�����ʾ��˳�� */
    HWND_TOPMOST = -1                            '/* �����������б�������λ���κ�������ڵ�ǰ�� */
    SWP_SHOWWINDOW = &H40                        '/* ��ʾ���� */
    SWP_HIDEWINDOW = &H80                        '/* ���ش��� */
    SWP_FRAMECHANGED = &H20                      '/* ǿ��һ��WM_NCCALCSIZE��Ϣ���봰�ڣ���ʹ���ڵĴ�Сû�иı� */
    SWP_NOACTIVATE = &H10                        '/* ������� */
    SWP_NOCOPYBITS = &H100                       '
    SWP_NOMOVE = &H2                             '/* ���ֵ�ǰλ�ã�x��y�趨�������ԣ� */
    SWP_NOOWNERZORDER = &H200                    '/* Don't do owner Z ordering */
    SWP_NOREDRAW = &H8                           '/* ���ڲ��Զ��ػ� */
    SWP_NOREPOSITION = SWP_NOOWNERZORDER         '
    SWP_NOSIZE = &H1                             '/* ���ֵ�ǰ��С��cx��cy�ᱻ���ԣ� */
    SWP_NOZORDER = &H4                           '/* ���ִ������б�ĵ�ǰλ�ã�hWndInsertAfter�������ԣ� */
    SWP_DRAWFRAME = SWP_FRAMECHANGED             '/* Χ�ƴ��ڻ�һ���� */
'    HWND_BROADCAST = &HFFFF&
'    HWND_DESKTOP = 0
End Enum

' ָ���������ڵķ��
 Enum KhanCreateWindowSytles
    ' CreateWindow
    WS_BORDER = &H800000                         '/* ����һ�����߿�Ĵ��ڡ�
    WS_CAPTION = &HC00000                        '/* ����һ���б����Ĵ��ڣ�����WS_BODER��񣩡�
    WS_CHILD = &H40000000                        '/* ����һ���Ӵ��ڡ�����������WS_POPVP�����á�
    WS_CHILDWINDOW = (WS_CHILD)                  '/* ��WS_CHILD��ͬ��
    WS_CLIPCHILDREN = &H2000000                  '/* ���ڸ������ڻ�ͼʱ���ų��Ӵ��������ڴ���������ʱʹ��������
    WS_CLIPSIBLINGS = &H4000000                  '/* �ų��Ӵ���֮����������Ҳ���ǣ���һ���ض��Ĵ��ڽ��յ�WM_PAINT��Ϣʱ��WS_CLIPSIBLINGS ������в�������ų��ڻ�ͼ֮�⣬ֻ�ػ�ָ�����Ӵ��ڡ����δָ��WS_CLIPSIBLINGS��񣬲����Ӵ����ǲ���ģ������ػ��Ӵ��ڵĿͻ���ʱ���ͻ��ػ��ڽ����Ӵ��ڡ�
    WS_DISABLED = &H8000000                      '/* ����һ����ʼ״̬Ϊ��ֹ���Ӵ��ڡ�һ����ֹ״̬�Ĵ��ղ��ܽ��������û���������Ϣ��
    WS_DLGFRAME = &H400000                       '/* ����һ�����Ի���߿���Ĵ��ڡ����ַ��Ĵ��ڲ��ܴ���������
    WS_GROUP = &H20000                           '/* ָ��һ����Ƶĵ�һ�����ơ�����������ɵ�һ�����ƺ������Ŀ�����ɣ��Եڶ������ƿ�ʼÿ�����ƣ�����WS_GROUP���ÿ����ĵ�һ�����ƴ���WS_TABSTOP��񣬴Ӷ�ʹ�û�����������ƶ����û�������ʹ�ù�������ڵĿ��Ƽ�ı���̽��㡣
    WS_HSCROLL = &H100000                        '/* ����һ����ˮƽ�������Ĵ��ڡ�
    WS_MAXIMIZE = &H1000000                      '/* ����һ��������󻯰�ť�Ĵ��ڡ��÷������WS_EX_CONTEXTHELP���ͬʱ���֣�ͬʱ����ָ��WS_SYSMENU���
    WS_MAXIMIZEBOX = &H10000                     '/*
    WS_MINIMIZE = &H20000000                     '/* ����һ����ʼ״̬Ϊ��С��״̬�Ĵ��ڡ�
    WS_ICONIC = WS_MINIMIZE                      '/* ����һ����ʼ״̬Ϊ��С��״̬�Ĵ��ڡ���WS_MINIMIZE�����ͬ��
    WS_MINIMIZEBOX = &H20000                     '/*
    WS_OVERLAPPED = &H0&                         '/* ����һ������Ĵ��ڡ�һ������Ĵ�����һ����������һ���߿���WS_TILED�����ͬ
    WS_POPUP = &H80000000                        '/* ����һ������ʽ���ڡ��÷������WS_CHLD���ͬʱʹ�á�
    WS_SYSMENU = &H80000                         '/* ����һ���ڱ������ϴ��д��ڲ˵��Ĵ��ڣ�����ͬʱ�趨WS_CAPTION���
    WS_TABSTOP = &H10000                         '/* ����һ�����ƣ�����������û�����Tab��ʱ���Ի�ü��̽��㡣����Tab����ʹ���̽���ת�Ƶ���һ����WS_TABSTOP���Ŀ��ơ�
    WS_THICKFRAME = &H40000                      '/* ����һ�����пɵ��߿�Ĵ��ڡ�
    WS_SIZEBOX = WS_THICKFRAME                   '/* ��WS_THICKFRAME�����ͬ
    WS_TILED = WS_OVERLAPPED                     '/* ����һ������Ĵ��ڡ�һ������Ĵ�����һ�������һ���߿���WS_OVERLAPPED�����ͬ��
    WS_VISIBLE = &H10000000                      '/* ����һ����ʼ״̬Ϊ�ɼ��Ĵ��ڡ�
    WS_VSCROLL = &H200000                        '/* ����һ���д�ֱ�������Ĵ��ڡ�
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW         '/* ����һ������WS_OVERLAPPED��WS_CAPTION��WS_SYSMENU MS_THICKFRAME��
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) '/* ����һ������WS_BORDER��WS_POPUP,WS_SYSMENU���Ĵ��ڣ�WS_CAPTION��WS_POPUPWINDOW����ͬʱ�趨����ʹ����ĳ���ɼ���
    ' CreateWindowEx
    WS_EX_ACCEPTFILES = &H10&                    '/* ָ���Ը÷�񴴽��Ĵ��ڽ���һ����ק�ļ���
    WS_EX_APPWINDOW = &H40000                    '/* �����ڿɼ�ʱ����һ�����㴰�ڷ��õ��������ϡ�
    WS_EX_CLIENTEDGE = &H200                     '/* ָ��������һ������Ӱ�ı߽硣
    WS_EX_CONTEXTHELP = &H400                    '/* �ڴ��ڵı���������һ���ʺű�־�����û�������ʺ�ʱ��������Ϊһ���ʺŵ�ָ�롢��������һ���Ӵ��ڣ����Ӵ��ս��յ�WM_HELP��Ϣ���Ӵ���Ӧ�ý������Ϣ���ݸ������ڹ��̣���������ͨ��HELP_WM_HELP�������WinHelp���������HelpӦ�ó�����ʾһ�������Ӵ��ڰ�����Ϣ�ĵ���ʽ���ڡ� WS_EX_CONTEXTHELP������WS_MAXIMIZEBOX��WS_MINIMIZEBOXͬʱʹ�á�
    WS_EX_CONTROLPARENT = &H10000                '/* �����û�ʹ��Tab���ڴ��ڵ��Ӵ��ڼ�������
    WS_EX_DLGMODALFRAME = &H1&                   '/* ����һ����˫�ߵĴ��ڣ��ô��ڿ�����dwStyle��ָ��WS_CAPTION���������һ����������
    WS_EX_LEFT = &H0                             '/* ���ھ�����������ԣ�����ȱʡ���õġ�
    WS_EX_LEFTSCROLLBAR = &H4000                 '/* ��������������Hebrew��Arabic��������֧��reading order alignment�����ԣ����������������ڣ����ڿͻ������󲿷֡������������ԣ��ڸ÷�񱻺��Բ��Ҳ���Ϊ������
    WS_EX_LTRREADING = &H0                       '/* �����ı���LEFT��RIGHT���������ң����Ե�˳����ʾ������ȱʡ���õġ�
    WS_EX_MDICHILD = &H40                        '/* ����һ��MDI�Ӵ��ڡ�
    WS_EX_NOACTIVATE = &H8000000                 '/*
    WS_EX_NOPATARENTNOTIFY = &H4&                '/* ָ���������񴴽��Ĵ����ڱ�����������ʱ���򸸴��ڷ���WM_PARENTNOTFY��Ϣ��
    WS_EX_OVERLAPPEDWINDOW = &H300               '/*
    WS_EX_PALETTEWINDOW = &H188                  '/* WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOW��WS_WX_TOPMOST�������WS_EX_RIGHT:���ھ�����ͨ���Ҷ������ԣ��������ڴ����ࡣֻ���������������Hebrew,Arabic������֧�ֶ�˳����루reading order alignment��������ʱ�÷�����Ч�����򣬺��Ըñ�־���Ҳ���Ϊ������
    WS_EX_RIGHT = &H1000                         '/*
    WS_EX_RIGHTSCROLLBAR = &H0                   '/* ��ֱ�������ڴ��ڵ��ұ߽硣����ȱʡ���õġ�
    WS_EX_RTLREADING = &H2000                    '/* ��������������Hebrew��Arabic��������֧�ֶ�˳����루reading order alignment�������ԣ��򴰿��ı���һ�������ң�RIGHT��LEFT˳��Ķ���˳�������������ԣ��ڸ÷�񱻺��Բ��Ҳ���Ϊ������
    WS_EX_STATICEDGE = &H20000                   '/* Ϊ�������û���������һ��3һά�߽���
    WS_EX_TOOLWINDOW = &H80                      '/*
    WS_EX_TOPMOST = &H8&                         '/* ָ���Ը÷�񴴽��Ĵ���Ӧ���������з���߲㴰�ڵ����沢��ͣ������L����ʹ����δ�����ʹ�ú���SetWindowPos�����ú���ȥ������
    WS_EX_TRANSPARENT = &H20&                    '/* ָ���������񴴽��Ĵ����ڴ����µ�ͬ���������ػ�ʱ���ô��ڲſ����ػ���
    WS_EX_WINDOWEDGE = &H100
End Enum

' Windows�����йص���Ϣ������GetSystemMetrics����
 Enum KhanSystemMetricsFlags
    SM_CXSCREEN = 0                              '/* ��Ļ��С */
    SM_CYSCREEN = 1                              '/* ��Ļ��С */
    SM_CXVSCROLL = 2                             '/* ��ֱ�������еļ�ͷ��ť�Ĵ�С */
    SM_CYHSCROLL = 3                             '/* ˮƽ�������ϵļ�ͷ��С */
    SM_CYCAPTION = 4                             '/* ���ڱ���ĸ߶� */
    SM_CXBORDER = 5                              '/* �ߴ粻�ɱ�߿�Ĵ�С */
    SM_CYBORDER = 6                              '/* �ߴ粻�ɱ�߿�Ĵ�С */
    SM_CXDLGFRAME = 7                            '/* �Ի���߿�Ĵ�С */
    SM_CYDLGFRAME = 8                            '/* �Ի���߿�Ĵ�С */
    SM_CYVTHUMB = 9                              '/* ��������ˮƽ�������ϵĴ�С */
    SM_CXHTHUMB = 10                             '/* ��������ˮƽ�������ϵĴ�С */
    SM_CXICON = 11                               '/* ��׼ͼ��Ĵ�С */
    SM_CYICON = 12                               '/* ��׼ͼ��Ĵ�С */
    SM_CXCURSOR = 13                             '/* ��׼ָ���С */
    SM_CYCURSOR = 14                             '/* ��׼ָ���С */
    SM_CYMENU = 15                               '/* �˵��߶� */
    SM_CXFULLSCREEN = 16                         '/* ��󻯴��ڿͻ����Ĵ�С */
    SM_CYFULLSCREEN = 17                         '/* ��󻯴��ڿͻ����Ĵ�С */
    SM_CYKANJIWINDOW = 18                        '/* Kanji���ڵĴ�С��Height of Kanji window�� */
    SM_MOUSEPRESENT = 19                         '/* �簲װ�������ΪTRUE */
    SM_CYVSCROLL = 20                            '/* ��ֱ�������еļ�ͷ��ť�Ĵ�С */
    SM_CXHSCROLL = 21                            '/* ˮƽ�������ϵļ�ͷ��С */
    SM_DEBUG = 22                                '/* ��windows�ĵ��԰��������У���ΪTRUE */
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28                                '/* ���ڵ���С�ߴ� */
    SM_CYMIN = 29                                '/* ���ڵ���С�ߴ� */
    SM_CXSIZE = 30                               '/* ������λͼ�Ĵ�С */
    SM_CYSIZE = 31                               '/* ������λͼ�Ĵ�С */
    SM_CXFRAME = 32                              '/* �ߴ�ɱ�߿�Ĵ�С����win95��nt 4.0��ʹ��SM_C?FIXEDFRAME�� */
    SM_CYFRAME = 33                              '/* �ߴ�ɱ�߿�Ĵ�С */
    SM_CXMINTRACK = 34                           '/* ���ڵ���С�켣��� */
    SM_CYMINTRACK = 35                           '/* ���ڵ���С�켣��� */
    SM_CXDOUBLECLK = 36                          '/* ˫������Ĵ�С��ָ����Ļ��һ���ض�����ʾ����ֻ���������������������������굥�������п��ܱ�����˫���¼����� */
    SM_CYDOUBLECLK = 37                          '/* ˫������Ĵ�С */
    SM_CXICONSPACING = 38                        '/* ����ͼ��֮��ļ�����롣��win95��nt 4.0����ָ��ͼ��ļ�� */
    SM_CYICONSPACING = 39                        '/* ����ͼ��֮��ļ�����롣��win95��nt 4.0����ָ��ͼ��ļ�� */
    SM_MENUDROPALIGNMENT = 40                    '/* �絯��ʽ�˵�����˵�����Ŀ����࣬��Ϊ�� */
    SM_PENWINDOWS = 41                           '/* ��װ����֧�ֱʴ��ڵ�DLL�����ʾ�ʴ��ڵľ�� */
    SM_DBCSENABLED = 42                          '/* ��֧��˫�ֽ���ΪTRUE */
    SM_CMOUSEBUTTONS = 43                        '/* ��갴ť������������������û����꣬��Ϊ�� */
    SM_CMETRICS = 44                             '/* ����ϵͳ���������� */
End Enum

' SetMapMode
 Enum KhanMapModeStyles
    MM_ANISOTROPIC = 8                           '/* �߼���λת���ɾ����������������ⵥλ����SetWindowExtEx��SetViewportExtEx������ָ����λ������ͱ�����
    MM_HIENGLISH = 5                             '/* ÿ���߼���λת��Ϊ0.001inch(Ӣ��)��X�����������ң�Y������������
    MM_HIMETRIC = 3                              '/* ÿ���߼���λת��Ϊ0.01millimeter(����)��X���������ң�Y�����������ϡ�
    MM_ISOTROPIC = 7                             '/* �ӿںʹ��ڷ�Χ���⣬ֻ��x��y�߼���Ԫ�ߴ�Ҫ��ͬ
    MM_LOENGLISH = 4                             '/* ÿ���߼���λת��ΪӢ�磬X���������ң�Y���������ϡ�
    MM_LOMETRIC = 2                              '/* ÿ���߼���λת��Ϊ���ף�X���������ң�Y���������ϡ�
    MM_TEXT = 1                                  '/* ÿ���߼���λת��Ϊһ�����ñ��أ�X���������ң�Y���������¡�
    MM_TWIPS = 6                                 '/* ÿ���߼���λת��Ϊ1 twip (1/1440 inch)��X���������ң�Y�������ϡ�
End Enum

' GetROP2,SetROP2
 Enum EnumDrawModeFlags
    R2_BLACK = 1                                 '/* ��ɫ
    R2_COPYPEN = 13                              '/* ������ɫ
    R2_LAST = 16
    R2_MASKNOTPEN = 3                            '/* ������ɫ�ķ�ɫ����ʾ��ɫ����AND����
    R2_MASKPEN = 9                               '/* ��ʾ��ɫ�뻭����ɫ����AND����
    R2_MASKPENNOT = 5                            '/* ��ʾ��ɫ�ķ�ɫ�뻭����ɫ����AND����
    R2_MERGENOTPEN = 12                          '/* ������ɫ�ķ�ɫ����ʾ��ɫ����OR����
    R2_MERGEPEN = 15                             '/* ������ɫ����ʾ��ɫ����OR����
    R2_MERGEPENNOT = 14                          '/* ��ʾ��ɫ�ķ�ɫ�뻭����ɫ����OR����
    R2_NOP = 11                                  '/* ����
    R2_NOT = 6                                   '/* ��ǰ��ʾ��ɫ�ķ�ɫ
    R2_NOTCOPYPEN = 4                            '/* R2_COPYPEN�ķ�ɫ
    R2_NOTMASKPEN = 8                            '/* R2_MASKPEN�ķ�ɫ
    R2_NOTMERGEPEN = 2                           '/* R2_MERGEPEN�ķ�ɫ
    R2_NOTXORPEN = 10                            '/* R2_XORPEN�ķ�ɫ
    R2_WHITE = 16                                '/* ��ɫ
    R2_XORPEN = 7                                '/* ��ʾ��ɫ�뻭����ɫ�����������
End Enum

' ======================================================================================
' Types
' ======================================================================================

Private Type tagINITCOMMONCONTROLSEX              '/* icc
   dwSize                   As Long              '/* size of this structure
   dwICC                    As Long              '/* flags indicating which classes to be initialized.
End Type

Private Type POINTAPI
   X                        As Long
   Y                        As Long
End Type

 Private Type RECT
   Left                     As Long
   Top                      As Long
   Right                    As Long
   Bottom                   As Long
End Type

Private Type LOGPEN
    lopnStyle               As Long
    lopnWidth               As POINTAPI
    lopnColor               As Long
End Type

Private Type LOGBRUSH
   lbStyle                  As Long
   lbColor                  As Long
   lbHatch                  As Long
End Type

' ����ṹ�����˸��ӵĻ�ͼ����������DrawTextEx
Private Type DRAWTEXTPARAMS
    cbSize                  As Long              '/* Specifies the structure size, in bytes */
    iTabLength              As Long              '/* Specifies the size of each tab stop, in units equal to the average character width */
    iLeftMargin             As Long              '/* Specifies the left margin, in units equal to the average character width */
    iRightMargin            As Long              '/* Specifies the right margin, in units equal to the average character width */
    uiLengthDrawn           As Long              '/* Receives the number of characters processed by DrawTextEx, including white-space characters. */
                                                 '/* The number can be the length of the string or the index of the first line that falls below the drawing area. */
                                                 '/* Note that DrawTextEx always processes the entire string if the DT_NOCLIP formatting flag is specified */
End Type

Private Const LF_FACESIZE   As Long = 32
 Private Type LOGFONT
   lfHeight                 As Long              '/* The font size (see below) */
   lfWidth                  As Long              '/* Normally you don't set this, just let Windows create the Default */
   lfEscapement             As Long              '/* The angle, in 0.1 degrees, of the font */
   lfOrientation            As Long              '/* Leave as default */
   lfWeight                 As Long              '/* Bold, Extra Bold, Normal etc */
   lfItalic                 As Byte              '/* As it says */
   lfUnderline              As Byte              '/* As it says */
   lfStrikeOut              As Byte              '/* As it says */
   lfCharSet                As Byte              '/* As it says */
   lfOutPrecision           As Byte              '/* Leave for default */
   lfClipPrecision          As Byte              '/* Leave for defaultv
   lfQuality                As Byte              '/* Leave for default */
   lfPitchAndFamily         As Byte              '/* Leave for default */
   lfFaceName(LF_FACESIZE)  As Byte              '/* The font name converted to a byte array */
End Type

Private Type ICONINFO
   fIcon                    As Long
   xHotspot                 As Long
   yHotspot                 As Long
   hbmMask                  As Long
   hbmColor                 As Long
End Type

Private Type IMAGEINFO
    hBitmapImage            As Long
    hBitmapMask             As Long
    cPlanes                 As Long
    cBitsPerPixel           As Long
    rcImage                 As RECT
End Type

'/* DIB ���ļ���С���ܹ�ѶϢ */
Private Type BITMAPFILEHEADER
    bfType                  As Integer           '/* ָ���ļ����ͣ����� BM("magic cookie" - must be "BM" (19778)) */
    bfSize                  As Long              '/* ָ��λͼ�ļ���С����λԪ��Ϊ��λ */
    bfReserved1             As Integer           '/* ������������Ϊ0 */
    bfReserved2             As Integer           '/* ͬ�� */
    bfOffBits               As Long              '/* �Ӵ˼ܹ���λͼ����λ��λԪ��ƫ���� */
End Type

'/* �豸�޹�λͼ (DIB)�Ĵ�С����ɫ��Ϣ  (��λ�� bmp �ļ��Ŀ�ͷ��) 40 bytes */
 Private Type BITMAPINFOHEADER
    biSize                  As Long              '/* �ṹ���� */
    biWidth                 As Long              '/* ָ��λͼ�Ŀ�ȣ�������Ϊ��λ */
    biHeight                As Long              '/* ָ��λͼ�ĸ߶ȣ�������Ϊ��λ */
    biPlanes                As Integer           '/* ָ��Ŀ���豸�ļ���(����Ϊ 1 ) */
    biBitCount              As Integer           '/* λͼ����ɫλ��,ÿһ�����ص�λ(1��4��8��16��24��32) */
    biCompression           As Long              '/* ָ��ѹ������(BI_RGB Ϊ��ѹ��) */
    biSizeImage             As Long              '/* ͼ��Ĵ�С,���ֽ�Ϊ��λ,����BI_RGB��ʽ��,������Ϊ0 */
    biXPelsPerMeter         As Long              '/* ָ���豸ˮ׼�ֱ��ʣ���ÿ�׵�����Ϊ��λ */
    biYPelsPerMeter         As Long              '/* ��ֱ�ֱ��ʣ�����ͬ�� */
    biClrUsed               As Long              '/* ˵��λͼʵ��ʹ�õĲ�ɫ���е���ɫ������,��Ϊ0�Ļ�,˵��ʹ�����е�ɫ���� */
    biClrImportant          As Long              '/* ˵����ͼ����ʾ����ҪӰ�����ɫ��������Ŀ�������0����ʾ����Ҫ */
End Type

'/* �������ɺ졢�̡�����ɵ���ɫ��� */
 Private Type RGBQUAD
    rgbBlue                 As Byte
    rgbGreen                As Byte
    rgbRed                  As Byte
    rgbReserved             As Byte              '/* '����������Ϊ 0 */
End Type

Private Type BITMAPINFO
    bmiHeader               As BITMAPINFOHEADER
    bmiColors               As RGBQUAD
End Type

 Private Type BITMAPINFO_1BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 1)        As RGBQUAD
End Type

 Private Type BITMAPINFO_4BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 15)       As RGBQUAD
End Type

 Private Type BITMAPINFO_8BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 255)      As RGBQUAD
End Type

 Private Type BITMAPINFO_ABOVE8
   bmiHeader                As BITMAPINFOHEADER
End Type

 Private Type BITMAP
    bmType                  As Long              '/* Type of bitmap */
    bmWidth                 As Long              '/* Pixel width */
    bmHeight                As Long              '/* Pixel height */
    bmWidthBytes            As Long              '/* Byte width = 3 x Pixel width */
    bmPlanes                As Integer           '/* Color depth of bitmap */
    bmBitsPixel             As Integer           '/* Bits per pixel, must be 16 or 24 */
    bmBits                  As Long              '/* This is the pointer to the bitmap data */
End Type

' AlphaBlend
 Private Type BLENDFUNCTION
   BlendOp                  As Byte
   BlendFlags               As Byte
   SourceConstantAlpha      As Byte
   AlphaFormat              As Byte
End Type

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1

Private Const MAX_PATH = 260

' ======================================================================================
' Types:
' ======================================================================================

Private Type PICTDESC
    cbSizeofStruct  As Long
    picType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

Private Type Guid
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

' ======================================================================================
' API declares:
' ======================================================================================

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
' ����ָ���豸�����Ļ�ͼģʽ����vb��DrawMode������ȫһ��
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

' ======================================================================================
' API declares:
' ======================================================================================

'����������������������������������������������������������������������������������������
'��-----------------------------��Ϣ��������Ϣ�жӺ���---------------------------------��
'��                                                                                    ��
'
' ����һ�����ڵĴ��ں�������һ����Ϣ�����Ǹ����ڡ�������Ϣ������ϣ�����ú������᷵�ء�
' SendMessageBynum�� SendMessageByString�Ǹú����ġ����Ͱ�ȫ��������ʽ
 Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
 Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ��һ����ϢͶ�ݵ�ָ�����ڵ���Ϣ���С�Ͷ�ݵ���Ϣ����Windows�¼���������еõ�����
' ���Ǹ�ʱ�򣬻���ͬͶ�ݵ���Ϣ����ָ�����ڵĴ��ں������ر��ʺ���Щ����Ҫ��������Ĵ�����Ϣ�ķ���
 Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��--------------------------------���ں���(Window)------------------------------------��
'��                                                                                    ��
'
' Creating new windows:
 Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' ��С��ָ���Ĵ��ڡ����ڲ�����ڴ������
 Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
' �ƻ����������ָ���Ĵ����Լ����������Ӵ���
 Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
' ��ָ���Ĵ�����������ֹ������꼰��������
 Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
' �ڴ����б���Ѱ����ָ����������ĵ�һ���Ӵ���
 Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' �ж�ָ�����ڵĸ�����
 Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
' ָ��һ�����ڵ��¸�����vb��ʹ�ã��������������vb���Զ�����ʽ֧���Ӵ��ڡ�
' ���磬�ɽ��ؼ���һ���������������е���һ��������������ڴ�����ƶ��ؼ����൱ð�յģ�
' ��ȴ��ʧΪһ����Ч�İ취������������������ڹر��κ�һ������֮ǰ��ע����SetParent���ؼ��ĸ����ԭ�����Ǹ���
 Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' ����ָ�����ڣ���ֹ�����¡�ͬʱֻ����һ�����ڴ�������״̬
 Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
' ǿ���������´��ڣ���������ǰ���ε��������򶼻��ػ�
' ��vb��ʹ�ã���vb�����ؼ����κβ�����Ҫ���£��ɿ���ֱ��ʹ��refresh����
 Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
' �ж�һ�����ھ���Ƿ���Ч
 Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
' ���ƴ��ڵĿɼ���
 Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' �ı�ָ�����ڵ�λ�úʹ�С���������ڿ�����������С�ߴ�����ƣ���Щ�ߴ��������������õĲ���
 Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' ���������Ϊ����ָ��һ����λ�ú�״̬����Ҳ�ɸı䴰�����ڲ������б��е�λ�á�
' �ú�����DeferWindowPos�������ƣ�ֻ�������������������ֳ�����
' ��vb��ʹ�ã����vb���壬��������win32�����λ���С���������������״̬��
' ���б�Ҫ������һ�����ദ��ģ�����������״̬)
' ����
' hwnd             ����λ�Ĵ���
' hWndInsertAfter  ���ھ�����ڴ����б��У�����hwnd������������ھ���ĺ��棬�ο���ģ��ö��KhanSetWindowPosStyles
' x                �����µ�x���ꡣ��hwnd��һ���Ӵ��ڣ���x�ø����ڵĿͻ��������ʾ
' y                �����µ�y���ꡣ��hwnd��һ���Ӵ��ڣ���y�ø����ڵĿͻ��������ʾ
' cx               ָ���µĴ��ڿ��
' cy               ָ���µĴ��ڸ߶�
' wFlags           ����������һ���������ο���ģ��ö��KhanSetWindowPosStyles
 Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' ��ָ�����ڵĽṹ��ȡ����Ϣ��nIndex�����ο���ģ�鳣������
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' �ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ��nIndex�����ο���ģ�鳣������
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��------------------------------�����ຯ��(Window Class)------------------------------��
'��                                                                                    ��
'
' Ϊָ���Ĵ���ȡ������
 Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��-----------------------------������뺯��(Mouse Input)------------------------------��
'
' ���һ�����ڵľ�����������λ�ڵ�ǰ�����̣߳���ӵ����겶������������գ�
 Private Declare Function GetCapture Lib "user32" () As Long
' ����겶�����õ�ָ���Ĵ��ڡ�����갴ť���µ�ʱ��������ڻ�Ϊ��ǰӦ�ó��������ϵͳ���������������
 Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
' Ϊ��ǰ��Ӧ�ó����ͷ���겶��
 Private Declare Function ReleaseCapture Lib "user32" () As Long
' ����ģ��һ������¼����������������˫�����Ҽ�������
 Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
' ��������ж�ָ���ĵ��Ƿ�λ�ھ���lpRect�ڲ�
' Private Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��-----------------------------�������뺯��(Mouse Input)------------------------------��
'
' ���ӵ�����뽹��Ĵ��ڵľ��
 Private Declare Function GetFocus Lib "user32" () As Long
' ���뽹���赽ָ���Ĵ���
 Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��----------------����ռ���任����(Coordinate Space Transtormation)-----------------��
'
' �жϴ������Կͻ��������ʾ��һ�������Ļ����
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
' �ж���Ļ��һ��ָ����Ŀͻ�������
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------�豸��������(Device Context)-----------------------------��
'
' ����һ�����ض��豸����һ�µ��ڴ��豸�������ڻ���֮ǰ����ҪΪ���豸����ѡ��һ��λͼ��
' ������Ҫʱ�����豸��������DeleteDC����ɾ����ɾ��ǰ�������ж���Ӧ�ظ���ʼ״̬
 Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
' Ϊר���豸�����豸����
 Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
' ��ȡָ�����ڵ��豸�������ñ�������ȡ���豸����һ��Ҫ��ReleaseDC�����ͷţ�������DeleteDC
 Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
' �ͷ��ɵ���GetDC��GetWindowDC������ȡ��ָ���豸�������������˽���豸������Ч���������ĵ��ò�������𺦣�
 Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
' ɾ��ר���豸��������Ϣ�������ͷ�������ش�����Դ����Ҫ��������GetDC����ȡ�ص��豸����
 Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
' ÿ���豸������������ѡ�����е�ͼ�ζ������а���λͼ��ˢ�ӡ����塢�����Լ�����ȵȡ�
' һ��ѡ���豸������ֻ����һ������ѡ���Ķ�������豸�����Ļ�ͼ������ʹ�á�
' ���磬��ǰѡ���Ļ��ʾ��������豸�����������߶���ɫ����ʽ
' ����ֵͨ�����ڻ��ѡ��DC�Ķ����ԭʼֵ��
' ��ͼ������ɺ�ԭʼ�Ķ���ͨ��ѡ���豸�����������һ���豸����ǰ�����ע��ָ�ԭʼ�Ķ���
 Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
' ���������ɾ��GDI���󣬱��续�ʡ�ˢ�ӡ����塢λͼ�������Լ���ɫ��ȵȡ�����ʹ�õ�����ϵͳ��Դ���ᱻ�ͷ�
' ��Ҫɾ��һ����ѡ���豸�����Ļ��ʡ�ˢ�ӻ�λͼ����ɾ����λͼΪ��������Ӱ��ͼ����ˢ�ӣ�
' λͼ�������������ɾ������ֻ��ˢ�ӱ�ɾ��
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'����ָ���豸����������豸�Ĺ��ܷ�����Ϣ
 Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
' ȡ�ö�ָ���������˵����һ���ṹ
' lpObject �κ����ͣ��������ɶ������ݵĽṹ��
' ��Ի��ʣ�ͨ����һ��LOGPEN�ṹ�������չ���ʣ�ͨ����EXTLOGPEN��
' ���������LOGBRUSH�����λͼ��BITMAP�����DIBSectionλͼ��DIBSECTION��
' ��Ե�ɫ�壬Ӧָ��һ�����ͱ����������ɫ���е���Ŀ����
 Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' �ڴ��ڣ����豸����������ˮƽ�ͣ��򣩴�ֱ��������
Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
' �������������Ϊһ��������
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
' ����һ���ɵ�X1��Y1��X2��Y2�����ľ������򣬲���ʱһ��Ҫ��DeleteObject����ɾ��������
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ����һ����lpRectȷ���ľ�������
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
' ����һ��Բ�Ǿ��Σ��þ�����X1��Y1-X2��Y2ȷ��������X3��Y3ȷ������Բ����Բ�ǻ���
' �øú�����������������RoundRect API��������Բ�Ǿ��β���ȫ��ͬ����Ϊ�����ε��ұߺ��±߲�����������֮��
 Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' ��ָ��ˢ�����ָ������
 Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
' ��ָ��ˢ��Χ��ָ������һ�����
 Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
' ������Щ��������ע�⵽�ĶԱ������˵�Ǹ��޴�ı��ص�������API�����е�һ�����������������ı䴰�ڵ�����
' ͨ�����д��ڶ��Ǿ��εġ�������һ�����ھͺ���һ���������򡣱���������������������
' ����ζ�������Դ���Բ�ġ����εĴ��ڣ�Ҳ���Խ�����Ϊ��������ಿ�֡���ʵ���Ͽ������κ���״
 Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
' �ú���ѡ��һ��������Ϊָ���豸�����ĵ�ǰ��������
 Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------------λͼ����(Bitmap)-----------------------------------��
'
' �ú���������ʾ͸�����͸�����ص�λͼ��
 Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
' ��һ��λͼ��һ���豸�������Ƶ���һ����Դ��Ŀ��DC�໥��������
' ��NT�����£�����һ�����紫����Ҫ����Դ�豸�����н��м��л���ת�������������ִ�л�ʧ��
' ��Ŀ���ԴDC��ӳ���ϵҪ����������صĴ�С�����ڴ�������иı䣬
' ��ô��������������Ҫ�Զ���������ת���۵������жϣ��Ա�������յĴ������
' dwRop��ָ����դ�������롣��Щ���뽫����Դ�����������ɫ���ݣ������Ŀ������������ɫ������������������ɫ��
 Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' ����һ�����豸�й�λͼ
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
' ����һ�����豸�й�λͼ������ָ�����豸��������
' �ڴ��豸���������ɫλͼ���ݣ�Ҳ�뵥ɫλͼ���ݡ���������������Ǵ���һ���뵱ǰѡ��hdc�еĳ������ݡ�
' ��һ���ڴ泡����˵��Ĭ�ϵ�λͼ�ǵ�ɫ�ġ������ڴ��豸������һ��DIBSectionѡ�����У�
' ��������ͻ᷵��DIBSection��һ���������hdc��һ���豸λͼ��
' ��ô������ɵ�λͼ�Ϳ϶��������豸��Ҳ����˵����ɫ�豸���ɵĿ϶��ǲ�ɫλͼ��
' ���nWidth��nHeightΪ�㣬���ص�λͼ����һ��1��1�ĵ�ɫλͼ
' һ��λͼ������Ҫ��һ����DeleteObject�����ͷ���ռ�õ��ڴ漰��Դ
 Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
' �ú��������豸�޹ص�λͼ��DIB���������豸�йص�λͼ��DDB����������ѡ���Ϊλͼ��λ��
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
' �ú�������Ӧ�ó������ֱ��д��ġ����豸�޹ص�λͼ��DIB����
' �ú����ṩһ��ָ�룬��ָ��ָ��λͼλ����ֵ�ĵط���
' ���Ը��ļ�ӳ������ṩ���������ʹ���ļ�ӳ�����������λͼ��������ϵͳΪλͼ�����ڴ档
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
' ����λͼ��ͼ���ָ�룬ͬʱ�ڸ��ƹ����н���һЩת������
 Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
' ����һ��λͼ��ͼ���ָ��
 Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
 Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��----------------------------------ͼ�꺯��(Icon)------------------------------------��
'
' ����ָ��ͼ������ָ���һ��������������������ڷ������õ�Ӧ�ó���
 Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
' ����һ��ͼ��
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
' �ú������ͼ����ͷ��κα�ͼ��ռ�õĴ洢�ռ䡣
 Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
' �ú������޶����豸�����Ĵ��ڵĿͻ��������ͼ��
 Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
' �ú������޶����豸�����Ĵ��ڵĿͻ��������ͼ�ִ꣬���޶��Ĺ�դ�����������ض�Ҫ���쳤��ѹ��ͼ����ꡣ
 Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
' ȡ����ͼ���йص���Ϣ
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------------��꺯��(Cursor)-----------------------------------��
'
 Private Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
' ��ָ����ģ���Ӧ�ó���ʵ��������һ�����ָ�롣LoadCursorBynum��LoadCursor���������Ͱ�ȫ����
 Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
' �ú�������һ����겢�ͷ���ռ�õ��κ��ڴ棬��Ҫʹ�øú���ȥ����һ�������ꡣ
 Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
' ��ȡ���ָ��ĵ�ǰλ��
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' �ú����ѹ���Ƶ���Ļ��ָ��λ�á������λ�ò����� ClipCursor�������õ���Ļ��������֮�ڣ�
' ��ϵͳ�Զ��������꣬ʹ�ù���ھ���֮�ڡ�
 Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��-----------------------------��ˢ����(Pen and Brush)---------------------------------��
'
' ��ָ������ʽ����Ⱥ���ɫ����һ�����ʣ���DeleteObject��������ɾ��
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' ����ָ����LOGPEN�ṹ����һ������
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
' ����һ����չ���ʣ�װ�λ򼸺Σ�
Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
' ��һ��LOGBRUSH���ݽṹ�Ļ����ϴ���һ��ˢ��
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
' �ú������Դ���һ������ָ����Ӱģʽ����ɫ���߼�ˢ�ӡ�
 Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
' �ú������Դ�������ָ��λͼģʽ���߼�ˢ�ӣ���λͼ������DIB���͵�λͼ��DIBλͼ����CreateDIBSection���������ġ�
 Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
' �ô�ɫ����һ��ˢ�ӣ�һ��ˢ�Ӳ�����Ҫ������DeleteObject��������ɾ��
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' Ϊ�κ�һ�ֱ�׼ϵͳ��ɫȡ��һ��ˢ�ӣ���Ҫ��DeleteObject����ɾ����Щˢ�ӡ�
' ��������ϵͳӵ�еĹ��ж��󡣲�Ҫ����Щˢ��ָ����һ�ִ������Ĭ��ˢ��
 Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------��������ĺ���(Font and Text)-----------------------------��
'
' ��ָ�������Դ���һ���߼����壬VB������������ѡ�������ʱ���Եø���Ч
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
' ���ı���浽ָ���ľ����У�wFormat��־�����ο�KhanDrawTextStyles
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
' �ú���ȡ��ָ���豸�����ĵ�ǰ������ɫ��
 Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
' ���õ�ǰ�ı���ɫ��������ɫҲ��Ϊ��ǰ��ɫ������ı���������ã�ע��ָ�VB�����ؼ�ԭʼ���ı���ɫ
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��------------------------------------��ͼ����----------------------------------------��
'
' �ú�����һ��Բ����Բ������һ����Բ��һ���߶Σ���֮Ϊ���ߣ��ཻ�޶��ıպ�����
' �˻��ɵ�ǰ�Ļ��ʻ��������ɵ�ǰ�Ļ�ˢ��䡣
 Private Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' ��ָ������ʽ���һ�����εı߿������������������û�б�Ҫ��ʹ�����3D�߿����塣
' ���Ծ���Դ���ڴ��ռ������˵�����������Ч��Ҫ�ߵöࡣ������һ���̶�����������
' hdc      Ҫ�����л�ͼ���豸����
' qrc      ҪΪ�����߿�ľ���
' edge     ����ǰ׺BDR_��������������ϡ�һ��ָ���ڲ��߿�����͹�����°�����һ����ָ���ⲿ�߿���ʱ�ܻ��ô�EDGE_ǰ׺�ĳ�����
' grfFlags ����BF_ǰ׺�ĳ��������
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
' ��һ��������Ρ�����������ڱ�־�������ʽ��ͨ�����������ɵģ�����ͨ����һ�����߱�ʾ��
' ����ͬ���Ĳ����ٴε�������������ͱ�ʾɾ���������
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
' ��������������һ����׼�ؼ�
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
' ���������Ϊһ��ͼ����ͼ����Ӧ�ø�ʽ������Ч��
 Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
' �ú������ڻ�һ����Բ����Բ���������޶����ε����ģ�ʹ�õ�ǰ���ʻ���Բ���õ�ǰ�Ļ�ˢ�����Բ��
 Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ��ָ����ˢ�����һ�����Σ����ε��ұߺ͵ױ߲������
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' ��ָ����ˢ��Χ��һ�����λ�һ���߿����һ��֡�����߿�Ŀ����һ���߼���λ
 Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' ȡ��ָ���豸������ǰ�ı�����ɫ
 Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
' ���ָ�����豸������ȡ�õ�ǰ�ı������ģʽ
 Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
' Ϊָ�����豸�������ñ�����ɫ��������ɫ���������Ӱˢ�ӡ����߻����Լ��ַ����米��ģʽΪOPAQUE���еĿ�϶��
' Ҳ��λͼ��ɫת���ڼ�ʹ�á�����ʵ�����豸�ܹ���ʾ����ӽ��� crColor ����ɫ
 Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
' ָ����Ӱˢ�ӡ����߻����Լ��ַ��еĿ�϶����䷽ʽ������ģʽ����Ӱ������չ������������
 Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
' ��ָ�����豸������ȡ��һ�����ص�RGBֵ
 Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
' ��ָ�����豸����������һ�����ص�RGBֵ
 Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
' ������һ��λͼ�Ķ�����λ���Ƶ�һ�����豸�޹ص�λͼ��
' Private Declare Function GetDIBits Lib "gdi32" ( ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
 Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' ���������豸�޹�λͼ�Ķ�����λ���Ƶ�һ�����豸�йص�λͼ��
 Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' ���ָ�����豸��������ö�������ģʽ��
 Private Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
' ���ö���ε����ģʽ
 Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
' ���ָ�����豸������ȡ�õ�ǰ�Ļ�ͼģʽ�������ɶ����ͼ���������������ʾ��ͼ��ϲ�����
' �������ֻ�Թ�դ�豸��Ч
 Private Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
' ����ָ���豸�����Ļ�ͼģʽ��

' �õ�ǰ���ʻ�һ���ߣ��ӵ�ǰλ������һ��ָ���ĵ㡣�������������ϣ���ǰλ�ñ��x,y��
 Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
' Ϊָ�����豸����ָ��һ���µĵ�ǰ����λ�á�ǰһ��λ�ñ�����lpPoint��
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
' �ú�����һ������Բ�������뾶�ཻ�պ϶��ɵı�״Ш��ͼ���˱�ͼ�ɵ�ǰ���ʻ��������ɵ�ǰ��ˢ��䡣
 Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' �ú�����һ����ֱ�����ŵ��������϶�����ɵĶ���Σ��õ�ǰ���ʻ������������
' �õ�ǰ��ˢ�Ͷ�������ģʽ������Ρ�
 Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
' �õ�ǰ�������һϵ���߶Ρ�ʹ��PolylineTo����ʱ����ǰλ�û���Ϊ���һ���߶ε��յ㡣
' ��������Polyline�����Ķ�
 Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
 Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
 Private Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
' �ú�����һ�����Σ��õ�ǰ�Ļ��ʻ������������õ�ǰ��ˢ������䡣
 Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ������һ����Բ�ǵľ��Σ��˾����ɵ�ǰ���ʻ����ȣ��ɵ�ǰ��ˢ��䡣
 Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' �����������������Сһ�����εĴ�С��
' x�����Ҳ����򣬲�����������ȥ����xΪ��������������εĿ�ȣ���xΪ�������ܼ�С����
' y�Զ�����ײ����������Ӱ���������Ƶ�
 Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
' �ú���ͨ��Ӧ��һ��ָ����ƫ�ƣ��Ӷ��þ����ƶ�������
' x����ӵ��Ҳ���������y��ӵ������͵ײ�����
' ƫ�Ʒ�����ȡ���ڲ������������Ǹ������Լ����õ���ʲô����ϵͳ
 Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
' ������windows�����йص���Ϣ��nIndexֵ�ο���ģ��ĳ�������
 Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
' ����������ڵķ�Χ���Σ����ڵı߿򡢱����������������˵��ȶ������������
 Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' ����ָ�����ڿͻ������εĴ�С
 Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' �����������һ�����ڿͻ�����ȫ���򲿷�������ᵼ�´������¼��ڼ䲿���ػ�
 Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
' �ж�ָ��windows��ʾ�������ɫ����ɫ���󿴱�ģ������
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
 Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

'Initializes the entire common control dynamic-link library.
'Exported by all versions of Comctl32.dll.
 Private Declare Sub InitCommonControls Lib "Comctl32" ()
'Initializes specific common controls classes from the common
'control dynamic-link library.
'Returns TRUE (non-zero) if successful, or FALSE otherwise.
'Began being exported with Comctl32.dll version 4.7 (IE3.0 & later).
 Private Declare Function InitCommonControlsEx Lib "Comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
 Private Declare Function ImageList_GetBkColor Lib "Comctl32" (ByVal hImageList As Long) As Long
 Private Declare Function ImageList_ReplaceIcon Lib "Comctl32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
 Private Declare Function ImageList_Convert Lib "Comctl32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
 Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
 Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
 Private Declare Function ImageList_Replace Lib "Comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
 Private Declare Function ImageList_Add Lib "Comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, hbmMask As Long) As Long
 Private Declare Function ImageList_Remove Lib "Comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long) As Long
 Private Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, pImageInfo As IMAGEINFO) As Long
 Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
 Private Declare Function ImageList_GetIcon Lib "Comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
 Private Declare Function ImageList_SetImageCount Lib "Comctl32" (ByVal hImageList As Long, uNewCount As Long)
 Private Declare Function ImageList_GetImageCount Lib "Comctl32" (ByVal hImageList As Long) As Long
 Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long
 Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
 Private Declare Function ImageList_SetIconSize Lib "Comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
 Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
' Draw an item in an ImageList with more control over positioning and colour:
 Private Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
 Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, prcImage As RECT) As Long
 Private Declare Function ImageList_LoadImage Lib "Comctl32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)
 Private Declare Function ImageList_SetBkColor Lib "Comctl32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
 Private Declare Function ImageList_Copy Lib "Comctl32" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long

' ======================================================================================
' Enums
' ======================================================================================

' �߿���ʽ
Public Enum GPTAB_BORDERSTYLE_METHOD
   GpTabBorderStyleNone = 0               ' û�б߿�
   GpTabBorderStyle3D = 1                 ' 3D
   GpTabBorderStyle3DThin = 2             ' 3DThin
End Enum

' ��ʽ
Public Enum GPTAB_STYLE_METHOD
    GpTabStyleStandard = 0                '/* Win32 ���
    GpTabStyleWinXP = 1                   '/* XP ���
End Enum

' ѡ�����
Public Enum GPTAB_PLACEMENT_METHOD
    GpTabPlacementTopleft = 0
    GpTabPlacementTopRight = 1
    GpTabPlacementBottomLeft = 2
    GpTabPlacementBottomRight = 3
    GpTabPlacementLeftTop = 4
    GpTabPlacementLeftBottom = 5
    GpTabPlacementRightTop = 6
    GpTabPlacementRightBottom = 7
End Enum

' ѡ���ʽ
Public Enum GPTAB_TABSTYLE_METHOD
    GpTabRectangle = 0                 '
    GpTabRoundRect = 1                 '
    GpTabTrapezoid = 2                 '
End Enum

' ѡ����
Public Enum GPTAB_TABWIDTHSTYLE_METHOD
    GpTabJustified = 0                 '
    GpTabnonJustified = 1              '
    GpTabFixed = 2                     '
End Enum

Public Enum GPTAB_XPCOLORSCHEME_METHOD
    GpTabUseWindows = 0
    GpTabCustom = 1
End Enum

' ======================================================================================
' Types
' ======================================================================================

' ���ڱ�־��չ
Private Type TabState
    Index As Long       ' ���ڴ��Tab��Index
End Type

Private Type TabListType
    list() As TabState  ' ȡ��ÿ����Tab��Index
    Count As Long       ' һ����Tab�ĸ���
End Type

' ======================================================================================
' Private variables:
' ======================================================================================

' Icons:
Private m_hIml                    As Long
Private m_lIconSizeX              As Long
Private m_lIconSizeY              As Long
Private m_lngFontHeight           As Long
Private m_lngDefaultTabHeight     As Long

Private m_lngXPFaceColor         As Long
Private m_oleBackColor As OLE_COLOR      ' �ؼ��ı�����ɫ
Private m_oleTabColor As OLE_COLOR       ' ѡ�����ѡʱ��ɫ
Private m_oleTabColorActive As OLE_COLOR ' ѡ�����ʱ��ɫ
Private m_oleTabColorHover As OLE_COLOR  ' ѡ��ȸ���ʱ��ɫ
Private m_oleTabBorderColor As OLE_COLOR ' XP���,GpTabBorderStyleNone�ؼ��߿����ɫ
Private m_blnAutoBackColor As Boolean ' �жϿؼ��ı�����ɫ�Ƿ��游����ı�����ɫ�ı���ı�
Private m_blnUserMode As Boolean ' �ؼ���������ƽ׶�?���н׶�?
Private m_blnEnabled As Boolean ' Enable
Private m_blnHotTracking As Boolean ' �ȸ���
Private m_blnMultiRow As Boolean

Private m_lngTabFixedHeight As Long     ' ����Tab�ĸ߶�
Private m_lngTabFixedWidth As Long      ' ����Tab�Ŀ��

Private m_udtMainRect         As RECT     ' ������
Private m_lngCurrentList        As Long  ' ��ǰTab�б������
Private m_lngListCount         As Long  ' Tab�м���
Private m_aryTabList()       As TabListType

Private m_udtBorderStyle As GPTAB_BORDERSTYLE_METHOD
Private m_udtXPColorScheme As GPTAB_XPCOLORSCHEME_METHOD
Private m_udtPlacement As GPTAB_PLACEMENT_METHOD
Private m_udtStyle As GPTAB_STYLE_METHOD
Private m_udtTabStyle As GPTAB_TABSTYLE_METHOD
Private m_udtTabWidthStyle As GPTAB_TABWIDTHSTYLE_METHOD
Private m_udtDrawTextParams           As DRAWTEXTPARAMS
Private m_clsSelectTab As cTabItem
Private m_clsHoverTab As cTabItem
Private WithEvents m_clsTabs As cTabItems
Attribute m_clsTabs.VB_VarHelpID = -1

' ======================================================================================
' Events
' ======================================================================================
Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event TabClick()

Public Property Get AutoBackColor() As Boolean
    AutoBackColor = m_blnAutoBackColor
End Property

Public Property Let AutoBackColor(ByVal New_AutoBackColor As Boolean)
    m_blnAutoBackColor = New_AutoBackColor
    PropertyChanged "AutoBackColor"
    Call pvDraw
End Property

'/* ������ɫ��Ĭ��Ϊ-1���游�������ɫ���ı䣩 */
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oleBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_oleBackColor = VerifyColor(New_BackColor)
    If Not m_blnAutoBackColor Then UserControl.BackColor = m_oleBackColor
    PropertyChanged "BackColor"
    Call pvDraw
End Property

Public Property Get BorderStyle() As GPTAB_BORDERSTYLE_METHOD
    BorderStyle = m_udtBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As GPTAB_BORDERSTYLE_METHOD)
    m_udtBorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get Enable() As Boolean
    Enable = m_blnEnabled
End Property

Public Property Let Enable(ByVal New_Enable As Boolean)
    m_blnEnabled = New_Enable
    PropertyChanged "Enable"
    Call pvDraw
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
    Dim udtFont As LOGFONT
     
    Set UserControl.Font = New_Font
    Set lblFont.Font = New_Font
    '/* ȡ�õ�ǰ���������ֵĸ߶� */
    m_lngFontHeight = lblFont.Height + 1
    If m_lngFontHeight > m_lIconSizeY Then
       m_lngDefaultTabHeight = m_lngFontHeight + InflateFontHeight
    Else
       m_lngDefaultTabHeight = m_lIconSizeY + InflateIconHeight
    End If
'    If m_lngFontHeight > m_lIconSizeY Then
'       m_cListItems.DefaultListitemHeight = m_lngFontHeight
'    Else
'       m_cListItems.DefaultListitemHeight = m_lIconSizeY
'    End If
    
'    If m_lngFontHeight > m_lngColumnIconHeight Then
'       If m_lngFontHeight > m_cHeader.Height Then
'          m_cHeader.Height = m_lngFontHeight
'       End If
'    Else
'       If m_lngColumnIconHeight > m_cHeader.Height Then
'          m_cHeader.Height = m_lngColumnIconHeight
'       End If
'    End If
    PropertyChanged "Font"
    Call pvDraw
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   UserControl.ForeColor = VerifyColor(New_ForeColor)
   PropertyChanged "ForeColor"
    Call pvDraw
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As cTabItem
    Dim lngI As Long
    Dim lngY As Long
    Dim tR As RECT
    Const ProcName = "HitTest"
    
    On Error GoTo ErrorHandle
    
    ' �ж�����Ƿ����
    If X < 0 Or X > UserControl.ScaleWidth Or Y < 0 Or Y > UserControl.ScaleHeight Then
       Set HitTest = Nothing
       Exit Function
    End If
    ' �ж�����Ƿ�����������
    If Y >= m_udtMainRect.Top Then
       Set HitTest = Nothing
       Exit Function
    End If
    If m_clsTabs Is Nothing Or m_clsTabs.Count <= 0 Then
       Set HitTest = Nothing
       Exit Function
    End If
    Select Case m_udtTabStyle
           Case GpTabRectangle, GpTabRoundRect
             For lngI = 1 To m_clsTabs.Count
                 With m_clsTabs.Item(lngI)
                      tR.Top = .Top
                      tR.Left = .Left
                      tR.Right = .Left + .Width
                      tR.Bottom = .Top + .Height
                 End With
                 If PtInRect(tR, X, Y) <> 0 Then
                    Set HitTest = m_clsTabs.Item(lngI)
                    Exit Function
                 End If
             Next lngI
           Case GpTabTrapezoid
             For lngI = 1 To m_clsTabs.Count
                 With m_clsTabs.Item(lngI)
                      tR.Top = .Top
                      tR.Left = .Left + .Height
                      tR.Right = .Left + .Width
                      tR.Bottom = .Top + .Height
                 End With
                 If PtInRect(tR, X, Y) <> 0 Then
                    Set HitTest = m_clsTabs.Item(lngI)
                    Exit Function
                 End If
                 
'                 With m_clsTabs.Item(lngI)
'                      For lngY = 1 To .Height
'                          If Y = lngY And X >= .Left - .Height - 1 Then
'                             Set HitTest = m_clsTabs.Item(lngI)
'                             Exit Function
'                          End If
'                      Next lngY
'                 End With
             Next lngI
    End Select
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
    End Select
End Function

Public Property Get HotTracking() As Boolean
    HotTracking = m_blnHotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    m_blnHotTracking = New_HotTracking
    PropertyChanged "HotTracking"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let ImageList(New_ImageList As Variant)
    Dim hIml As Long
    
    ' Set the ImageList handle property either from a VB
    ' image list or directly:
    If VarType(New_ImageList) = vbObject Then
       ' Assume VB ImageList control.  Note that unless
       ' some call has been made to an object within a
       ' VB ImageList the image list itself is not
       ' created.  Therefore hImageList returns error. So
       ' ensure that the ImageList has been initialised by
       ' drawing into nowhere:
       On Error Resume Next
       ' Get the image list initialised..
       New_ImageList.ListImages(1).Draw 0, 0, 0, 1
       hIml = New_ImageList.hImageList
       If (Err.Number <> 0) Then
           hIml = 0
       End If
       On Error GoTo 0
    ElseIf VarType(New_ImageList) = vbLong Then
       ' Assume ImageList handle:
       hIml = New_ImageList
    Else
       Err.Raise vbObjectError + 1049, "GpTabs." & App.EXEName, "ImageList property expects ImageList object or long hImageList handle."
    End If
    
    ' If we have a valid image list, then associate it with the control:
    If (hIml <> 0) Then
       m_hIml = hIml
       Call ImageList_GetIconSize(m_hIml, m_lIconSizeX, m_lIconSizeY)
       m_lIconSizeY = m_lIconSizeY + 2
       If m_lngFontHeight > m_lIconSizeY Then
          m_lngDefaultTabHeight = m_lngFontHeight + InflateFontHeight
       Else
          m_lngDefaultTabHeight = m_lIconSizeY + InflateIconHeight
       End If
    End If
End Property

' ���Icon
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

' �����ʽ
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

' ������ʾ
Public Property Get MultiRow() As Boolean
    MultiRow = m_blnMultiRow
End Property

Public Property Let MultiRow(ByVal New_MultiRow As Boolean)
    m_blnMultiRow = New_MultiRow
    PropertyChanged "MultiRow"
    Call pvDraw
End Property

' ѡ�����
Public Property Get Placement() As GPTAB_PLACEMENT_METHOD
    Placement = m_udtPlacement
End Property

Public Property Let Placement(ByVal New_Placement As GPTAB_PLACEMENT_METHOD)
    m_udtPlacement = New_Placement
    PropertyChanged "Placement"
    Call pvDraw
End Property

' ���㹹������������������
Private Sub pvCalculateRect(ByRef DstPoint() As POINTAPI, _
                            ByVal Count As Long, _
                            ByVal Top As Long, _
                            ByVal Left As Long, _
                            ByVal Right As Long, _
                            ByVal Bottom As Long, _
                            ByVal LeftTop As Boolean, _
                            ByVal LeftBottom As Boolean, _
                            ByVal RightBottom As Boolean, _
                            ByVal RightTop As Boolean)
    Dim lngindex            As Long
    
    ReDim DstPoint(Count - 1)
    lngindex = 0
    If LeftTop Then
       With DstPoint(lngindex)
            .X = Left + RoundRectSize
            .Y = Top
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .X = Left
            .Y = Top + RoundRectSize
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .X = Left
            .Y = Top
       End With
       lngindex = lngindex + 1
    End If
    
    If LeftBottom Then
       With DstPoint(lngindex)
            .X = Left
            .Y = Bottom - RoundRectSize
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .X = Left + RoundRectSize
            .Y = Bottom
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .X = Left
            .Y = Bottom
       End With
       lngindex = lngindex + 1
    End If
    If RightBottom Then
       With DstPoint(lngindex)
            .X = Right - RoundRectSize
            .Y = Bottom
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .X = Right
            .Y = Bottom - RoundRectSize
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .X = Right
            .Y = Bottom
       End With
       lngindex = lngindex + 1
    End If
    If RightTop Then
       With DstPoint(lngindex)
            .X = Right
            .Y = Top + RoundRectSize
       End With
       lngindex = lngindex + 1
       With DstPoint(lngindex)
            .X = Right - RoundRectSize
            .Y = Top
       End With
       lngindex = lngindex + 1
    Else
       With DstPoint(lngindex)
            .X = Right
            .Y = Top
       End With
       lngindex = lngindex + 1
    End If
    
    If LeftTop Then
       With DstPoint(lngindex)
            .X = Left + RoundRectSize
            .Y = Top
       End With
    Else
       With DstPoint(lngindex)
            .X = Left
            .Y = Top
       End With
    End If
End Sub

Private Sub pvCalculateRoundPoint(ByRef DstPoint() As POINTAPI, _
                                  ByVal Count As Long, _
                                  ByVal Top As Long, _
                                  ByVal Left As Long, _
                                  ByVal Right As Long, _
                                  ByVal Bottom As Long)
    Dim lngindex As Long
    
    lngindex = 0
    ReDim DstPoint(Count - 1)
    With DstPoint(lngindex)
         .X = Left
         .Y = Bottom
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .X = Left
         .Y = Top + RoundRectSize
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .X = Left + RoundRectSize
         .Y = Top
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .X = Right - RoundRectSize
         .Y = Top
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .X = Right
         .Y = Top + RoundRectSize
    End With
    lngindex = lngindex + 1
    With DstPoint(lngindex)
         .X = Right
         .Y = Bottom
    End With
End Sub

Private Sub pvCalculateSize()
    Dim lngI                 As Long  ' ѭ������
    Dim lngY                 As Long
    Dim lngTabCount          As Long  ' Tab���ܸ���
    Dim lngAllWidth          As Long  ' ���е�Tab�Ŀ��
    Dim lngListIndex         As Long  ' û��Tab������
    Dim lngListTabIndex      As Long  ' һ��Tab������
    Dim lngListWidth         As Long  ' �ۼ�һ��Tab�Ŀ��
    Dim lngManualWidth       As Long  ' ����Tab�Ŀ��
    Dim lngManualHeight      As Long  ' ����Tab�ĸ߶�
    Dim lngWidth             As Long  ' �ؼ��Ŀ��
    Dim lngHeight            As Long  ' �ؼ��ĸ߶�
    Dim lngDiscrepancy       As Long
    Dim tR                   As RECT
    
    If m_udtBorderStyle = GpTabBorderStyleNone Then
       lngDiscrepancy = 0
    Else
       lngDiscrepancy = DiscrepancyHeight
    End If
    Const ProcName = "pvCalculateSize"
    
    On Error GoTo ErrorHandle
    
    lngWidth = UserControl.ScaleWidth - 1
    lngHeight = UserControl.ScaleHeight - 1
    With m_udtMainRect
         .Top = 0
         .Left = 0
         .Right = lngWidth
         .Bottom = lngHeight
    End With
    If m_clsTabs Is Nothing Then
       m_udtMainRect.Top = m_lngDefaultTabHeight
       Exit Sub
    End If
    
    m_lngListCount = 0
    Erase m_aryTabList
    lngManualWidth = m_lngTabFixedWidth \ Screen.TwipsPerPixelX
    lngManualHeight = m_lngTabFixedHeight \ Screen.TwipsPerPixelY
    With m_clsTabs
         lngTabCount = m_clsTabs.Count
         ' ����ÿ��Tab��ʵ����С���
         For lngI = 1 To lngTabCount
             Call DrawTextEx(UserControl.hdc, .Item(lngI).Caption & vbNullChar, -1, tR, _
                             DT_CALCRECT Or DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, _
                             m_udtDrawTextParams)
             If m_udtTabStyle = GpTabTrapezoid Then
                .Item(lngI).DefaultWidth = tR.Right - tR.Left + InflateFontWidth + m_lngDefaultTabHeight
             Else
                .Item(lngI).DefaultWidth = tR.Right - tR.Left + InflateFontWidth
             End If
             lngAllWidth = lngAllWidth + .Item(lngI).DefaultWidth
         Next lngI
         Select Case m_udtPlacement
                Case GpTabPlacementTopleft
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                           ' ����ÿ��Tab�ĸ߶ȺͿ��
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = m_lngDefaultTabHeight
                               .Item(lngI).Width = .Item(lngI).DefaultWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                              ' ��������
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(0)
                              ' ����,������ÿ��Tab������
                              m_lngListCount = 0
                              lngListTabIndex = 0
                              lngListWidth = .Item(1).Width
                           Else
                              ' ����������Ķ���
                              m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ����ÿ��Tab������
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                              Next lngI
                           End If
                         Case GpTabnonJustified
                           ' ����ÿ��Tab�ĸ߶ȺͿ��
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = m_lngDefaultTabHeight
                               .Item(lngI).Width = .Item(lngI).DefaultWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                              ' ��������
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(0)
                              ' ����,������ÿ��Tab������
                              m_lngListCount = 0
                              lngListTabIndex = 0
                              lngListWidth = .Item(1).Width
                              For lngI = 1 To lngTabCount
                                  ReDim Preserve m_aryTabList(m_lngListCount).list(lngListTabIndex)
                                  With m_aryTabList(m_lngListCount).list(lngListTabIndex)
                                       .Index = lngI
                                  End With
                                  If lngListTabIndex > 0 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                                  If lngI + 1 <= lngTabCount Then
                                     lngListWidth = lngListWidth + .Item(lngI + 1).Width
                                  End If
                                  lngListTabIndex = lngListTabIndex + 1
                                  If lngListWidth > lngWidth Then
                                     ' �洢ÿ����Tab�ĸ���
                                     m_aryTabList(m_lngListCount).Count = lngListTabIndex
                                     lngListTabIndex = 0
                                     m_lngListCount = m_lngListCount + 1
                                     ReDim Preserve m_aryTabList(m_lngListCount)
                                     ReDim Preserve m_aryTabList(m_lngListCount).list(0)
                                     m_aryTabList(m_lngListCount).Count = 1
                                     lngListWidth = .Item(lngI + 1).Width
                                  End If
                              Next lngI
                              ' ����ÿ�еĸ߶�
                              For lngI = m_lngListCount To 0 Step -1
                                  For lngY = 0 To m_aryTabList(lngI).Count - 1
                                      m_clsTabs.Item(m_aryTabList(lngI).list(lngY).Index).Top = (m_lngListCount - lngI) * m_lngDefaultTabHeight + lngDiscrepancy
                                  Next lngY
                              Next lngI
                              m_lngListCount = m_lngListCount + 1
                              ' ����������Ķ���
                              m_udtMainRect.Top = m_lngListCount * m_lngDefaultTabHeight + lngDiscrepancy
                           Else
                              ' ����������Ķ���
                              m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ����ÿ��Tab������
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                              Next lngI
                           End If
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ����ÿ��Tab�ĸ߶ȺͿ��
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                              ' ����Tab������
                              m_lngListCount = CLng(lngAllWidth \ lngWidth)
                              If lngAllWidth Mod lngWidth > 0 Then m_lngListCount = m_lngListCount + 1
                              m_udtMainRect.Top = m_lngListCount * lngManualHeight + lngDiscrepancy
                              ' ��������
                              ReDim m_aryTabList(m_lngListCount - 1)
                              For lngI = 0 To m_lngListCount - 1
                                  ReDim m_aryTabList(lngI).list(0)
                              Next lngI
                              ' ����,������ÿ��Tab������
                              lngListIndex = 0
                              lngListTabIndex = 0
                              lngListWidth = .Item(1).Width
                              For lngI = 1 To lngTabCount
                                  ReDim Preserve m_aryTabList(lngListIndex).list(lngListTabIndex)
                                  With m_aryTabList(lngListIndex).list(lngListTabIndex)
                                       .Index = lngI
                                  End With
                                  .Item(lngI).Top = (m_lngListCount - lngListIndex - 1) * m_lngDefaultTabHeight + 1
                                  If lngListTabIndex > 0 Then
                                     .Item(lngI).Left = .Item(lngI - 1).Left + .Item(lngI - 1).Width + TabsInterval
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                                  If lngI + 1 <= lngTabCount Then
                                     lngListWidth = lngListWidth + .Item(lngI + 1).Width
                                  End If
                                  If lngListWidth > lngWidth Then
                                     lngListIndex = lngListIndex + 1
                                     lngListTabIndex = 0
                                  Else
                                     lngListTabIndex = lngListTabIndex + 1
                                  End If
                                  ' �洢ÿ����Tab�ĸ���
                                  If lngListIndex <= m_lngListCount - 1 Then m_aryTabList(lngListIndex).Count = lngListTabIndex
                              Next lngI
                           Else
                              ' ����������Ķ���
                              m_udtMainRect.Top = lngManualHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ����ÿ��Tab������
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = (lngManualWidth + TabsInterval) * (lngI - 1)
                                  Else
                                     .Item(lngI).Left = 0
                                  End If
                              Next lngI
                           End If
                  End Select
                Case GpTabPlacementTopRight
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Top = m_lngDefaultTabHeight + lngDiscrepancy
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                         Case GpTabnonJustified
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ����ÿ��Tab�ĸ߶ȺͿ��
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                           Else
                              ' ����������Ķ���
                              m_udtMainRect.Top = lngManualHeight + lngDiscrepancy
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ����ÿ��Tab������
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Top = lngDiscrepancy
                                  If lngI > 1 Then
                                     .Item(lngI).Left = m_udtMainRect.Right - 10 - (lngManualWidth + TabsInterval) * lngI
                                  Else
                                     .Item(lngI).Left = m_udtMainRect.Right - 10 - .Item(lngI).Width
                                  End If
                              Next lngI
                           End If
                  End Select
                Case GpTabPlacementLeftTop
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Left = lngManualWidth
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                         Case GpTabnonJustified
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ����ÿ��Tab�ĸ߶ȺͿ��
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                           Else
                              ' ����������Ķ���
                              m_udtMainRect.Left = lngManualWidth
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ����ÿ��Tab������
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Left = 0
                                  If lngI > 1 Then
                                     .Item(lngI).Top = m_udtMainRect.Top + lngManualHeight * (lngI - 1)
                                  Else
                                     .Item(lngI).Top = 0
                                  End If
                              Next lngI
                           End If
                  End Select
                Case GpTabPlacementLeftBottom
                Case GpTabPlacementBottomLeft
                Case GpTabPlacementBottomRight
                Case GpTabPlacementRightTop
                Case GpTabPlacementRightBottom
                  If lngTabCount <= 0 Then
                     m_udtMainRect.Right = m_udtMainRect.Right - lngManualWidth
                     Exit Sub
                  End If
                  Select Case m_udtTabWidthStyle
                         Case GpTabJustified
                         Case GpTabnonJustified
                         Case GpTabFixed
                           lngAllWidth = lngManualWidth * lngTabCount
                           ' ����ÿ��Tab�ĸ߶ȺͿ��
                           For lngI = 1 To lngTabCount
                               .Item(lngI).Height = lngManualHeight
                               .Item(lngI).Width = lngManualWidth
                           Next lngI
                           If lngAllWidth > lngWidth Then
                           Else
                              ' ����������Ķ���
                              m_udtMainRect.Left = lngManualWidth
                              m_lngCurrentList = 0
                              ReDim m_aryTabList(0)
                              ReDim m_aryTabList(0).list(lngTabCount - 1)
                              m_aryTabList(0).Count = lngTabCount
                              ' ����ÿ��Tab������
                              For lngI = 1 To lngTabCount
                                  With m_aryTabList(0).list(lngI - 1)
                                       .Index = lngI - 1
                                  End With
                                  .Item(lngI).Left = m_udtMainRect.Right
                                  If lngI > 1 Then
                                     .Item(lngI).Top = m_udtMainRect.Top + lngManualHeight * (lngI - 1)
                                  Else
                                     .Item(lngI).Top = 0
                                  End If
                              Next lngI
                           End If
                  End Select
         End Select
    End With
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub pvCalculateTrapezoidPoint(ByRef DstPoint() As POINTAPI, _
                                      ByVal Count As Long, _
                                      ByVal Top As Long, _
                                      ByVal Left As Long, _
                                      ByVal Right As Long, _
                                      ByVal Bottom As Long)
    Dim lngindex As Long
    
    Select Case m_udtPlacement
           Case GpTabPlacementTopleft
             lngindex = 0
             ReDim DstPoint(Count - 1)
             With DstPoint(lngindex)
                  .X = Left
                  .Y = Bottom
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Left
                  .Y = Top + RoundRectSize
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Left + RoundRectSize
                  .Y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right - (Bottom - Top)
                  .Y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right
                  .Y = Bottom
             End With
           Case GpTabPlacementTopRight
             lngindex = 0
             ReDim DstPoint(Count - 1)
             With DstPoint(lngindex)
                  .X = Left
                  .Y = Bottom
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Left + Bottom - Top
                  .Y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right - RoundRectSize
                  .Y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right
                  .Y = Top + RoundRectSize
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right
                  .Y = Bottom
             End With
           Case GpTabPlacementBottomLeft
           Case GpTabPlacementBottomRight
           Case GpTabPlacementLeftTop
           Case GpTabPlacementLeftBottom
           Case GpTabPlacementRightTop
           Case GpTabPlacementRightBottom
             lngindex = 0
             ReDim DstPoint(Count - 1)
             With DstPoint(lngindex)
                  .X = Left
                  .Y = Top
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right
                  .Y = Top + Right - Left
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right
                  .Y = Bottom - RoundRectSize
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Right - RoundRectSize
                  .Y = Bottom
             End With
             lngindex = lngindex + 1
             With DstPoint(lngindex)
                  .X = Left
                  .Y = Bottom
             End With
    End Select
End Sub

' ���ƿؼ�����
Private Sub pvDraw()
    Dim lngI                As Long
    Dim lngTop1             As Long
    Dim lngTop2             As Long
    Dim lngBottom1          As Long
    Dim lngBottom2          As Long
    Dim lngHeight           As Long
    Dim lngPointCount       As Long
    Dim lngBrush            As Long
    Dim lngPen              As Long
    Dim lngXPColor          As Long
    Dim lngStepXP           As Single
    Dim blnHover            As Boolean
    Dim blnLeftTop          As Boolean
    Dim blnLeftBottom       As Boolean
    Dim blnRightBottom      As Boolean
    Dim blnRightTop         As Boolean
    Dim tBR                 As RECT
    Dim udtPointA()         As POINTAPI
    Dim udtPointB()         As POINTAPI
    Const ProcName = "pvDraw"
    
    On Error GoTo ErrorHandle
    If m_lngDefaultTabHeight <= 0 Then Exit Sub
    With UserControl
         ' ���ÿؼ�������ɫ
         If m_blnAutoBackColor Then
            .BackColor = .Ambient.BackColor
         Else
            .BackColor = m_oleBackColor
         End If
         ' ���
         .Cls
    End With
    
    ' �ؼ�Win32���
    If m_udtStyle = GpTabStyleStandard Then
       ' ����ˢ��
       lngBrush = CreateSolidBrush(TranslateColor(m_oleTabColorActive))
       ' ���
       Call FillRect(UserControl.hdc, m_udtMainRect, lngBrush)
       ' ɾ����ˢ
       Call DeleteObject(lngBrush): lngBrush = 0
       If m_udtBorderStyle = GpTabBorderStyle3D Then
          Call DrawEdge(UserControl.hdc, m_udtMainRect, EDGE_RAISED, BF_RECT)
       ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
          Call DrawEdge(UserControl.hdc, m_udtMainRect, BDR_RAISEDINNER, BF_RECT)
       End If
    ' �ؼ�WinXP���
    ElseIf m_udtStyle = GpTabStyleWinXP Then
       If m_udtBorderStyle = GpTabBorderStyleNone Then
          ' ����ˢ��
          lngBrush = CreateSolidBrush(TranslateColor(XPFlatTabColorActive))
          ' ���
          Call FillRect(UserControl.hdc, m_udtMainRect, lngBrush)
          ' ɾ����ˢ
          Call DeleteObject(lngBrush): lngBrush = 0
          With m_udtMainRect
               tBR.Top = .Top
               tBR.Left = .Left
               tBR.Right = .Right
               tBR.Bottom = .Top + 3
          End With
          ' ����ˢ��
          lngBrush = CreateSolidBrush(TranslateColor(XPFlatBorderColor))
          ' ���
          Call FillRect(UserControl.hdc, tBR, lngBrush)
          ' ɾ����ˢ
          Call DeleteObject(lngBrush): lngBrush = 0
       Else
          blnLeftTop = False
          blnLeftBottom = True
          blnRightBottom = True
          blnRightTop = True
          lngPointCount = 8
          ' ���õ�ǰ����,�߿���ɫ
          'lngPen = CreatePen(PS_SOLID, 1, TranslateColor(XPBorderColor))
          'Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
          With m_udtMainRect
               Call pvCalculateRect(udtPointA, lngPointCount, .Top, .Left, .Right, .Bottom, _
                                    blnLeftTop, blnLeftBottom, blnRightBottom, blnRightTop)
               lngHeight = .Bottom - .Top
               lngStepXP = 25 / lngHeight
               lngTop1 = 1
               lngTop2 = 2
               lngBottom1 = lngHeight - 1
               lngBottom2 = lngHeight
               For lngI = 1 To lngHeight
                   Select Case lngI
                            'HERE
                          Case lngTop1
                            lngXPColor = vbWhite 'BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI)
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftTop, .Left + 2, .Left), .Top, _
                                          IIf(blnRightTop, .Right - 2, .Right), .Top, _
                                          lngXPColor)
                          Case lngTop2
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftTop, .Left + 1, .Left), .Top + 1, _
                                          IIf(blnRightTop, .Right - 1, .Right), .Top + 1, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                          Case lngBottom1
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftBottom, .Left + 2, .Left), .Bottom - 1, _
                                          IIf(blnRightBottom, .Right - 2, .Right), .Bottom - 1, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                          Case lngBottom2
                            'Call DrawLine(UserControl.hdc, IIf(blnLeftBottom, .Left + 1, .Left), .Bottom - 2, _
                                          IIf(blnRightBottom, .Right - 1, .Right), .Bottom - 2, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                          Case Else
                            'Call DrawLine(UserControl.hdc, .Left, lngI + .Top - 1, .Right, lngI + .Top - 1, _
                                          BrightnessColor(m_lngXPFaceColor, -lngStepXP * lngI))
                   End Select
               Next lngI
               Call Polyline(UserControl.hdc, udtPointA(0), lngPointCount)
          End With
       End If
    End If
    If m_clsTabs.Count <= 0 Then
       Call pvDrawTab("", 0, DiscrepancyHeight, 60, m_lngDefaultTabHeight, m_oleTabColorActive, lngXPColor, True, False)
    Else
       For lngI = m_clsTabs.Count To 1 Step -1
           blnHover = False
           With m_clsTabs.Item(lngI)
                If Not (m_clsHoverTab Is Nothing) Then
                   If m_clsHoverTab.Index = lngI Then blnHover = True
                End If
                Call pvDrawTab(.Caption, .Left, .Top, .Width, .Height, m_oleTabColorActive, lngXPColor, .Selected, blnHover)
           End With
       Next lngI
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub pvDrawTab(ByVal Caption As String, _
                      ByVal Left As Long, _
                      ByVal Top As Long, _
                      ByVal Width As Long, _
                      ByVal Height As Long, _
                      ByVal TabColor As Long, _
                      ByVal TranXPColor As Long, _
                      ByVal Selected As Boolean, _
                      ByVal Hover As Boolean)
    Dim lngTop1             As Long
    Dim lngTop2             As Long
    Dim lngEdgeStyle        As Long
    Dim lngEdgeFlag         As Long
    Dim lngPen              As Long
    Dim lngOldBrush         As Long
    Dim lngXPBorderBrush    As Long
    Dim lngTabBorderBrush   As Long
    
    Dim lngFlatBrush        As Long
    Dim lngFlatActiveBrush  As Long
    Dim lngFlatHoverBrush   As Long
    Dim lngFlatBorderBrush  As Long
    
    Dim lngI                As Long
    Dim lngLightColor       As Long
    Dim lngHighLightColor   As Long
    Dim lngShadowColor      As Long
    Dim lngDarkShadowColor  As Long
    Dim lngXPColor          As Long
    Dim lngStepXP           As Single
    Dim udtPointA()         As POINTAPI
    Dim udtPointB()         As POINTAPI
    Dim udtTabRect          As RECT
    Dim udtCaptionRect      As RECT
    Const ProcName = "pvDrawTab"
    
    On Error GoTo ErrorHandle
    
    ' ȡ��ϵͳ��ʾ�������ɫ
    lngShadowColor = GetSysColor(COLOR_BTNSHADOW)
    lngLightColor = GetSysColor(COLOR_BTNLIGHT)
    lngDarkShadowColor = GetSysColor(COLOR_BTNDKSHADOW)
    lngHighLightColor = GetSysColor(COLOR_BTNHIGHLIGHT)
    
    ' ������ˢ
    lngXPBorderBrush = CreateSolidBrush(TranslateColor(XPBorderColor))
    
    lngFlatBrush = CreateSolidBrush(TranslateColor(XPFlatTabColor))
    lngFlatActiveBrush = CreateSolidBrush(TranslateColor(XPFlatTabColorActive))
    lngFlatHoverBrush = CreateSolidBrush(TranslateColor(XPFlatTabColorHover))
    lngFlatBorderBrush = CreateSolidBrush(TranslateColor(XPFlatBorderColor))
    
    lngTabBorderBrush = CreateSolidBrush(TranslateColor(TabColor))
    lngOldBrush = SelectObject(UserControl.hdc, lngTabBorderBrush)
    With udtTabRect
         .Left = Left
         .Top = Top
         .Right = Left + Width
         .Bottom = Top + Height
    End With
    
    ' �ؼ�Win32���
    If m_udtStyle = GpTabStyleStandard Then
       Select Case m_udtTabStyle
              Case GpTabRectangle
                If m_udtBorderStyle = GpTabBorderStyle3D Then
                   lngEdgeStyle = EDGE_RAISED
                   'lngEdgeStyle = BDR_RAISEDINNER
                   If Selected Then udtTabRect.Bottom = Top + Height + 2
                ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                   lngEdgeStyle = BDR_RAISEDINNER
                   If Selected Then udtTabRect.Bottom = Top + Height + 1
                End If
                Call FillRect(UserControl.hdc, udtTabRect, lngTabBorderBrush)
                If m_udtBorderStyle <> GpTabBorderStyleNone And Selected = True Then Call DrawEdge(UserControl.hdc, udtTabRect, lngEdgeStyle, BF_LEFT Or BF_TOP Or BF_RIGHT)
              Case GpTabRoundRect
                ReDim udtPointB(5)
                With udtTabRect
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        .Left = .Left + 1
                        .Bottom = .Bottom + 2
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        .Bottom = .Bottom + 1
                     End If
                     ' ȡ��Բ�ǵĸ����������
                     Call pvCalculateRoundPoint(udtPointA, 6, .Top, _
                                                .Left, .Right, _
                                                .Bottom)
                     ' ���
                     lngPen = CreatePen(PS_SOLID, 1, lngHighLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call SelectObject(UserControl.hdc, lngTabBorderBrush)
                     Call Polygon(UserControl.hdc, udtPointA(0), 6)
                     lngPen = CreatePen(PS_SOLID, 1, lngLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call pvInflatePoint(udtPointB(0), udtPointA(0), 1, 0)
                     Call pvInflatePoint(udtPointB(1), udtPointA(1), 1, 1)
                     Call pvInflatePoint(udtPointB(2), udtPointA(2), 1, 1)
                     Call pvInflatePoint(udtPointB(3), udtPointA(3), -1, 1)
                     Call pvInflatePoint(udtPointB(4), udtPointA(4), -1, 1)
                     Call pvInflatePoint(udtPointB(5), udtPointA(5), -1, 0)
                     Call Polyline(UserControl.hdc, udtPointB(0), 6)
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        ' ȥ�ױ�
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ����Ӱ
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 3, Top, .Right - 1, .Top + 3, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 1, .Top + 2, .Right - 1, .Bottom - 1, TranslateColor(lngShadowColor))
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        ' ȥ�ױ�
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ����Ӱ
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngShadowColor))
                     End If
                End With
              Case GpTabTrapezoid
                ReDim udtPointB(4)
                With udtTabRect
                     If m_udtBorderStyle = GpTabBorderStyleNone Then
                        ' ȡ��Բ�ǵĸ����������
                        Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom)
                        ' ���
                        lngPen = CreatePen(PS_SOLID, 1, lngHighLightColor)
                        Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                        Call SelectObject(UserControl.hdc, lngTabBorderBrush)
                        Call Polygon(UserControl.hdc, udtPointA(0), 5)
                     Else
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        .Left = .Left + 1
                        .Bottom = .Bottom + 2
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        .Bottom = .Bottom + 1
                     End If
                     ' ȡ��Բ�ǵĸ����������
                     Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom)
                     ' ���
                     lngPen = CreatePen(PS_SOLID, 1, lngHighLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call SelectObject(UserControl.hdc, lngTabBorderBrush)
                     Call Polygon(UserControl.hdc, udtPointA(0), 5)
                     lngPen = CreatePen(PS_SOLID, 1, lngLightColor)
                     Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                     Call pvInflatePoint(udtPointB(0), udtPointA(0), 1, 0)
                     Call pvInflatePoint(udtPointB(1), udtPointA(1), 1, 1)
                     Call pvInflatePoint(udtPointB(2), udtPointA(2), -1, 1)
                     Call pvInflatePoint(udtPointB(3), udtPointA(3), -1, 1)
                     Call pvInflatePoint(udtPointB(4), udtPointA(4), -1, 0)
                     Call Polyline(UserControl.hdc, udtPointB(0), 5)
                     If m_udtBorderStyle = GpTabBorderStyle3D Then
                        ' ȥ�ױ�
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 1, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ����Ӱ
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 3, .Top, .Right - 1, .Top + 3, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngDarkShadowColor))
                        Call DrawLine(UserControl.hdc, .Right - 1, .Top + 2, .Right - 1, .Bottom - 1, TranslateColor(lngShadowColor))
                     ElseIf m_udtBorderStyle = GpTabBorderStyle3DThin Then
                        ' ȥ�ױ�
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom, .Right, .Bottom, TranslateColor(m_oleTabColorActive))
                        Call DrawLine(UserControl.hdc, .Left + 2, .Bottom + 1, .Right, .Bottom + 1, TranslateColor(m_oleTabColorActive))
                        ' ����Ӱ
                        Call DrawLine(UserControl.hdc, .Right - 2, .Top, .Right, .Top + 2, TranslateColor(lngShadowColor))
                        Call DrawLine(UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, TranslateColor(lngShadowColor))
                     End If
                     End If
                End With
       End Select
    ' �ؼ�WinXP���
    'LAST HERE -> LOOK FOR THE BLUE EDGE ON TAB HOW TO REMOVE
    ElseIf m_udtStyle = GpTabStyleWinXP Then
       Select Case m_udtTabStyle
              Case GpTabRectangle
                If m_udtBorderStyle = GpTabBorderStyleNone Then
                Else
                   With udtTabRect
                        ' ���
                        lngStepXP = 25 / Height
                        For lngI = Height To 1 Step -1
                            Call DrawLine(UserControl.hdc, .Left + 1, lngI + .Top, .Right, lngI + .Top, _
                                          BrightnessColor(TranXPColor, lngStepXP * lngI))
                        Next lngI
                        lngXPColor = BrightnessColor(TranXPColor, lngStepXP * 1)
                        Call SelectObject(UserControl.hdc, lngXPBorderBrush)
                        .Right = .Right + 1  ' ʹ����С
                        .Bottom = Top + Height + 1
                        ' ���߿�
                        Call FrameRect(UserControl.hdc, udtTabRect, lngXPBorderBrush)
                        If Selected Then
                           ' ȥ����
                           'Call DrawLine(UserControl.hdc, .Left, .Top, .Right, .Top, lngXPColor)
                           ' ��������
                           'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 2, Top - 2, &HFF6633)
                           'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right - 1, Top - 1, &HFF855D)
                           'Call DrawLine(UserControl.hdc, .Left, Top, .Right, Top, &HFEA588)
                           'Call DrawLine(UserControl.hdc, .Left - 1, Top + 1, .Right - 1, Top + 1, &HFFC5B2)
                           ' ��ѡ��ʱȥ�ױ�
                           'Call DrawLine(UserControl.hdc, .Left + 1, .Bottom - 1, .Right - 1, .Bottom - 1, TranXPColor)
                        Else
                           If Hover Then
                              ' ȥ����
                              'Call DrawLine(UserControl.hdc, .Left, .Top, .Right, .Top, lngXPColor)
                              ' Hover line
                              'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 2, Top - 2, &H138DEB)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right - 1, Top - 1, &H3399FF)
                              'Call DrawLine(UserControl.hdc, .Left, Top, .Right, Top, &H66CCFF)
                              'Call DrawLine(UserControl.hdc, .Left - 1, Top + 1, .Right - 1, Top + 1, &H9DDBFF)
                           End If
                        End If
                   End With
                End If
              Case GpTabRoundRect
                If m_udtBorderStyle = GpTabBorderStyleNone Then
                Else
                   lngStepXP = 25 / Height
                   With udtTabRect
                        ' ���
                        lngTop1 = .Top
                        lngTop2 = .Top + 1
                        For lngI = Height To 1 Step -1
                            If lngI = lngTop1 Then
                               Call DrawLine(UserControl.hdc, .Left + 2, lngI + .Top, .Right - 2, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                            ElseIf lngI = lngTop2 Then
                               Call DrawLine(UserControl.hdc, .Left + 1, lngI + .Top, .Right - 1, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                            Else
                               Call DrawLine(UserControl.hdc, .Left, lngI + .Top, .Right, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                            End If
                        Next lngI
                        lngXPColor = BrightnessColor(TranXPColor, lngStepXP * .Top)
                        ' ȡ��Բ�ǵĸ����������
                        Call pvCalculateRoundPoint(udtPointA, 6, .Top, .Left, .Right, .Bottom)
                        Call SelectObject(UserControl.hdc, lngXPBorderBrush)
                        ' ���߿�
                        Call Polyline(UserControl.hdc, udtPointA(0), 6)
                        If Selected Then
                           ' ȥ����
                           'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                           ' ��������
                           'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 1, Top - 2, &HFF6633)
                           'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right, Top - 1, &HFF855D)
                           'Call DrawLine(UserControl.hdc, .Left, Top, .Right + 1, Top, &HFEA588)
                           'Call DrawLine(UserControl.hdc, .Left - 1, Top + 1, .Right, Top + 1, &HFFC5B2)
                        Else
                           If Hover Then
                              ' ȥ����
                              'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                              ' Hover line
                              'Call DrawLine(UserControl.hdc, .Left + 2, Top - 2, .Right - 1, Top - 2, &H138DEB)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top - 1, .Right, Top - 1, &H3399FF)
                              'Call DrawLine(UserControl.hdc, .Left, Top, .Right + 1, Top, &H66CCFF)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top + 1, .Right, Top + 1, &H9DDBFF)
                           End If
                           ' û��ѡ��ʱ�ӵױ�
                           Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, XPBorderColor)
                        End If
                   End With
                End If
              Case GpTabTrapezoid
                If m_udtBorderStyle = GpTabBorderStyleNone Then
                   ReDim udtPointB(4)
                   With udtTabRect
                        ' ȡ��Բ�ǵĸ����������
                        Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom - 1)
                        ' ���
                        If Selected Then
                          
                           lngPen = CreatePen(PS_SOLID, 1, vbWhite)
                           Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                           Call SelectObject(UserControl.hdc, vbWhite)
                           Call pvInflatePoint(udtPointB(0), udtPointA(0), 2, 2)
                           Call pvInflatePoint(udtPointB(1), udtPointA(1), 2, 2)
                           Call pvInflatePoint(udtPointB(2), udtPointA(2), -2, 2)
                           Call pvInflatePoint(udtPointB(3), udtPointA(3), -2, 2)
                           Call pvInflatePoint(udtPointB(4), udtPointA(4), -2, 2)
                           Call Polygon(UserControl.hdc, udtPointB(0), 5)
                           
                           'Call DrawLine(UserControl.hdc, .Left, .Bottom + 2, .Right, .Bottom + 2, XPFlatTabColorActive)
                        Else
                           lngPen = CreatePen(PS_SOLID, 0, &H808080)
                           Call DeleteObject(SelectObject(UserControl.hdc, lngPen))
                           Call SelectObject(UserControl.hdc, &H808080)
                           Call Polygon(UserControl.hdc, udtPointA(0), 5)
                        End If
                   End With
                Else
                   lngStepXP = 25 / Height
                   With udtTabRect
                        ' ���
                        lngTop1 = .Top
                        lngTop2 = .Top + 1
                        For lngI = Height To 1 Step -1
                            If lngI = lngTop1 Then
                                If Selected Then
                               Call DrawLine(UserControl.hdc, .Left + Height, lngI + .Top, .Right - 2, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                                End If
                            ElseIf lngI = lngTop2 Then
                                If Selected Then
                               Call DrawLine(UserControl.hdc, .Left + Height - 1, lngI + .Top, .Right - 1, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                                End If
                            Else
                                If Selected Then
                               Call DrawLine(UserControl.hdc, .Left + (.Bottom - lngI), lngI + .Top, .Right, lngI + .Top, _
                                             BrightnessColor(TranXPColor, lngStepXP * lngI))
                                End If
                            End If
                        Next lngI
                        ' ȡ��Բ�ǵĸ����������
                        Call pvCalculateTrapezoidPoint(udtPointA, 5, .Top, .Left, .Right, .Bottom)
                        Call SelectObject(UserControl.hdc, lngXPBorderBrush)
                        ' ���߿�
                        Call Polyline(UserControl.hdc, udtPointA(0), 5)
                        If Selected Then
                           ' ȥ����
                           'HERE
                           'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                           ' ��������
                           'Call DrawLine(UserControl.hdc, .Left + 2, Top, .Right - 1, Top, &HFF6633)
                           'Call DrawLine(UserControl.hdc, .Left + 1, Top + 1, .Right, Top + 1, &HFF855D)
                           'Call DrawLine(UserControl.hdc, .Left, Top + 2, .Right + 1, Top + 2, &HFEA588)
                           'Call DrawLine(UserControl.hdc, .Left - 1, Top + 3, .Right, Top + 3, &HFFC5B2)
                        Else
                           If Hover Then
                              ' ȥ����
                              'Call DrawLine(UserControl.hdc, .Left + 2, .Top, .Right - 2, .Top, lngXPColor)
                              ' Hover line
                              'Call DrawLine(UserControl.hdc, .Left + 2, Top, .Right - 1, Top, &H138DEB)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top + 1, .Right, Top + 1, &H3399FF)
                              'Call DrawLine(UserControl.hdc, .Left, Top + 2, .Right + 1, Top + 2, &H66CCFF)
                              'Call DrawLine(UserControl.hdc, .Left + 1, Top + 3, .Right, Top + 3, &H9DDBFF)
                           End If
                           ' û��ѡ��ʱ�ӵױ�
                           Call DrawLine(UserControl.hdc, .Left + 1, .Bottom, .Right, .Bottom, XPBorderColor)
                        End If
                   End With
                End If
       End Select
    End If
    
    If m_udtTabStyle = GpTabTrapezoid Then
       With udtTabRect
            .Left = .Left + m_lngDefaultTabHeight
       End With
    End If
    ' Draw Caption
    Call DrawTextEx(UserControl.hdc, Caption & vbNullString, -1, udtTabRect, _
                    pvGetCaptionFlags(GpTabCaptionRight), m_udtDrawTextParams)
    If lngOldBrush <> 0 Then Call SelectObject(UserControl.hdc, lngOldBrush): lngOldBrush = 0
    Call DeleteObject(lngXPBorderBrush)
    Call DeleteObject(lngTabBorderBrush)
    Call DeleteObject(lngFlatBrush)
    Call DeleteObject(lngFlatActiveBrush)
    Call DeleteObject(lngFlatHoverBrush)
    Call DeleteObject(lngFlatBorderBrush)
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             If lngOldBrush <> 0 Then Call SelectObject(UserControl.hdc, lngOldBrush)
             Call DeleteObject(lngXPBorderBrush)
             Call DeleteObject(lngTabBorderBrush)
             Call DeleteObject(lngFlatBrush)
             Call DeleteObject(lngFlatActiveBrush)
             Call DeleteObject(lngFlatHoverBrush)
             Call DeleteObject(lngFlatBorderBrush)
    End Select
End Sub

Private Function pvGetCaptionFlags(ByVal Alignment As GPTAB_ALIGNMENT_METHOD) As Long
    Select Case Alignment
           Case GpTabCaptionLeft
             pvGetCaptionFlags = DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_LEFT Or DT_VCENTER
           Case GpTabCaptionRight
             pvGetCaptionFlags = DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_RIGHT Or DT_VCENTER
           Case GpTabCaptionCenter
             pvGetCaptionFlags = DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CENTER Or DT_VCENTER
    End Select
End Function

Private Sub pvInflatePoint(ByRef DstPoint As POINTAPI, ByRef SrcPoint As POINTAPI, ByVal X As Long, ByVal Y As Long)
    With DstPoint
         .X = SrcPoint.X + X
         .Y = SrcPoint.Y + Y
    End With
End Sub

Private Sub pvSetColor()
    If m_udtXPColorScheme = GpTabUseWindows Then
       m_lngXPFaceColor = BrightnessColor(GetSysColor(COLOR_BTNFACE), &H30)
    Else
       m_lngXPFaceColor = BrightnessColor(m_oleTabColorActive, &H30)
    End If
End Sub

' ˢ��
Public Sub Refresh()
    Call pvDraw
End Sub

Public Property Get SelectTabItem() As cTabItem
    If m_clsSelectTab Is Nothing Then
       Set SelectTabItem = Nothing
    Else
       Set SelectTabItem = m_clsSelectTab
    End If
End Property

Public Property Get Style() As GPTAB_STYLE_METHOD
    Style = m_udtStyle
End Property

Public Property Let Style(ByVal New_Style As GPTAB_STYLE_METHOD)
    m_udtStyle = New_Style
    PropertyChanged "Style"
    Call pvDraw
End Property

Public Property Get TabBorderColor() As OLE_COLOR
    TabBorderColor = m_oleTabBorderColor
End Property

Public Property Let TabBorderColor(ByVal New_TabBorderColor As OLE_COLOR)
    m_oleTabBorderColor = New_TabBorderColor
    PropertyChanged "TabBorderColor"
End Property

Public Property Get TabColor() As OLE_COLOR
    TabColor = m_oleTabColor
End Property

Public Property Let TabColor(ByVal New_TabColor As OLE_COLOR)
    m_oleTabColor = New_TabColor
    PropertyChanged "TabColor"
    Call pvDraw
End Property

Public Property Get TabColorActive() As OLE_COLOR
    TabColorActive = m_oleTabColorActive
End Property

Public Property Let TabColorActive(ByVal New_TabColorActive As OLE_COLOR)
    m_oleTabColorActive = New_TabColorActive
    PropertyChanged "TabColorActive"
    Call pvSetColor
    Call pvDraw
End Property

Public Property Get TabColorHover() As OLE_COLOR
    TabColorHover = m_oleTabColorHover
End Property

Public Property Let TabColorHover(ByVal New_TabColorHover As OLE_COLOR)
    m_oleTabColorHover = New_TabColorHover
    PropertyChanged "TabColorHover"
End Property

Public Property Get TabFixedHeight() As Long
    TabFixedHeight = m_lngTabFixedHeight
End Property

Public Property Let TabFixedHeight(ByVal New_TabFixedHeight As Long)
    m_lngTabFixedHeight = New_TabFixedHeight
    PropertyChanged "TabFixedHeight"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get TabFixedWidth() As Long
    TabFixedWidth = m_lngTabFixedWidth
End Property

Public Property Let TabFixedWidth(ByVal New_TabFixedWidth As Long)
    m_lngTabFixedWidth = New_TabFixedWidth
    PropertyChanged "TabFixedWidth"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get Tabs() As cTabItems
    Set Tabs = m_clsTabs
End Property

Public Property Get TabStyle() As GPTAB_TABSTYLE_METHOD
    TabStyle = m_udtTabStyle
End Property

Public Property Let TabStyle(ByVal New_TabStyle As GPTAB_TABSTYLE_METHOD)
    m_udtTabStyle = New_TabStyle
    PropertyChanged "TabStyle"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get TabWidthStyle() As GPTAB_TABWIDTHSTYLE_METHOD
    TabWidthStyle = m_udtTabWidthStyle
End Property

Public Property Let TabWidthStyle(ByVal New_TabWidthStyle As GPTAB_TABWIDTHSTYLE_METHOD)
    m_udtTabWidthStyle = New_TabWidthStyle
    PropertyChanged "TabWidthStyle"
    Call pvCalculateSize
    Call pvDraw
End Property

Public Property Get XPColorScheme() As GPTAB_XPCOLORSCHEME_METHOD
    XPColorScheme = m_udtXPColorScheme
End Property

Public Property Let XPColorScheme(ByVal New_XPColorScheme As GPTAB_XPCOLORSCHEME_METHOD)
    m_udtXPColorScheme = New_XPColorScheme
    PropertyChanged "XPColorScheme"
    Call pvSetColor
    Call pvDraw
End Property

Private Sub m_clsTabs_TabAddNew()
    If m_clsSelectTab Is Nothing Then
       Set m_clsSelectTab = m_clsTabs.Item(1)
       m_clsSelectTab.Selected = True
    End If
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabAlignmentChanged(ByVal Index As Long)
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabCaptionChanged(ByVal Index As Long)
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabIconAlignChanged(ByVal Index As Long)
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabIconChanged(ByVal Index As Long)
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabRemove()
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub m_clsTabs_TabSelectedChanged(ByVal Index As Long)
    '
End Sub

Private Sub UserControl_Click()
    If m_blnEnabled Then RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Set m_clsTabs = New cTabItems
    With m_udtDrawTextParams
         .iLeftMargin = 1
         .iRightMargin = 1
         .iTabLength = 1
         .cbSize = Len(m_udtDrawTextParams)
    End With
    m_blnAutoBackColor = True
    m_blnEnabled = True
    m_blnHotTracking = True
    m_blnMultiRow = True
    m_lngTabFixedHeight = 0
    m_lngTabFixedWidth = 0
    m_udtXPColorScheme = GpTabUseWindows
    m_udtBorderStyle = GpTabBorderStyle3D
    m_udtPlacement = GpTabPlacementTopleft
    m_udtStyle = GpTabStyleStandard
    m_oleTabBorderColor = vbWhite
    m_oleTabColor = vbButtonShadow
    m_oleTabColorActive = vbButtonFace
    m_oleTabColorHover = vbHighlight
    m_udtTabStyle = GpTabRectangle
    m_udtTabWidthStyle = GpTabJustified
    Call pvSetColor
End Sub

Private Sub UserControl_InitProperties()
'    Me.AutoBackColor = True
    Call pvCalculateSize
    Me.BackColor = UserControl.Ambient.BackColor
'    Me.BorderStyle = GpTabBorderStyle3D
    Set Me.Font = UserControl.Ambient.Font
'    Me.ForeColor = vbWindowText
'    Me.Enable = True
'    Me.MultiRow = True
'    Me.Placement = GpTabPlacementTopLeft
'    Me.Style = GpTabStyleStandard
'    Me.TabFixedHeight = 0
'    Me.TabFixedWidth = 0
'    Me.TabStyle = GpTabStandard
'    Me.TabWidthStyle = GpTabJustified
End Sub

Public Sub SelectTab(Index As Integer)
    Dim lngI As Long
    Dim clsTemp As cTabItem
    
    'On Error GoTo ErrorHandle
    
        '
        For lngI = 1 To m_clsTabs.Count
            m_clsTabs.Item(lngI).Selected = False
        Next lngI
        
        m_clsTabs.Item(Index).Selected = True
        'clsTemp = m_clsTabs.Item(Index)
        Set m_clsSelectTab = m_clsTabs.Item(Index)
        m_clsSelectTab.Selected = True
        Call pvDraw
        'Set clsTemp = Nothing
    Exit Sub
ErrorHandle:
    Select Case ShowError("SelectTab", MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort: Set clsTemp = Nothing
    End Select
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngI As Long
    Dim clsTemp As cTabItem
    Const ProcName = "UserControl_MouseDown"
    On Error GoTo ErrorHandle
    DoEvents
    If m_blnEnabled Then RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = vbLeftButton Then
       Set clsTemp = HitTest(X, Y)
       If Not (clsTemp Is Nothing) Then
          If Not clsTemp Is m_clsSelectTab Then
             For lngI = 1 To m_clsTabs.Count
             DoEvents
                 m_clsTabs.Item(lngI).Selected = False
             Next lngI
             Set m_clsSelectTab = clsTemp
             m_clsSelectTab.Selected = True
             Call pvDraw
             RaiseEvent TabClick
          End If
       End If
    End If
    Set clsTemp = Nothing
    frmMain.Code.SetFocus
   Exit Sub
ErrorHandle:
DoEvents
   'Select Case ShowError(ProcName, MODULE_NAME)
'           Case vbRetry: Resume
'           Case vbIgnore: Resume Next
'           Case vbAbort: Set clsTemp = Nothing
'    End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim clsTemp As cTabItem
    Const ProcName = "UserControl_MouseMove"
    
    On Error GoTo ErrorHandle
    If m_blnEnabled Then RaiseEvent MouseMove(Button, Shift, X, Y)
    Set clsTemp = HitTest(X, Y)
    If Not clsTemp Is m_clsHoverTab Then
       Set m_clsHoverTab = clsTemp
       If m_blnHotTracking Then Call pvDraw
    End If
    ' ��겶��
    If X >= 0 And X < UserControl.ScaleWidth And Y >= 0 And Y < UserControl.ScaleHeight And Button = 0 Then
        'If GetCapture() <> UserControl.hWnd Then Call SetCapture(UserControl.hWnd)
    Else
        'If GetCapture() = UserControl.hWnd And Button = 0 Then Call ReleaseCapture
    End If
    'If GetCapture() = UserControl.hWnd And Button = 0 Then Call ReleaseCapture
    Set clsTemp = Nothing
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort: Set clsTemp = Nothing
    End Select
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_blnEnabled Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    UserControl.AutoRedraw = True
    Call pvDraw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim clsFont As New StdFont
    
    Call pvCalculateSize
    m_blnUserMode = UserControl.Ambient.UserMode
    With PropBag
         Me.AutoBackColor = .ReadProperty("AutoBackColor", True)
         Me.BackColor = .ReadProperty("BackColor", UserControl.Ambient.BackColor)
         Me.BorderStyle = .ReadProperty("BorderStyle", GpTabBorderStyle3D)
         Me.Enable = .ReadProperty("Enable", True)
         Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
         Me.HotTracking = .ReadProperty("HotTracking", True)
         Set Me.MouseIcon = .ReadProperty("MouseIcon", Nothing)
         Me.MousePointer = .ReadProperty("MousePointer", 0)
         Me.MultiRow = .ReadProperty("MultiRow", True)
         Me.Placement = .ReadProperty("Placement", GpTabPlacementTopleft)
         Me.Style = .ReadProperty("Style", GpTabStyleStandard)
         Me.TabBorderColor = .ReadProperty("TabBorderColor", vbWhite)
         Me.TabColor = .ReadProperty("TabColor", vbButtonShadow)
         Me.TabColorActive = .ReadProperty("TabColorActive", vbButtonFace)
         Me.TabColorHover = .ReadProperty("TabColorHover", vbHighlight)
         Me.TabFixedHeight = .ReadProperty("TabFixedHeight", 0)
         Me.TabFixedWidth = .ReadProperty("TabFixedWidth", 0)
         Me.TabStyle = .ReadProperty("TabStyle", GpTabRectangle)
         Me.TabWidthStyle = .ReadProperty("TabWidthStyle", GpTabJustified)
         Me.XPColorScheme = .ReadProperty("XPColorScheme", GpTabUseWindows)
         With clsFont
              .Name = "MS Sans Serif"
              .Size = 8
         End With
         Set Me.Font = .ReadProperty("Font", clsFont)
    End With
    Set clsFont = Nothing
    Call pvSetColor
End Sub

Private Sub UserControl_Resize()
    Call pvCalculateSize
    Call pvDraw
End Sub

Public Sub Redraw()
    Call pvCalculateSize
    Call pvDraw
End Sub

Private Sub UserControl_Terminate()
    Set m_clsSelectTab = Nothing
    Set m_clsHoverTab = Nothing
    Set m_clsTabs = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim clsFont As New StdFont
    
    With PropBag
         .WriteProperty "AutoBackColor", m_blnAutoBackColor, True
         .WriteProperty "BackColor", m_oleBackColor, UserControl.Ambient.BackColor
         .WriteProperty "BorderStyle", m_udtBorderStyle, GpTabBorderStyle3D
         .WriteProperty "Enable", m_blnEnabled, True
         .WriteProperty "ForeColor", Me.ForeColor, vbWindowText
         .WriteProperty "HotTracking", m_blnHotTracking, True
         .WriteProperty "MouseIcon", Me.MouseIcon, Nothing
         .WriteProperty "MousePointer", UserControl.MousePointer, 0
         .WriteProperty "MultiRow", m_blnMultiRow, True
         .WriteProperty "Placement", m_udtPlacement, GpTabPlacementTopleft
         .WriteProperty "Style", m_udtStyle, GpTabStyleStandard
         .WriteProperty "TabBorderColor", m_oleTabBorderColor, vbWhite
         .WriteProperty "TabColor", Me.TabColor, vbButtonShadow
         .WriteProperty "TabColorActive", Me.TabColorActive, vbButtonFace
         .WriteProperty "TabColorHover", Me.TabColorHover, vbHighlight
         .WriteProperty "TabFixedHeight", m_lngTabFixedHeight, 0
         .WriteProperty "TabFixedWidth", m_lngTabFixedWidth, 0
         .WriteProperty "TabStyle", m_udtTabStyle, GpTabRectangle
         .WriteProperty "TabWidthStyle", m_udtTabWidthStyle, GpTabJustified
         .WriteProperty "XPColorScheme", m_udtXPColorScheme, GpTabUseWindows
         With clsFont
              .Name = "MS Sans Serif"
              .Size = 8
         End With
         .WriteProperty "Font", Me.Font, clsFont
    End With
    Set clsFont = Nothing
End Sub

Private Function ShowError(ByVal strFunc As String, ByVal strModule As String) As VbMsgBoxResult
    Dim lngErrNumber             As Long
    Dim strErrDescription        As String
    Dim strErrSource             As String
    
    lngErrNumber = Err.Number
    strErrDescription = Err.Description
    strErrSource = IIf(Len(strModule) > 0, _
                            "[\\" & ErrComputerName() & "] " & _
                            App.EXEName & "." & _
                            strModule & "." & _
                            strFunc & _
                            IIf(Erl <> 0, "(" & Erl & ")", ""), "") & "--" & Err.Source
    ShowError = MsgBox( _
            strErrDescription & vbCrLf & vbCrLf & _
            "Error: 0x" & Hex(lngErrNumber) & vbCrLf & vbCrLf & _
            "Call stack:" & vbCrLf & _
            strErrSource, vbCritical Or vbAbortRetryIgnore, "Error")
End Function

Private Function ErrComputerName() As String
    Static sName        As String
        
    If Len(sName) = 0 Then
        sName = String(256, 0)
        GetComputerName sName, Len(sName)
        sName = Left$(sName, InStr(sName, Chr(0)) - 1)
    End If
    ErrComputerName = sName
End Function

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture
    Dim IGuid    As Guid
    Dim NewPic   As Picture
    Dim tPicConv As PICTDESC
    
    If (hBmp = 0) Then Exit Function
   
    ' Fill PictDesc structure with necessary parts:
    With tPicConv
         .cbSizeofStruct = Len(tPicConv)
         .picType = vbPicTypeBitmap
         .hImage = hBmp
    End With
    
    ' Fill in IDispatch Interface ID
    With IGuid
         .Data1 = &H20400
         .Data4(0) = &HC0
         .Data4(7) = &H46
    End With
   
    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
   
    ' Return it:
    Set BitmapToPicture = NewPic
End Function

Public Function BrightnessColor(ByVal ColorValue As Long, ByVal Increment As Long) As Long
    Dim R, g, b As Long
    
    b = ((ColorValue \ &H10000) Mod &H100): b = b + ((b * Increment) \ &HC0)
    g = ((ColorValue \ &H100) Mod &H100) + Increment
    R = (ColorValue And &HFF) + Increment
    If R < 0 Then R = 0
    If R > 255 Then R = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
    BrightnessColor = RGB(R, g, b)
End Function

Private Sub DrawDragImage(ByRef rcNew As RECT, _
                         ByVal bFirst As Boolean, _
                         ByVal bLast As Boolean)
    Static rcCurrent     As RECT
    Dim hdc              As Long
    Dim lngReturn        As Long
    
    On Error Resume Next
    ' First get the Desktop DC:
    hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ' Set the draw mode to XOR:
    lngReturn = SetROP2(hdc, R2_NOTXORPEN)
    '// Draw over and erase the old rectangle
    If Not (bFirst) Then
       lngReturn = Rectangle(hdc, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom)
    End If
    If Not (bLast) Then
       '// Draw the new rectangle
       lngReturn = Rectangle(hdc, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom)
    End If
    ' Store this position so we can erase it next time:
    LSet rcCurrent = rcNew
    ' Free the reference to the Desktop DC we got (make sure you do this!)
    lngReturn = DeleteDC(hdc)
End Sub

Public Sub DrawImage(ByVal hIml As Long, _
                     ByVal iIndex As Long, _
                     ByVal hdc As Long, _
                     ByVal xPixels As Integer, _
                     ByVal yPixels As Integer, _
                     ByVal lIconSizeX As Long, _
                     ByVal lIconSizeY As Long, _
                     Optional ByVal bSelected = False, _
                     Optional ByVal bCut = False, _
                     Optional ByVal bDisabled = False, _
                     Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
                     Optional ByVal hExternalIml As Long = 0)
    Dim hIcon        As Long
    Dim lFlags       As Long
    Dim lhIml        As Long
    Dim lColor       As Long
    Dim iImgIndex    As Long
    Dim lngReturn    As Long
    
    ' Draw the image at 1 based index or key supplied in vKey.
    ' on the hDC at xPixels,yPixels with the supplied options.
    ' You can even draw an ImageList from another ImageList control
    ' if you supply the handle to hExternalIml with this function.
    On Error Resume Next
    iImgIndex = iIndex
    If (iImgIndex > -1) Then
       If (hExternalIml <> 0) Then
          lhIml = hExternalIml
       Else
          lhIml = hIml
       End If
       lFlags = ILD_TRANSPARENT
       If (bSelected) Or (bCut) Then
          lFlags = lFlags Or ILD_SELECTED
       End If
       If (bCut) Then
          ' Draw dithered:
          lColor = TranslateColor(oCutDitherColour)
          If (lColor = -1) Then lColor = TranslateColor(vbWindowBackground)
          lngReturn = ImageList_DrawEx(lhIml, iImgIndex, hdc, xPixels, yPixels, 0, 0, CLR_NONE, lColor, lFlags)
       ElseIf (bDisabled) Then
          ' extract a copy of the icon:
          hIcon = ImageList_GetIcon(hIml, iImgIndex, 0)
          ' Draw it disabled at x,y:
          lngReturn = DrawState(hdc, 0, 0, hIcon, 0, xPixels, yPixels, lIconSizeX, lIconSizeY, DST_ICON Or DSS_DISABLED)
          ' Clear up the icon:
          lngReturn = DestroyIcon(hIcon)
       Else
          ' Standard draw:
          lngReturn = ImageList_Draw(lhIml, iImgIndex, hdc, xPixels, yPixels, lFlags)
       End If
    End If
End Sub

Public Sub DrawLine(ByVal hdc As Long, _
                    ByVal X1 As Long, _
                    ByVal Y1 As Long, _
                    ByVal X2 As Long, _
                    ByVal Y2 As Long, _
                    ByVal Color As Long, _
                    Optional Width As Long = 1)
    Dim lngPen      As Long
    Dim lngPenOld   As Long
    Dim pt          As POINTAPI
    Const FuncName = "DrawLine"
    
    On Error GoTo ErrorHandle
    
    '/* ����һ������ */
    lngPen = CreatePen(PS_SOLID, Width, Color)
    If lngPen <> 0 Then lngPenOld = SelectObject(hdc, lngPen)
    '/* ָ��һ���µĵ�ǰ����λ��X1,Y1��ǰһ��λ�ñ�����pt�� */
    MoveToEx hdc, X1, Y1, pt
    '/* ��һ���� */
    LineTo hdc, X2, Y2
    If lngPenOld <> 0 Then SelectObject hdc, lngPenOld
    lngPenOld = 0
    If lngPen <> 0 Then DeleteObject lngPen
    Exit Sub
ErrorHandle:
    Select Case ShowError(FuncName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             If lngPenOld <> 0 Then SelectObject hdc, lngPenOld
             lngPenOld = 0
             If lngPen <> 0 Then DeleteObject lngPen
    End Select
End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_NONE
End Function

' Returns Color as long, accepts SystemColorConstants
Public Function VerifyColor(ByVal ColorVal As Long) As Long
    VerifyColor = ColorVal
    If ColorVal > &HFFFFFF Or ColorVal < 0 Then VerifyColor = GetSysColor(ColorVal And &HFFFFFF)
End Function


