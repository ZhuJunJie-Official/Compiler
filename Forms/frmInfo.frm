VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "摘要"
   ClientHeight    =   4020
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   6105
   ClipControls    =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6105
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAlwaysBack 
      Caption         =   "返回.."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2430
      TabIndex        =   4
      Top             =   3375
      Visible         =   0   'False
      Width           =   1770
   End
   Begin RichTextLib.RichTextBox rtfSummary 
      Height          =   2580
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   675
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   4551
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmInfo.frx":2CFA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "运行"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4275
      TabIndex        =   0
      Top             =   3375
      Width           =   1770
   End
   Begin VB.Label lblNumErrors 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 错误.."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   3
      Top             =   180
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmInfo.frx":2D90
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "摘要:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin VB.Shape shpHead 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   645
      Left            =   45
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_transparencyKey As Long
Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = (-20)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long




Private Sub Form_Paint()
    Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long
    hBrush = CreateSolidBrush(m_transparencyKey)
    hBrushOld = SelectObject(Me.hdc, hBrush)
    GetClientRect Me.hwnd, m_Rect

    FillRect Me.hdc, m_Rect, hBrush
    SelectObject Me.hdc, hBrushOld

    DeleteObject hBrush
End Sub

Private Sub cmdAction_Click()
    Select Case cmdAction.Caption
        Case "返回..": frmMain.RunEnabled = False: Unload Me
        Case "执行.."
        Dim ProgramID As Long
        sFileToRun = """" & sFileToRun & """"
        ProgramID = Shell(sFileToRun, vbNormalFocus)
        hWndProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProgramID)
        Unload Me
    End Select
End Sub

Private Sub cmdAlwaysBack_Click()
    frmMain.RunEnabled = False
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdAction.Left = frmInfo.Width - cmdAction.Width - 200
    cmdAction.Top = frmInfo.Height - cmdAction.Height - 200
    cmdAlwaysBack.Left = 50
    cmdAlwaysBack.Top = frmInfo.Height - cmdAlwaysBack.Height - 200
    rtfSummary.Width = frmInfo.Width - 200
    rtfSummary.Height = frmInfo.Height - 1600
    shpHead.Width = frmInfo.Width - 200
End Sub

Private Sub Form_Load()
        m_transparencyKey = RGB(255, 255, 1)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributesByColor Me.hwnd, m_transparencyKey, 0, LWA_COLORKEY

   

    On Error GoTo ern

    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
    'MsgBox "1"
    DwmIsCompositionEnabled en
    If en Then
        'MsgBox "2"
        DwmExtendFrameIntoClientArea Me.hwnd, mg
        'MsgBox "OK!"
        

    End If

    Exit Sub

ern:


End Sub

