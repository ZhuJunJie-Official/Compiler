VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6345
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AZ Studio 32-位 编译器 3.0.0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   1575
      TabIndex        =   1
      Top             =   0
      Width           =   4740
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "确定"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3120
         TabIndex        =   6
         Tag             =   "OK"
         Top             =   2760
         Width           =   1230
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Height          =   135
         Left            =   135
         TabIndex        =   5
         Top             =   2430
         Width           =   4455
      End
      Begin VB.TextBox txtCredits 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   945
         Width           =   4470
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   "查看本程序开源项目地址"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   360
         TabIndex        =   8
         Tag             =   "Warning"
         Top             =   2880
         Width           =   2385
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "x.xx"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4005
         TabIndex        =   4
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "版本信息:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "32-位程序开发工具"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   7
         Tag             =   "Anwendungsbeschreibung"
         Top             =   585
         Width           =   3735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3555
      Left            =   0
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   3495
      ScaleWidth      =   1470
      TabIndex        =   0
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "frmAbout"
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
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "3.0.0"
    txtCredits = "Thanks:" & vbNewLine & _
    "AZ Studio"
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

Private Sub lblDisclaimer_Click()
    Dim Result
    Result = ShellExecute(0, vbNullString, "https://github.com/JasonZhuJunJie/32-Bit-Compiler", vbNullString, vbNullString, SW_SHOWNORMAL)
    If Result <= 32 Then
        
    End If
End Sub


