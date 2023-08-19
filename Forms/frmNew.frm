VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "新建工程"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "frmNew.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6615
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   90
      TabIndex        =   3
      Top             =   2655
      Width           =   6450
   End
   Begin Zhujunjie_官方.McToolBar McToolBar1 
      Height          =   2505
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4419
      BackColor       =   16777215
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Button_Count    =   4
      ButtonsWidth    =   90
      ButtonsHeight   =   160
      ButtonsPerRow   =   4
      HoverColor      =   10520954
      TooTipStyle     =   0
      BackGradient    =   5
      BackGradientCol =   12549429
      ButtonsStyle    =   2
      BorderColor     =   16777215
      ButtonCaption0  =   "Windows GUI"
      ButtonIcon0     =   "frmNew.frx":038A
      ButtonCaption1  =   "动态链接库文件"
      ButtonIcon1     =   "frmNew.frx":1064
      ButtonCaption2  =   "Windows 控制台"
      ButtonIcon2     =   "frmNew.frx":1D3E
      ButtonCaption3  =   "空白工程"
      ButtonIcon3     =   "frmNew.frx":2A18
   End
   Begin VB.CommandButton cmdExist 
      Caption         =   "打开现有的 .."
      Height          =   510
      Left            =   90
      TabIndex        =   1
      Top             =   2790
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   510
      Left            =   4980
      TabIndex        =   0
      Top             =   2790
      Width           =   1545
   End
End
Attribute VB_Name = "frmNew"
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
Private Sub cmdCancel_Click()
    frmMain.MakeInitControls
    Unload Me
End Sub

Private Sub cmdExist_Click()
    frmMain.MakeInitControls
    frmMain.OpenProject
    Unload Me
End Sub

Sub TemplateGUI()
    frmMain.MakeInitControls
    SetVirtualFileContent "Entry Point", _
                "application PE GUI;" & vbCrLf & vbCrLf & _
                "import MessageBox ascii lib ""USER32.DLL"",4;" & vbCrLf & vbCrLf & _
                "entry" & vbCrLf & vbCrLf & _
                vbTab & "MessageBox(0,""Hello World!"",""GUI"",$20);" & vbCrLf & _
                vbCrLf & _
                "end."
    frmMain.SelectEntryFile
End Sub

Sub TemplateDLL()
    frmMain.MakeInitControls
    SetVirtualFileContent "Entry Point", _
                "application PE GUI DLL;" & vbCrLf & vbCrLf & _
                "export IsInitialized();" & vbCrLf & _
                vbTab & "return(TRUE);" & vbCrLf & _
                "end;" & vbCrLf
    frmMain.SelectEntryFile
End Sub

Sub TemplateCUI()
    frmMain.MakeInitControls
    SetVirtualFileContent "Entry Point", _
                "application PE CUI;" & vbCrLf & vbCrLf & _
                "include ""Windows.inc"", ""Console.inc"";" & vbCrLf & vbCrLf & _
                "entry" & vbCrLf & vbCrLf & _
                vbTab & "Console.Init(""Console"");" & vbCrLf & _
                vbTab & "Console.Write(""Hello World!"");" & vbCrLf & _
                vbTab & "Console.Read();" & vbCrLf & _
                vbCrLf & _
                "end."
    frmMain.SelectEntryFile
End Sub


Private Sub Form_Load()
    frmMain.ucTab.Tabs.Clear
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

Private Sub McToolBar1_Click(ByVal vButton_Index As Long)
    Select Case vButton_Index
        Case 0: TemplateGUI
        Case 1: TemplateDLL
        Case 2: TemplateCUI
        Case 3: frmMain.MakeInitControls: frmMain.SelectEntryFile
        Case Else: Exit Sub
    End Select
    Unload Me
End Sub












