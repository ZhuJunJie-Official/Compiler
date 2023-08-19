VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog cdCmd 
      Left            =   6930
      Top             =   3555
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5490
      Top             =   2970
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "本程序开源项目地址：https://github.com/JasonZhuJunJie/32-Bit-Compiler"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00959595&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   6390
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "AZ Studio 32-位 编译器"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2000-2023 AZ Studio. All rights reserved. 保留所有权利。"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   795
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3 . 0 . 0 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1035
      TabIndex        =   0
      Top             =   90
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   5280
      Left            =   0
      Top             =   0
      Width           =   8160
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   420
      Left            =   0
      Top             =   4860
      Width           =   8160
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "frmSplash.frx":2CFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
    Dim CmdParams As Variant
    Dim SourceCon As String
    Dim SourFile As String
    Dim DestFile As String

    If Command() <> "" Then
        CmdParams = Split(Command(), " ")
        If UBound(CmdParams) = 1 Then
            Timer1.Enabled = False
            If CmdParams(0) = "/c" Then
                SourFile = CmdParams(1)
                InitVirtualFiles
                If Not Dir(CmdParams(1)) <> "" Then MsgBox "File '" & CmdParams(1) & "' does not exist.": End
                Open SourFile For Binary As #1
                    SourceCon = Space(LOF(1))
                    Get #1, , SourceCon
                Close #1
                CreateVirtualFile "Entry Point", EX_ENTRY, SourceCon
            
                On Error GoTo CmdExeError
                With cdCmd
                .Filter = "可执行文件|*.exe"
                .CancelError = True
                .ShowSave
                End With
                
                If Dir(cdCmd.FileName) <> "" Then
                    If MsgBox("File already exists! Overwrite?", vbYesNo + vbCritical) = vbNo Then
                        Exit Sub
                    End If
                End If
                IsCmdCompile = True
                Compile cdCmd.FileName, False
CmdExeError:
                End
            Else
                MsgBox "Usage: /c source.* destination.*"
            End If
        End If
        Unload Me
        frmMain.Show
    End If
End Sub

Private Sub Label1_Click()
    Dim Result
    Result = ShellExecute(0, vbNullString, "https://github.com/JasonZhuJunJie/32-Bit-Compiler", vbNullString, vbNullString, SW_SHOWNORMAL)
    If Result <= 32 Then
        
    End If
End Sub

Private Sub Timer1_Timer()
    Unload Me
    frmMain.Show
End Sub
