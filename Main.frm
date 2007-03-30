VERSION 5.00
Begin VB.Form Main 
   Caption         =   "FoolCraft"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblDir 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   7935
   End
   Begin VB.TextBox txtMask 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Text            =   "*.*"
      Top             =   360
      Width           =   495
   End
   Begin VB.FileListBox File 
      BackColor       =   &H00C0FFC0&
      Height          =   2625
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.DriveListBox Drive 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.DirListBox Dir 
      BackColor       =   &H00C0FFC0&
      Height          =   2565
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Маска"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Menu mFile 
      Caption         =   "Файл"
      Begin VB.Menu Search 
         Caption         =   "Найти обьект"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu Open 
         Caption         =   "Открыть"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Правка"
      Begin VB.Menu DeleteFile 
         Caption         =   "Удалить файл"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu DeletePath 
         Caption         =   "Удалить папку"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu SysFiles 
         Caption         =   "Скрывать системные файлы"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Справка"
      Begin VB.Menu About 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFile As New FileSystemObject
Dim objShell As New Shell32.Shell
Dim strDirec, _
    strDrive As String

Private Sub About_Click()
    MsgBox "VitekSoft, 2007. All rights are resived.", vbInformation, "About"
End Sub

Private Sub DeletePath_Click()
    If MsgBox("Вы точно хотите удалить папку: " & Dir.List(Dir.ListIndex), vbYesNo, "Предупреждение") <> vbYes Then Exit Sub
    objFile.DeleteFolder Dir.List(Dir.ListIndex)
    Dir.Refresh
    File.Refresh
    If Len(Dir.List(Dir.ListIndex)) = 3 Then
        lblDir = Dir.List(Dir.ListIndex) & File.FileName
    Else
        lblDir = Dir.List(Dir.ListIndex) & "\" & File.FileName
    End If
End Sub

Private Sub DeleteFile_Click()
    If File.ListIndex <> -1 Then
        If MsgBox("Вы точно хотите удалить файл: " & strDirec & "\" & File.FileName, vbYesNo, "Предупреждение") <> vbYes Then Exit Sub
        Kill strDirec & "\" & File.FileName
    End If
    Dir.Refresh
    File.Refresh
    If Len(Dir.List(Dir.ListIndex)) = 3 Then
        lblDir = Dir.List(Dir.ListIndex) & File.FileName
    Else
        lblDir = Dir.List(Dir.ListIndex) & "\" & File.FileName
    End If
End Sub

Private Sub Dir_Change()
    
    File.Path = strDirec
    File.Refresh
    If Len(Dir.List(Dir.ListIndex)) = 3 Then
        lblDir = Dir.List(Dir.ListIndex) & File.FileName
    Else
        lblDir = Dir.List(Dir.ListIndex) & "\" & File.FileName
    End If
End Sub

Private Sub Dir_Click()
    File.ListIndex = -1
    strDirec = Dir.List(Dir.ListIndex)
    Dir_Change
End Sub

Private Sub Drive_Change()
    On Error GoTo Err
    Dir.Path = Drive.List(Drive.ListIndex)
    strDrive = Drive.List(Drive.ListIndex)
    Dir.Refresh
    File.Refresh
    If Len(Dir.List(Dir.ListIndex)) = 3 Then
        lblDir = Dir.List(Dir.ListIndex) & File.FileName
    Else
        lblDir = Dir.List(Dir.ListIndex) & "\" & File.FileName
    End If
    Exit Sub
Err:
    MsgBox Err.Number & ":  " & Err.Description, vbOKOnly, "FoolCraft"
End Sub

Private Sub File_Click()
    If Len(Dir.List(Dir.ListIndex)) = 3 Then
        lblDir = Dir.List(Dir.ListIndex) & File.FileName
    Else
        lblDir = Dir.List(Dir.ListIndex) & "\" & File.FileName
    End If
End Sub

Private Sub Form_Load()
    File.Pattern = txtMask
    File.System = SysFiles.Checked
    On Error Resume Next
    Drive.ListIndex = GetSetting("FoolCraft", "Default", "drive")
    Dir.Path = GetSetting("FoolCraft", "Default", "Directory")
    File.Path = GetSetting("FoolCraft", "Default", "Directory")
End Sub

Private Sub Form_Resize()
    Drive.Top = 200
    On Error GoTo Ext
    Dir.Height = Height - Drive.Top - Drive.Height - 1000
    Drive.Left = 200
    Drive.Width = (Width - 500) / 2
    Dir.Top = Drive.Top + Drive.Height
    Dir.Left = Drive.Left
    Dir.Width = Drive.Width
    File.Width = Dir.Width
    File.Height = Dir.Height
    File.Left = Dir.Left + Dir.Width
    File.Top = Dir.Top
    Label1.Top = Drive.Top
    Label1.Left = Drive.Left + Drive.Width + 200
    txtMask.Top = Drive.Top
    txtMask.Left = Label1.Left + Label1.Width
    lblDir.Width = Width
    lblDir.Top = Height - 1100
Ext:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "FoolCraft", "Default", "Directory", strDirec
    SaveSetting "FoolCraft", "Default", "Drive", Drive.ListIndex
End Sub

Private Sub Open_Click()
    If File.ListIndex <> -1 Then
        objShell.FileRun
    End If
End Sub

Private Sub Search_Click()
    objShell.Explore Dir.List(Dir.ListIndex)
End Sub

Private Sub SysFiles_Click()
    SysFiles.Checked = Not (SysFiles.Checked)
    File.System = SysFiles.Checked
End Sub

Private Sub txtMask_Change()
    File.Pattern = txtMask
End Sub
