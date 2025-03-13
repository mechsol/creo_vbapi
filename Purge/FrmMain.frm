VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creo旧版本清除工具"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3735
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1395
   ScaleWidth      =   3735
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 在窗体代码顶部添加 API 声明
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
' 引用 Shell 类型库，用于将文件移动到回收站
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
' 定义 SHFILEOPSTRUCT 结构体
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
' 定义常量
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10 ' 禁止确认对话框

' 递归删除旧文件
Sub DeleteOldFiles(ByVal FolderPath As String, Optional ByVal Recurse As Boolean = True)
    Dim FileName As String
    Dim FileList As New Collection
    Dim FileParts() As String
    Dim FileBase As String
    Dim FileNumber As Long
    Dim MaxNumber As Long
    Dim MaxFileName As String
    Dim Item As Variant
    Dim SubFolderPath As String
    
    ' 确保文件夹路径以反斜杠结尾
    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If
    
    ' 获取文件夹下的所有文件
    FileName = Dir(FolderPath & "*.*")
    Do While FileName <> ""
        ' 检查文件名是否符合 文件名.后缀名.数字 格式
        FileParts = Split(FileName, ".")
        If UBound(FileParts) >= 2 Then
            If IsNumeric(FileParts(UBound(FileParts))) Then
                ' 提取文件名和数字
                FileBase = Left(FileName, Len(FileName) - Len(FileParts(UBound(FileParts))) - 1)
                FileNumber = CLng(FileParts(UBound(FileParts)))
                
                ' 检查是否已经有相同文件名的文件
                Dim Found As Boolean
                Found = False
                For Each Item In FileList
                    If Left(Item, Len(FileBase)) = FileBase Then
                        Dim ItemParts() As String
                        ItemParts = Split(Item, ".")
                        Dim ItemNumber As Long
                        ItemNumber = CLng(ItemParts(UBound(ItemParts)))
                        If FileNumber > ItemNumber Then
                            ' 删除数字较小的文件
                            MoveToRecycleBin FolderPath & Item
                        Else
                            ' 删除当前文件
                            MoveToRecycleBin FolderPath & FileName
                            Exit For ' 无需继续检查，已找到并删除
                        End If
                        Found = True
                    End If
                Next Item
                
                If Not Found Then
                    ' 如果没有相同文件名的文件，添加到集合中
                    FileList.Add FileName
                End If
            End If
        End If
        FileName = Dir
    Loop
    
    ' 如果需要递归处理子文件夹
    If Recurse Then
        SubFolderPath = Dir(FolderPath & "*", vbDirectory)
        Do While SubFolderPath <> ""
            ' 忽略 "." 和 ".." 目录
            If SubFolderPath <> "." And SubFolderPath <> ".." Then
                ' 如果是文件夹，则递归调用
                If (GetAttr(FolderPath & SubFolderPath) And vbDirectory) <> 0 Then
                    DeleteOldFiles FolderPath & SubFolderPath, True
                End If
            End If
            SubFolderPath = Dir
        Loop
    End If
End Sub

Sub MoveToRecycleBin(ByVal FilePath As String)
    Dim FileOp As SHFILEOPSTRUCT
    FileOp.hWnd = 0
    FileOp.wFunc = FO_DELETE
    FileOp.pFrom = FilePath & vbNullChar
    FileOp.pTo = vbNullString
    ' 组合标志，允许移动到回收站且不弹出确认对话框
    FileOp.fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
    FileOp.fAnyOperationsAborted = False
    FileOp.hNameMappings = 0
    FileOp.lpszProgressTitle = vbNullString
    SHFileOperation FileOp
End Sub

Private Sub Form_Load()
    Me.OLEDropMode = vbOLEDropManual
End Sub

Private Sub Form_Paint()
    Dim text As String
    text = "拖动要处理的文件夹到此处"  ' 要居中的文字
    ' 计算居中坐标（ScaleWidth/Height自动适配单位）
    Me.CurrentX = (Me.ScaleWidth - Me.TextWidth(text)) / 2
    Me.CurrentY = (Me.ScaleHeight - Me.TextHeight(text)) / 2
    Me.Print text  ' 绘制文字
End Sub

Private Sub Form_Resize()
    Me.Refresh  ' 窗体大小变化时强制重绘
End Sub

' 拖放事件处理
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next ' 防止无效路径导致崩溃
    Dim vFile As Variant
    Dim sPath As String
    
    ' 遍历所有拖放的路径
    For Each vFile In Data.Files
        sPath = vFile
        ' 判断类型
        If PathIsDirectory(sPath) Then
            DeleteOldFiles sPath, True ' 递归调用
        End If
    Next vFile
    MsgBox "操作已完成。", vbExclamation
End Sub
