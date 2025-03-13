VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creo�ɰ汾�������"
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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' �ڴ�����붥����� API ����
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
' ���� Shell ���Ϳ⣬���ڽ��ļ��ƶ�������վ
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
' ���� SHFILEOPSTRUCT �ṹ��
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
' ���峣��
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10 ' ��ֹȷ�϶Ի���

' �ݹ�ɾ�����ļ�
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
    
    ' ȷ���ļ���·���Է�б�ܽ�β
    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If
    
    ' ��ȡ�ļ����µ������ļ�
    FileName = Dir(FolderPath & "*.*")
    Do While FileName <> ""
        ' ����ļ����Ƿ���� �ļ���.��׺��.���� ��ʽ
        FileParts = Split(FileName, ".")
        If UBound(FileParts) >= 2 Then
            If IsNumeric(FileParts(UBound(FileParts))) Then
                ' ��ȡ�ļ���������
                FileBase = Left(FileName, Len(FileName) - Len(FileParts(UBound(FileParts))) - 1)
                FileNumber = CLng(FileParts(UBound(FileParts)))
                
                ' ����Ƿ��Ѿ�����ͬ�ļ������ļ�
                Dim Found As Boolean
                Found = False
                For Each Item In FileList
                    If Left(Item, Len(FileBase)) = FileBase Then
                        Dim ItemParts() As String
                        ItemParts = Split(Item, ".")
                        Dim ItemNumber As Long
                        ItemNumber = CLng(ItemParts(UBound(ItemParts)))
                        If FileNumber > ItemNumber Then
                            ' ɾ�����ֽ�С���ļ�
                            MoveToRecycleBin FolderPath & Item
                        Else
                            ' ɾ����ǰ�ļ�
                            MoveToRecycleBin FolderPath & FileName
                            Exit For ' ���������飬���ҵ���ɾ��
                        End If
                        Found = True
                    End If
                Next Item
                
                If Not Found Then
                    ' ���û����ͬ�ļ������ļ�����ӵ�������
                    FileList.Add FileName
                End If
            End If
        End If
        FileName = Dir
    Loop
    
    ' �����Ҫ�ݹ鴦�����ļ���
    If Recurse Then
        SubFolderPath = Dir(FolderPath & "*", vbDirectory)
        Do While SubFolderPath <> ""
            ' ���� "." �� ".." Ŀ¼
            If SubFolderPath <> "." And SubFolderPath <> ".." Then
                ' ������ļ��У���ݹ����
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
    ' ��ϱ�־�������ƶ�������վ�Ҳ�����ȷ�϶Ի���
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
    text = "�϶�Ҫ������ļ��е��˴�"  ' Ҫ���е�����
    ' ����������꣨ScaleWidth/Height�Զ����䵥λ��
    Me.CurrentX = (Me.ScaleWidth - Me.TextWidth(text)) / 2
    Me.CurrentY = (Me.ScaleHeight - Me.TextHeight(text)) / 2
    Me.Print text  ' ��������
End Sub

Private Sub Form_Resize()
    Me.Refresh  ' �����С�仯ʱǿ���ػ�
End Sub

' �Ϸ��¼�����
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next ' ��ֹ��Ч·�����±���
    Dim vFile As Variant
    Dim sPath As String
    
    ' ���������Ϸŵ�·��
    For Each vFile In Data.Files
        sPath = vFile
        ' �ж�����
        If PathIsDirectory(sPath) Then
            DeleteOldFiles sPath, True ' �ݹ����
        End If
    Next vFile
    MsgBox "��������ɡ�", vbExclamation
End Sub
