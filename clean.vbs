Option Explicit

Dim fso, folder, files, file, fileInfo, baseName, numberPart
Dim fileDict
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(".")
Set files = folder.Files
Set fileDict = CreateObject("Scripting.Dictionary")

' 遍历当前文件夹下的所有文件
For Each file In files
    If IsValidFileName(file.Name) Then
        baseName = GetBaseName(file.Name)
        numberPart = GetNumberPart(file.Name)
        
        ' 如果字典中不存在该基础文件名，则添加
        If Not fileDict.Exists(baseName) Then
            fileDict.Add baseName, Array(file.Path, numberPart)
        Else
            ' 如果当前文件的数字大于字典中记录的数字
            If numberPart > CInt(fileDict(baseName)(1)) Then
                ' 将之前记录的文件移动到回收站
                MoveToRecycleBin fileDict(baseName)(0)
                ' 更新字典记录
                fileDict(baseName) = Array(file.Path, numberPart)
            Else
                ' 将当前文件移动到回收站
                MoveToRecycleBin file.Path
            End If
        End If
    End If
Next

' 检查文件名是否符合 “文件名.后缀名.数字” 格式
Function IsValidFileName(fileName)
    Dim pattern
    pattern = "^.*\.\w+\.\d+$"
    Dim re
    Set re = New RegExp
    re.Pattern = pattern
    IsValidFileName = re.Test(fileName)
End Function

' 获取文件名的基础部分（去掉最后的数字部分）
Function GetBaseName(fileName)
    Dim parts
    parts = Split(fileName, ".")
    ReDim Preserve parts(UBound(parts) - 1)
    GetBaseName = Join(parts, ".")
End Function

' 获取文件名中的数字部分
Function GetNumberPart(fileName)
    Dim parts
    parts = Split(fileName, ".")
    GetNumberPart = CInt(parts(UBound(parts)))
End Function

' 将文件移动到回收站
Sub MoveToRecycleBin(filePath)
    Dim shell
    Set shell = CreateObject("Shell.Application")
    shell.Namespace(10).MoveHere filePath
End Sub