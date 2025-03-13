# 定义函数，检查文件名是否符合 “文件名.后缀名.数字” 格式
function IsValidFileName {
    param (
        [string]$fileName
    )
    $pattern = '^.*\.\w+\.\d+$'
    return $fileName -match $pattern
}

# 定义函数，获取文件名的基础部分（去掉最后的数字部分）
function GetBaseName {
    param (
        [string]$fileName
    )
    $parts = $fileName.Split('.')
    $newParts = $parts[0..($parts.Length - 2)]
    return $newParts -join '.'
}

# 定义函数，获取文件名中的数字部分
function GetNumberPart {
    param (
        [string]$fileName
    )
    $parts = $fileName.Split('.')
    return [int]$parts[-1]
}

# 定义函数，将文件移动到回收站
function MoveToRecycleBin {
    param (
        [string]$filePath
    )
    $shell = New-Object -ComObject Shell.Application
    $recycleBin = $shell.Namespace(10)
    $item = $shell.Namespace((Get-Item $filePath).Directory.FullName).ParseName((Get-Item $filePath).Name)
    $recycleBin.MoveHere($item)
}

# 主程序逻辑
$fileDict = @{}

# 获取当前文件夹下的所有文件
$files = Get-ChildItem -Path . -File

# 遍历当前文件夹下的所有文件
foreach ($file in $files) {
    if (IsValidFileName $file.Name) {
        $baseName = GetBaseName $file.Name
        $numberPart = GetNumberPart $file.Name
        
        # 如果字典中不存在该基础文件名，则添加
        if (-not $fileDict.ContainsKey($baseName)) {
            $fileDict[$baseName] = @($file.FullName, $numberPart)
        } else {
            # 如果当前文件的数字大于字典中记录的数字
            if ($numberPart -gt $fileDict[$baseName][1]) {
                # 将之前记录的文件移动到回收站
                MoveToRecycleBin $fileDict[$baseName][0]
                # 更新字典记录
                $fileDict[$baseName] = @($file.FullName, $numberPart)
            } else {
                # 将当前文件移动到回收站
                MoveToRecycleBin $file.FullName
            }
        }
    }
}
