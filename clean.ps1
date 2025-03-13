# ���庯��������ļ����Ƿ���� ���ļ���.��׺��.���֡� ��ʽ
function IsValidFileName {
    param (
        [string]$fileName
    )
    $pattern = '^.*\.\w+\.\d+$'
    return $fileName -match $pattern
}

# ���庯������ȡ�ļ����Ļ������֣�ȥ���������ֲ��֣�
function GetBaseName {
    param (
        [string]$fileName
    )
    $parts = $fileName.Split('.')
    $newParts = $parts[0..($parts.Length - 2)]
    return $newParts -join '.'
}

# ���庯������ȡ�ļ����е����ֲ���
function GetNumberPart {
    param (
        [string]$fileName
    )
    $parts = $fileName.Split('.')
    return [int]$parts[-1]
}

# ���庯�������ļ��ƶ�������վ
function MoveToRecycleBin {
    param (
        [string]$filePath
    )
    $shell = New-Object -ComObject Shell.Application
    $recycleBin = $shell.Namespace(10)
    $item = $shell.Namespace((Get-Item $filePath).Directory.FullName).ParseName((Get-Item $filePath).Name)
    $recycleBin.MoveHere($item)
}

# �������߼�
$fileDict = @{}

# ��ȡ��ǰ�ļ����µ������ļ�
$files = Get-ChildItem -Path . -File

# ������ǰ�ļ����µ������ļ�
foreach ($file in $files) {
    if (IsValidFileName $file.Name) {
        $baseName = GetBaseName $file.Name
        $numberPart = GetNumberPart $file.Name
        
        # ����ֵ��в����ڸû����ļ����������
        if (-not $fileDict.ContainsKey($baseName)) {
            $fileDict[$baseName] = @($file.FullName, $numberPart)
        } else {
            # �����ǰ�ļ������ִ����ֵ��м�¼������
            if ($numberPart -gt $fileDict[$baseName][1]) {
                # ��֮ǰ��¼���ļ��ƶ�������վ
                MoveToRecycleBin $fileDict[$baseName][0]
                # �����ֵ��¼
                $fileDict[$baseName] = @($file.FullName, $numberPart)
            } else {
                # ����ǰ�ļ��ƶ�������վ
                MoveToRecycleBin $file.FullName
            }
        }
    }
}
