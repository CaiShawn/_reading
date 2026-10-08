# 用脚本所在目录作为根
$rootPath = $PSScriptRoot
Set-Location $rootPath

# 输出子目录
$outputDir = "output"
$outputFullPath = Join-Path $rootPath $outputDir
if (-not (Test-Path $outputFullPath)) {
    New-Item -ItemType Directory -Path $outputFullPath | Out-Null
}

# 时间戳
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$mdFile  = Join-Path $outputFullPath "$timestamp-readlist.md"
$csvFile = Join-Path $outputFullPath "$timestamp-readlist.csv"

# 排除规则
$excludePattern = '\.(csv|txt|ps1|md)$|^\.'

# 收集文件
$files = Get-ChildItem -Path $rootPath -File -Recurse |
    Where-Object {
        $_.Name -notmatch $excludePattern -and
        $_.DirectoryName -ne $outputFullPath -and
        $_.DirectoryName -notlike "$outputFullPath\*"
    } |
    Select-Object `
        @{Name="文件名";   Expression={$_.Name}}, `
        @{Name="大小(MB)"; Expression={[math]::Round($_.Length / 1MB, 2)}}, `
        @{Name="类型";     Expression={$_.Extension}}, `
        @{Name="所在目录"; Expression={
            if ($_.DirectoryName.Length -ge $rootPath.Length) {
                $rel = $_.DirectoryName.Substring($rootPath.Length).TrimStart('\')
            } else {
                $rel = $_.DirectoryName
            }
            if ($rel -eq '') { '.' } else { $rel }
        }}

# ---- 生成 Markdown 表格 ----
# 转义函数：处理 markdown 表格里会冲突的字符
function ConvertTo-MdCell {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    # | 要转义为 \|，换行替换为空格
    return ($Text -replace '\|', '\|') -replace "`r?`n", ' '
}

$md = [System.Text.StringBuilder]::new()
[void]$md.AppendLine("# 文件清单")
[void]$md.AppendLine("")
[void]$md.AppendLine("- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$md.AppendLine("- 根目录：``$rootPath``")
[void]$md.AppendLine("- 文件总数：$($files.Count)")
[void]$md.AppendLine("")
[void]$md.AppendLine("| 文件名 | 大小(MB) | 类型 | 所在目录 |")
[void]$md.AppendLine("|--------|---------:|------|----------|")

foreach ($f in $files) {
    $name = ConvertTo-MdCell $f.文件名
    $size = $f.'大小(MB)'
    $type = ConvertTo-MdCell $f.类型
    $dir  = ConvertTo-MdCell $f.所在目录
    [void]$md.AppendLine("| $name | $size | $type | $dir |")
}

# ---- 输出 ----
try {
    $md.ToString() | Out-File -FilePath $mdFile -Encoding UTF8 -ErrorAction Stop
    $files | Export-Csv -Path $csvFile -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
} catch {
    Write-Warning "写入失败: $_"
}

Write-Host "Done. Total $($files.Count) files."
Write-Host "Output: $mdFile"
Write-Host "Output: $csvFile"