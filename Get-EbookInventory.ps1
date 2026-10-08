# 用脚本所在目录作为根
$rootPath = $PSScriptRoot
Set-Location $rootPath

# 输出子目录
$outputDir      = "output"
$outputFullPath = Join-Path $rootPath $outputDir
if (-not (Test-Path $outputFullPath)) {
    New-Item -ItemType Directory -Path $outputFullPath | Out-Null
}

# 时间戳
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$mdFile    = Join-Path $outputFullPath "$timestamp-readlist.md"
$csvFile   = Join-Path $outputFullPath "$timestamp-readlist.csv"
$xlsxFile  = Join-Path $outputFullPath "$timestamp-readlist.xlsx"

# 排除规则（含 xlsx，避免把自己扫进去）
$excludePattern = '\.(csv|txt|ps1|md|xlsx)$|^\.'

# ================= 书名/作者解析 =================
function Parse-BookFile {
    param([string]$BaseName)

    # z-lib 尾巴（允许前后有空格）
    $zlibRx = [regex]'\s*\(z-library\.sk,\s*1lib\.sk,\s*z-lib\.sk\)\s*$'
    $hasZlib = $zlibRx.IsMatch($BaseName)
    $stripped = if ($hasZlib) { $zlibRx.Replace($BaseName, '') } else { $BaseName }
    $stripped = $stripped.TrimEnd()

    # 从末尾往回找"最外层"一对括号：
    #   - 末尾必须是 ')'
    #   - 从末尾往前找到第一个与之配对的 '('（考虑括号嵌套深度）
    $author = $null
    $title  = $null
    if ($stripped.EndsWith(')')) {
        $depth = 0
        $openIdx = -1
        for ($i = $stripped.Length - 1; $i -ge 0; $i--) {
            $ch = $stripped[$i]
            if ($ch -eq ')') { $depth++ }
            elseif ($ch -eq '(') {
                $depth--
                if ($depth -eq 0) { $openIdx = $i; break }
            }
        }
        if ($openIdx -gt 0) {
            $beforeOpen = $stripped.Substring(0, $openIdx).TrimEnd()
            # 只有当 '(' 前面是空格或行首，才认为它是"作者括号"
            if ($beforeOpen.Length -eq 0 -or $stripped[$openIdx - 1] -eq ' ') {
                $author = $stripped.Substring($openIdx + 1, $stripped.Length - $openIdx - 2).Trim()
                $title  = $beforeOpen.Trim()
            }
        }
    }

    if ($title -and $author) {
        $note = @()
        if (-not $hasZlib)             { $note += '缺少 z-library 标记' }
        if ($author -match '^etc\.?$') { $note += '作者为 etc.' }
        return [pscustomobject]@{
            Title      = $title
            Author     = $author
            IsStandard = ($hasZlib -and ($author -notmatch '^etc\.?$'))
            Note       = ($note -join '; ')
        }
    }

    # 兜底：实在解析不出来，保留原始名，标记异常
    $note = @()
    if (-not $hasZlib) { $note += '缺少 z-library 标记' }
    $note += '不符合标准格式'
    return [pscustomobject]@{
        Title      = $BaseName
        Author     = ''
        IsStandard = $false
        Note       = ($note -join '; ')
    }
}

# ================= 收集并解析 =================
$rawFiles = Get-ChildItem -Path $rootPath -File -Recurse |
    Where-Object {
        $_.Name -notmatch $excludePattern -and
        $_.DirectoryName -ne $outputFullPath -and
        $_.DirectoryName -notlike "$outputFullPath\*"
    }

$fileRows = @(
    foreach ($f in $rawFiles) {
        $base  = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $parse = Parse-BookFile -BaseName $base
        $dirRel = if ($f.DirectoryName.Length -ge $rootPath.Length) {
            $r = $f.DirectoryName.Substring($rootPath.Length).TrimStart('\')
            if ([string]::IsNullOrEmpty($r)) { '.' } else { $r }
        } else { $f.DirectoryName }

        [pscustomobject]@{
            文件名     = $f.Name
            书名       = $parse.Title
            作者       = $parse.Author
            格式       = if ($f.Extension) { $f.Extension.TrimStart('.').ToUpper() } else { '无扩展名' }
            '大小(MB)' = [math]::Round($f.Length / 1MB, 2)
            所在目录   = $dirRel
            标准       = if ($parse.IsStandard) { '是' } else { '否' }
            备注       = $parse.Note
        }
    }
)

$fileRows = @($fileRows | Sort-Object 所在目录, 书名, 作者, 格式)

# ================= Markdown（文件级，含解析结果）=================
function ConvertTo-MdCell {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return ($Text -replace '\|', '\|') -replace "`r?`n", ' '
}

$md = [System.Text.StringBuilder]::new()
[void]$md.AppendLine("# 文件清单")
[void]$md.AppendLine("")
[void]$md.AppendLine("- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$md.AppendLine("- 根目录：``$rootPath``")
[void]$md.AppendLine("- 文件总数：$($fileRows.Count)")
[void]$md.AppendLine("- 不符合标准格式：$(($fileRows | Where-Object { $_.标准 -eq '否' }).Count)")
[void]$md.AppendLine("")
[void]$md.AppendLine("| 书名 | 作者 | 格式 | 大小(MB) | 所在目录 | 标准 | 备注 |")
[void]$md.AppendLine("|------|------|------|---------:|----------|:----:|------|")
foreach ($f in $fileRows) {
    $t = ConvertTo-MdCell $f.书名
    $a = ConvertTo-MdCell $f.作者
    $e = ConvertTo-MdCell $f.格式
    $s = $f.'大小(MB)'
    $d = ConvertTo-MdCell $f.所在目录
    $k = $f.标准
    $n = ConvertTo-MdCell $f.备注
    [void]$md.AppendLine("| $t | $a | $e | $s | $d | $k | $n |")
}

try {
    $md.ToString() | Out-File -FilePath $mdFile -Encoding UTF8 -ErrorAction Stop
    $fileRows | Export-Csv -Path $csvFile -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
} catch {
    Write-Warning "写入 Markdown/CSV 失败: $_"
}

# ================= Sheet 名称清洗与去重 =================
function Get-ValidSheetName {
    param([string]$RawName, [hashtable]$Used)
    $name = if ([string]::IsNullOrWhiteSpace($RawName) -or $RawName -eq '.') { 'Root' } else { $RawName }
    $name = $name -replace '[\\/:*?\[\]]', '_'
    if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }

    $base = $name
    $i = 1
    while ($Used.ContainsKey($name)) {
        $suffix = "_$i"
        $keep = 31 - $suffix.Length
        $name = if ($base.Length -gt $keep) { $base.Substring(0, $keep) + $suffix } else { $base + $suffix }
        $i++
    }
    $Used[$name] = $true
    return $name
}

# ================= 构建 Excel 数据 =================
$allExts   = @($fileRows | Select-Object -ExpandProperty 格式 -Unique | Sort-Object)
$dirGroups = @($fileRows | Group-Object 所在目录 | Sort-Object Name)

$usedNames   = @{ 'Summary' = $true }
$sheetDefs   = @()
$summaryRows = @()

# 汇总表：总计行
$summaryRows += [pscustomobject]@{
    "目录"       = "(全部)"
    "书目数"     = @($fileRows | Group-Object 所在目录, 书名, 作者).Count
    "文件数"     = $fileRows.Count
    "总大小(MB)" = [math]::Round((@($fileRows | Measure-Object -Property "大小(MB)" -Sum).Sum), 2)
    "不符合标准" = @($fileRows | Where-Object { $_.标准 -eq '否' }).Count
}

foreach ($dg in $dirGroups) {
    $sheetName  = Get-ValidSheetName -RawName $dg.Name -Used $usedNames
    $bookGroups = @($dg.Group | Group-Object 书名, 作者 | Sort-Object Name)

    $rows = @(
        foreach ($bg in $bookGroups) {
            $first = $bg.Group[0]
            $row = [ordered]@{
                "书名" = $first.书名
                "作者" = $first.作者
            }
            foreach ($ext in $allExts) {
                $hit = $bg.Group | Where-Object { $_.'格式' -eq $ext }
                $row[$ext] = if ($hit) { "✔" } else { "" }
            }
            $row["大小(MB)"] = [math]::Round((@($bg.Group | Measure-Object -Property "大小(MB)" -Sum).Sum), 2)
            $row["文件数"]   = $bg.Group.Count
            $notes = @($bg.Group | Where-Object { $_.备注 } | Select-Object -ExpandProperty 备注 -Unique)
            $row["备注"]     = ($notes -join '; ')
            [pscustomobject]$row
        }
    )

    $sheetDefs += [pscustomobject]@{
        Name = $sheetName
        Rows = @($rows)
    }

    $summaryRows += [pscustomobject]@{
        "目录"       = $dg.Name
        "书目数"     = $bookGroups.Count
        "文件数"     = $dg.Group.Count
        "总大小(MB)" = [math]::Round((@($dg.Group | Measure-Object -Property "大小(MB)" -Sum).Sum), 2)
        "不符合标准" = @($dg.Group | Where-Object { $_.标准 -eq '否' }).Count
    }
}

# Summary 放最前
$sheetDefs = @(
    [pscustomobject]@{ Name = 'Summary'; Rows = @($summaryRows) }
) + $sheetDefs

# ================= 写入 Excel =================
Import-Module ImportExcel -Force -ErrorAction Stop

if (Test-Path $xlsxFile) { Remove-Item $xlsxFile -Force }

$isFirst = $true
foreach ($s in $sheetDefs) {
    if (-not $s.Rows -or $s.Rows.Count -eq 0) { continue }

    $params = @{
        Path          = $xlsxFile
        WorksheetName = $s.Name
        AutoSize      = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        InputObject   = $s.Rows
        ErrorAction   = 'Stop'
    }
    if (-not $isFirst) { $params['Append'] = $true }

    try {
        Export-Excel @params
        $isFirst = $false
    } catch {
        Write-Warning "写入 Sheet [$($s.Name)] 失败: $_"
    }
}

# ================= 验证 =================
$sheetNames = (Get-ExcelSheetInfo -Path $xlsxFile).Name
Write-Host ""
Write-Host "Output: $xlsxFile"
Write-Host "Sheets ($($sheetNames.Count)): $($sheetNames -join ', ')"
Write-Host ""

$badCount = @($fileRows | Where-Object { $_.标准 -eq '否' }).Count
$etcCount = @($fileRows | Where-Object { $_.备注 -like '*作者为 etc.*' }).Count

Write-Host "Done. Total $($fileRows.Count) files."
if ($badCount -gt 0) { Write-Host "  - 不符合标准格式: $badCount 个（详见备注列）" }
if ($etcCount -gt 0) { Write-Host "  - 作者为 etc. 占位: $etcCount 个" }
Write-Host "Output: $mdFile"
Write-Host "Output: $csvFile"