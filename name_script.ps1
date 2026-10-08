# 获取当前目录（根目录）下所有子目录内的文件
$files = Get-ChildItem -Path . -File -Recurse | Select-Object `
    @{Name="文件名"; Expression={$_.Name}}, `
    @{Name="大小(KB)"; Expression={[math]::Round($_.Length / 1KB, 2)}}, `
    @{Name="类型"; Expression={$_.Extension}}, `
    @{Name="所在目录"; Expression={$_.DirectoryName}}

# 输出到 txt 文件（格式化为表格）
$files | Format-Table -AutoSize | Out-File -FilePath "filelist.txt" -Encoding UTF8

# 同时输出为 CSV（便于后续处理）
$files | Export-Csv -Path "filelist.csv" -Encoding UTF8 -NoTypeInformation

Write-Host "已生成 filelist.txt 和 filelist.csv，共 $($files.Count) 个文件"