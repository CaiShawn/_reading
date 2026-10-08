# 获取当前目录下所有文件的名称（不包括文件夹）
$files = Get-ChildItem -File | Select-Object -ExpandProperty Name

# 输出到 txt 文件
$files | Out-File -FilePath "filelist.txt" -Encoding UTF8

Write-Host "已生成 filelist.txt，共 $($files.Count) 个文件"