param (
    [string]$InputFile,
    [string]$OutputFile
)

# Resolve input path
$InputPath = (Resolve-Path $InputFile).Path

# Nếu OutputFile không có thư mục, mặc định là thư mục hiện tại
$OutputPath = (Resolve-Path ".\").Path + "\" + (Split-Path $OutputFile -Leaf)

# Hằng số cho PDF
$xlTypePDF = 0

# Khởi tạo Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Mở workbook
$workbook = $excel.Workbooks.Open($InputPath)

# Xuất PDF
$workbook.ExportAsFixedFormat($xlTypePDF, $OutputPath)

# Đóng và dọn bộ nhớ
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
