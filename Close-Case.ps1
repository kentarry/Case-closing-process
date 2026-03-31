<#
.SYNOPSIS
    案件結案流程自動化腳本 - 本機網路磁碟操作 (步驟 5 & 6)

.DESCRIPTION
    處理以下結案步驟：
    步驟 5: 在 0_GT測試交接 中找到對應案件資料夾，移至已結案，並複製問題圖檔
    步驟 6: 在 圖影片資料存放區\GT 中建立對應日期資料夾，放入問題圖檔

.PARAMETER CaseName
    案件名稱，格式為 "6位數字:案件描述" 或 "6位數字 案件描述"
    範例: "235927:滿貫大亨_活動序號兌換測試"

.PARAMETER CaseNumber
    案件單號 (6位數字)，搭配 CaseDescription 使用

.PARAMETER CaseDescription
    案件描述，搭配 CaseNumber 使用

.EXAMPLE
    .\Close-Case.ps1 -CaseName "235927:滿貫大亨_活動序號兌換測試"
    .\Close-Case.ps1 -CaseNumber "235927" -CaseDescription "滿貫大亨_活動序號兌換測試"
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$CaseName,

    [Parameter()]
    [string]$CaseNumber,

    [Parameter()]
    [string]$CaseDescription
)

# ===== 設定 UTF-8 編碼 =====
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

# ===== 載入設定檔 =====
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = Join-Path $scriptDir "config.json"

if (-not (Test-Path $configPath)) {
    Write-Error "❌ 找不到設定檔: $configPath"
    exit 1
}

$config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$networkBase = $config.network_base
$gtTestRelPath = $config.gt_test_path
$imageStoreRelPath = $config.image_store_path
$excludedFoldersGT = $config.excluded_folders_gt

# ===== 解析案件名稱 =====
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  案件結案流程自動化" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# 支援兩種輸入方式：
# 1. -CaseNumber + -CaseDescription (來自 GUI 拆分欄位)
# 2. -CaseName (合併格式: "235927:滿貫大亨_活動序號兌換測試")
if ($CaseNumber -and $CaseDescription) {
    $caseNumber = $CaseNumber.Trim()
    $caseDescription = $CaseDescription.Trim()
}
elseif ($CaseName) {
    # 支援格式: "235927:滿貫大亨_活動序號兌換測試" 或 "235927 滿貫大亨_活動序號兌換測試" 或 "235927滿貫大亨_活動序號兌換測試"
    if ($CaseName -match '^(\d{6})[:\s]*(.+)$') {
        $caseNumber = $Matches[1]
        $caseDescription = $Matches[2].Trim()
    }
    else {
        Write-Error "❌ 無效的案件名稱格式。預期格式: 6位數字 + 案件描述"
        Write-Error "   範例: 235927:滿貫大亨_活動序號兌換測試"
        exit 1
    }
}
else {
    Write-Error "❌ 請提供案件資訊。使用 -CaseName 或 -CaseNumber + -CaseDescription"
    exit 1
}

if ($caseNumber -notmatch '^\d{6}$') {
    Write-Error "❌ 案件單號必須為 6 位數字，目前為: $caseNumber"
    exit 1
}

Write-Host "📋 案件編號: $caseNumber" -ForegroundColor Green
Write-Host "📋 案件描述: $caseDescription" -ForegroundColor Green


# ===== 取得日期資訊 =====
$now = Get-Date
$yearMonth = $now.ToString("yyyyMM")           # e.g., 202603
$yearChinese = "$($now.Year)年"                 # e.g., 2026年
$monthChinese = "$($now.Month)月"               # e.g., 3月
$dayFolder = $now.ToString("MMdd")              # e.g., 0330

Write-Host "📅 當前年月: $yearMonth ($yearChinese $monthChinese)" -ForegroundColor Yellow
Write-Host "📅 當日資料夾: $dayFolder`n" -ForegroundColor Yellow

# ===== 步驟 5: 在 0_GT測試交接 找到並處理案件 =====
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Magenta
Write-Host "  步驟 5: 處理 0_GT測試交接" -ForegroundColor Magenta
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor Magenta

$gtTestBase = Join-Path $networkBase $gtTestRelPath
$yearMonthPath = Join-Path $gtTestBase $yearMonth
$closedCasePath = Join-Path $yearMonthPath "已結案"

# 檢查年月資料夾是否存在
if (-not (Test-Path $yearMonthPath)) {
    Write-Error "❌ 找不到年月資料夾: $yearMonthPath"
    exit 1
}

Write-Host "🔍 搜尋案件資料夾 (案件編號: $caseNumber)..." -ForegroundColor White

# 搜尋範圍：年月資料夾 (排除特定資料夾) + 外包子資料夾
$projectFolder = $null
$foundInOutsource = $false

# 先搜尋主資料夾
$allFolders = Get-ChildItem $yearMonthPath -Directory -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -notin ($excludedFoldersGT + @("外包", "已結案"))
}

foreach ($folder in $allFolders) {
    if ($folder.Name -match [regex]::Escape($caseNumber)) {
        $projectFolder = $folder
        Write-Host "✅ 找到案件資料夾: $($folder.Name)" -ForegroundColor Green
        Write-Host "   路徑: $($folder.FullName)" -ForegroundColor DarkGray
        break
    }
}

# 若主資料夾沒找到，搜尋外包資料夾
if (-not $projectFolder) {
    $outsourcePath = Join-Path $yearMonthPath "外包"
    if (Test-Path $outsourcePath) {
        $outsourceFolders = Get-ChildItem $outsourcePath -Directory -ErrorAction SilentlyContinue
        foreach ($folder in $outsourceFolders) {
            if ($folder.Name -match [regex]::Escape($caseNumber)) {
                $projectFolder = $folder
                $foundInOutsource = $true
                Write-Host "✅ 在外包資料夾找到案件: $($folder.Name)" -ForegroundColor Green
                Write-Host "   路徑: $($folder.FullName)" -ForegroundColor DarkGray
                break
            }
        }
    }
}

if (-not $projectFolder) {
    Write-Warning "⚠️ 在 0_GT測試交接\$yearMonth 中找不到案件編號 $caseNumber 的資料夾"
    Write-Warning "   已搜尋: $yearMonthPath 及 外包 子資料夾"
    Write-Host "`n跳過步驟 5 和步驟 6..." -ForegroundColor Yellow
}
else {
    # 5a. 建立已結案資料夾 (如果不存在)
    if (-not (Test-Path $closedCasePath)) {
        Write-Host "📁 建立已結案資料夾: $closedCasePath" -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $closedCasePath -Force | Out-Null
    }

    # 5b. 移動專案資料夾至已結案
    $destProjectPath = Join-Path $closedCasePath $projectFolder.Name
    if (Test-Path $destProjectPath) {
        Write-Warning "⚠️ 目標已存在: $destProjectPath，將覆蓋"
        Remove-Item $destProjectPath -Recurse -Force
    }

    Write-Host "`n📦 移動案件資料夾至已結案..." -ForegroundColor Cyan
    Write-Host "   來源: $($projectFolder.FullName)" -ForegroundColor DarkGray
    Write-Host "   目標: $destProjectPath" -ForegroundColor DarkGray
    Move-Item -Path $projectFolder.FullName -Destination $destProjectPath -Force
    Write-Host "   ✅ 移動完成!" -ForegroundColor Green

    # ===== 步驟 6: 在 圖影片資料存放區\GT 建立資料夾並放入圖檔 =====
    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Magenta
    Write-Host "  步驟 6: 處理 圖影片資料存放區\GT" -ForegroundColor Magenta
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor Magenta

    $imageStoreBase = Join-Path $networkBase $imageStoreRelPath

    # 建立目標路徑: GT\YYYY年\M月\MMDD\專案名稱\問題圖檔
    $yearPath = Join-Path $imageStoreBase $yearChinese
    $monthPath = Join-Path $yearPath $monthChinese
    $dayPath = Join-Path $monthPath $dayFolder
    $projectImagePath = Join-Path $dayPath $projectFolder.Name
    $issueImageDest = Join-Path $projectImagePath "問題圖檔"

    # 依序建立資料夾 (含問題圖檔子資料夾)
    foreach ($pathToCreate in @($yearPath, $monthPath, $dayPath, $projectImagePath, $issueImageDest)) {
        if (-not (Test-Path $pathToCreate)) {
            Write-Host "📁 建立資料夾: $pathToCreate" -ForegroundColor Cyan
            New-Item -ItemType Directory -Path $pathToCreate -Force | Out-Null
        }
    }
    Write-Host "✅ 已建立: $issueImageDest" -ForegroundColor Green

    # 從已結案的資料夾中複製問題圖檔
    $issueImageSource = Join-Path $destProjectPath "問題圖檔"
    if (Test-Path $issueImageSource) {
        $imageFiles = Get-ChildItem $issueImageSource -File -Recurse -ErrorAction SilentlyContinue
        if ($imageFiles.Count -gt 0) {
            Write-Host "`n📸 從已結案資料夾複製問題圖檔 ($($imageFiles.Count) 個)..." -ForegroundColor Cyan
            foreach ($img in $imageFiles) {
                $relativePath = $img.FullName.Substring($issueImageSource.Length + 1)
                $finalDest = Join-Path $issueImageDest $relativePath
                $finalDestDir = Split-Path $finalDest -Parent
                if (-not (Test-Path $finalDestDir)) {
                    New-Item -ItemType Directory -Path $finalDestDir -Force | Out-Null
                }
                Copy-Item $img.FullName -Destination $finalDest -Force
                Write-Host "   ✓ $relativePath" -ForegroundColor DarkGreen
            }
            Write-Host "   ✅ 圖檔複製完成!" -ForegroundColor Green
        }
        else {
            Write-Host "📸 問題圖檔資料夾為空" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "📸 已結案資料夾中無問題圖檔子資料夾" -ForegroundColor Yellow
    }
}

# ===== 完成 =====
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  ✅ 本機結案流程完成!" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "📝 摘要:" -ForegroundColor White
Write-Host "   案件: $caseNumber $caseDescription" -ForegroundColor White
if ($projectFolder) {
    Write-Host "   ✅ 步驟 5: 已將案件資料夾移至已結案" -ForegroundColor Green
    Write-Host "   ✅ 步驟 6: 已建立圖影片資料夾 (含問題圖檔)" -ForegroundColor Green
}
else {
    Write-Host "   ⏭️ 步驟 5-6: 未找到案件資料夾，已跳過" -ForegroundColor Yellow
}


