<#
===============================================================================
 流放之路 1 (PoE1) 国际服 —— 繁体中文一键汉化脚本
 方案 1：PoeChinese3 (LibGGPK3 / aianlinb)
===============================================================================

【方案原理】
 PoE1 国际服客户端里其实自带完整的繁体中文语系数据（zh-TW），只是被锁住。
 PoeChinese3 不下载任何翻译包，只做三件事：
   1. 把 Data/Languages.dat 里 French 与 Traditional Chinese 两行的显示名互换
      —— 于是游戏里选「第二个（法国）国旗」实际就是繁体中文
   2. 若工具同目录存在 Font.ttf / Font.ttc，则替换游戏内字体
   3. 交换登录画面的国旗图标（Art/UIImages1.txt）
 最后写回 Content.ggpk（独立客户端）或 Bundles2/_.index.bin（Steam/Epic）。

【为什么适合「每次更新后跑一次」】
 - 无需 PIN 码、无需等编年史出汉化包、不受国服版本进度影响
 - 工具源码带幂等判断：Languages.dat 已套用过就直接跳过，重复运行不会叠坏档
 - 游戏每次更新会重置 Languages.dat / _.index.bin，所以更新后重跑本脚本即可
 - 只改极少数文件，可随时在游戏内切回英文原版

【前置条件（一次性）】
 - Windows；若报 VCRUNTIME140_1.dll 缺失，安装 Microsoft Visual C++ Redistributable
 - 运行前完全关闭游戏与启动器（脚本会自动检测并拒绝执行）
 - 杀毒软件可能误报，必要时临时关闭或加白名单
 - 官网市集设定 https://www.pathofexile.com/trade/settings
   把 Whisper Language 设为 English，否则交易密语会错乱

【用法】
  双击 run_hanhua.cmd                       # 一键：自动找游戏 -> 更新工具 -> 汉化
  powershell -File Poe1_TW.ps1 -ListTargets          # 只列出探测到的游戏文件
  powershell -File Poe1_TW.ps1 -DryRun               # 预演，不修改任何文件
  powershell -File Poe1_TW.ps1 -GamePath "D:\Games\Path of Exile"
  powershell -File Poe1_TW.ps1 -SetLanguageFr        # 顺带把 language 改成 fr（不想手动点国旗时用）
  powershell -File Poe1_TW.ps1 -SkipDownload         # 离线，只用 tools 里缓存的工具

 若自动探测找错路径：在脚本同目录建 poe1_path.txt，
 第一行写游戏根目录，或直接写 Content.ggpk / _.index.bin 的完整路径。

【进游戏后】
 选择第二个（法国）国旗 = 繁体中文。若无法点选，加 -SetLanguageFr 重跑
 （等价于改 Documents\My Games\Path of Exile\production_Config.ini 的 language=fr）。
 若启动器又开始读条还原文件：重跑本脚本；或先 PackCheck.exe 再重跑；
 也可直接用游戏目录下的 Client.exe 启动以绕过启动器覆盖检查。

【资料来源】
 - PoEDB 流亡编年史中文化页: https://poedb.tw/cn/chinese
 - LibGGPK3 项目 / 工具下载: https://github.com/aianlinb/LibGGPK3/releases
 - PoeChinese3 源码（命令行参数、幂等判断）: Examples/PoeChinese3/Program.cs
 注：修改游戏文件属官方 ToU 灰色地带，纯语系替换目前为默许状态，风险自负。
===============================================================================
#>

[CmdletBinding()]
param(
    [string]$GamePath,
    [string]$ToolUrl = 'https://github.com/aianlinb/LibGGPK3/releases/latest/download/PoeChinese3_win-x64.exe',
    [switch]$SkipDownload,
    [switch]$SetLanguageFr,
    [switch]$DryRun,
    [switch]$ListTargets,
    [switch]$Force,
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
# 固定控制台为 UTF-8：中文提示 + 工具输出都不乱码（PS 5.1 / PS 7 通用）
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8
# Windows PowerShell 5.1 默认不启用 TLS 1.2，GitHub 下载会失败
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
$script:Root = $PSScriptRoot
$script:ToolPath = Join-Path $script:Root 'tools\PoeChinese3_win-x64.exe'
$script:PathFile = Join-Path $script:Root 'poe1_path.txt'
$script:VersionUrl = 'https://raw.githubusercontent.com/aianlinb/LibGGPK3/main/.github/Version.txt'

function Write-Step { param([string]$Msg) Write-Host "  $Msg" -ForegroundColor Gray }
function Write-Ok   { param([string]$Msg) Write-Host "[OK]   $Msg" -ForegroundColor Green }
function Write-Warn2{ param([string]$Msg) Write-Host "[WARN] $Msg" -ForegroundColor Yellow }
function Write-Err  { param([string]$Msg) Write-Host "[FAIL] $Msg" -ForegroundColor Red }
function Write-Head { param([string]$Msg) Write-Host ''; Write-Host "== $Msg ==" -ForegroundColor Cyan }

# ---------- 1. 路径解析 ----------
function Get-IndexBin {
    param([string]$Dir)
    $bundles = Join-Path $Dir 'Bundles2'
    if (-not (Test-Path -LiteralPath $bundles -PathType Container)) { return $null }
    # 官方客户端文件名是 _.index.bin；个别版本/渠道可能略有差异，做兼容探测
    foreach ($n in @('_.index.bin', '_index.bin')) {
        $p = Join-Path $bundles $n
        if (Test-Path -LiteralPath $p -PathType Leaf) { return (Get-Item -LiteralPath $p) }
    }
    return (Get-ChildItem -LiteralPath $bundles -File -Filter '*index.bin' -ErrorAction SilentlyContinue |
            Select-Object -First 1)
}

function Resolve-GameTarget {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    if (-not $item.PSIsContainer) {
        if ($item.Extension -in @('.ggpk', '.bin')) { return $item }
        return $null
    }
    $ggpk = Join-Path $item.FullName 'Content.ggpk'
    if (Test-Path -LiteralPath $ggpk -PathType Leaf) { return (Get-Item -LiteralPath $ggpk) }
    return (Get-IndexBin -Dir $item.FullName)
}

function Find-GameCandidates {
    $roots = New-Object System.Collections.Generic.List[string]

    # 从正在运行的 PoE 进程反查安装目录（能覆盖任意自定义安装路径）
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.ProcessName -notlike 'PathOfExile*' -and $p.ProcessName -ne 'Client') { continue }
        try {
            $exePath = $p.Path
            if ($exePath) {
                $dir = Split-Path -Parent $exePath
                # 排除 PoE2 客户端，避免误汉化
                if ($dir -notmatch '(?i)poe2|path\s*of\s*exile\s*2') { $roots.Add($dir) }
            }
        } catch {}
    }

    # Steam 安装目录 + 其他库（libraryfolders.vdf）
    $steam = $null
    try { $steam = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam' -Name 'InstallPath' -ErrorAction Stop).InstallPath } catch {}
    if (-not $steam) {
        try { $steam = (Get-ItemProperty -Path 'HKCU:\Software\Valve\Steam' -Name 'InstallPath' -ErrorAction Stop).InstallPath } catch {}
    }
    if ($steam) {
        $roots.Add((Join-Path $steam 'steamapps\common\Path of Exile'))
        $vdf = Join-Path $steam 'steamapps\libraryfolders.vdf'
        if (Test-Path -LiteralPath $vdf) {
            $txt = Get-Content -LiteralPath $vdf -Raw -ErrorAction SilentlyContinue
            if ($txt) {
                foreach ($m in [regex]::Matches($txt, '"path"\s*"([^"]+)"')) {
                    $lib = $m.Groups[1].Value.Replace('\\', '\')
                    $roots.Add((Join-Path $lib 'steamapps\common\Path of Exile'))
                }
            }
        }
    }

    # 常见安装位置（只做浅层检查，不递归扫描磁盘）
    $drives = @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | ForEach-Object { $_.Root })
    foreach ($d in $drives) {
        $roots.Add((Join-Path $d 'Program Files (x86)\Grinding Gear Games\Path of Exile'))
        $roots.Add((Join-Path $d 'Program Files\Grinding Gear Games\Path of Exile'))
        $roots.Add((Join-Path $d 'Program Files (x86)\Steam\steamapps\common\Path of Exile'))
        $roots.Add((Join-Path $d 'Program Files\Steam\steamapps\common\Path of Exile'))
        $roots.Add((Join-Path $d 'Program Files\Epic Games\Path of Exile'))
        $roots.Add((Join-Path $d 'SteamLibrary\steamapps\common\Path of Exile'))
        $roots.Add((Join-Path $d 'Games\Path of Exile'))
        $roots.Add((Join-Path $d 'Path of Exile'))
    }

    $found = @{}
    foreach ($r in $roots) {
        $t = Resolve-GameTarget -Path $r
        if ($null -ne $t -and -not $found.ContainsKey($t.FullName)) { $found[$t.FullName] = $t }
    }
    return @($found.Values | Sort-Object -Property LastWriteTime -Descending)
}

function Get-ConfiguredTarget {
    if (-not [string]::IsNullOrWhiteSpace($GamePath)) { return (Resolve-GameTarget -Path $GamePath) }
    if (-not [string]::IsNullOrWhiteSpace($env:POE1_PATH)) { return (Resolve-GameTarget -Path $env:POE1_PATH) }
    if (Test-Path -LiteralPath $script:PathFile) {
        foreach ($line in (Get-Content -LiteralPath $script:PathFile -ErrorAction SilentlyContinue)) {
            $s = $line.Trim().Trim('"')
            if ($s -and -not $s.StartsWith('#')) { return (Resolve-GameTarget -Path $s) }
        }
    }
    return $null
}

# ---------- 2. 进程检测 ----------
function Get-RunningGameProcesses {
    param([string]$GameRoot)
    $list = @()
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
        $n = $p.ProcessName
        if ($n -like 'PathOfExile*') {
            $list += "$n (PID $($p.Id))"
        } elseif ($n -eq 'Client' -and $GameRoot) {
            try {
                if ($p.Path -and $p.Path.StartsWith($GameRoot, [StringComparison]::OrdinalIgnoreCase)) {
                    $list += "$n (PID $($p.Id))"
                }
            } catch {}
        }
    }
    return $list
}

# ---------- 3. 工具获取/更新 ----------
function Update-Tool {
    if (-not (Test-Path -LiteralPath (Split-Path -Parent $script:ToolPath))) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $script:ToolPath) -Force | Out-Null
    }
    if ($SkipDownload) {
        if (Test-Path -LiteralPath $script:ToolPath) { Write-Ok "离线模式：使用缓存工具"; return $true }
        Write-Err "离线模式但 tools 目录下没有工具，请先联网运行一次或手动放置 PoeChinese3_win-x64.exe"
        return $false
    }
    $tmp = "$($script:ToolPath).download"
    try {
        Write-Step "检查/下载最新版 PoeChinese3 ..."
        Invoke-WebRequest -Uri $ToolUrl -OutFile $tmp -UseBasicParsing -TimeoutSec 90 -ErrorAction Stop
        $len = (Get-Item -LiteralPath $tmp).Length
        if ($len -lt 1MB) { throw "下载内容异常（$len 字节），疑似被网络劫持" }
        Move-Item -LiteralPath $tmp -Destination $script:ToolPath -Force
        Write-Ok "工具已更新 ($([math]::Round($len / 1MB, 2)) MB)"
        return $true
    } catch {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $script:ToolPath) {
            Write-Warn2 "下载失败，继续使用缓存版本：$($_.Exception.Message)"
            return $true
        }
        Write-Err "工具下载失败且无缓存：可手动下载 $ToolUrl 放到 tools\PoeChinese3_win-x64.exe，或用 -ToolUrl 指定镜像"
        return $false
    }
}

function Test-ToolVersion {
    if (-not (Test-Path -LiteralPath $script:ToolPath)) { return }
    try {
        $fi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($script:ToolPath)
        $local = [version]$fi.FileVersion
        Write-Step "本地工具版本: $local"
        if ($SkipDownload) { return }
        $txt = (Invoke-WebRequest -Uri $script:VersionUrl -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop).Content
        $lines = @($txt -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { return }
        $latest = [version]$lines[0]
        $minimum = if ($lines.Count -gt 1) { [version]$lines[1] } else { $null }
        if ($local -lt $latest) { Write-Warn2 "线上最新版为 $latest，建议允许脚本自动更新工具" }
        if ($minimum -and $local -lt $minimum) {
            throw "本地工具版本 $local 低于最低可用版本 $minimum，旧版本存在已知 Bug，请更新后再运行"
        }
    } catch {
        if ($_.Exception.Message -like '*低于最低可用版本*') { throw }
        Write-Warn2 "版本检查跳过（网络不可达）：$($_.Exception.Message)"
    }
}

# ---------- 4. 备份 ----------
function Backup-IndexBin {
    param([System.IO.FileInfo]$Target)
    if ($Target.Extension -ine '.bin') {
        Write-Step '独立客户端 Content.ggpk 体积过大，跳过整体备份（工具改动极小且幂等，可用启动器/PackCheck.exe 还原）'
        return
    }
    try {
        $dir = Join-Path $script:Root 'backups'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $dst = Join-Path $dir "index_$stamp.bak"
        Copy-Item -LiteralPath $Target.FullName -Destination $dst -Force
        Write-Ok "已备份 _.index.bin -> $dst"
        Get-ChildItem -LiteralPath $dir -Filter 'index_*.bak' -File |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip 3 |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Warn2 "备份失败（继续汉化）：$($_.Exception.Message)"
    }
}

# ---------- 5. 执行汉化 ----------
function Invoke-ChinesePatch {
    param([System.IO.FileInfo]$Target)
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Step "目标文件: $($Target.FullName)"
    Write-Step "文件大小: $([math]::Round($Target.Length / 1MB, 1)) MB，最后修改 $($Target.LastWriteTime)"
    # 工具结尾会 Console.ReadLine() 等待回车，管道送入一个空行避免卡住
    $out = '' | & $script:ToolPath $Target.FullName 2>&1 | Out-String
    $code = $LASTEXITCODE
    Write-Host "----- PoeChinese3 输出 -----" -ForegroundColor DarkGray
    Write-Host $out.TrimEnd()
    Write-Host "---------------------------" -ForegroundColor DarkGray
    if ($out -match '(?m)Done!') {
        return @{ Ok = $true; Output = $out }
    }
    return @{ Ok = $false; Output = $out; Code = $code }
}

# ---------- 6. 可选：自动写入 language=fr ----------
function Set-LanguageConfig {
    $ini = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) 'My Games\Path of Exile\production_Config.ini'
    if (-not (Test-Path -LiteralPath $ini)) {
        Write-Warn2 "未找到 $ini ，请进游戏手动选择第二个（法国）国旗"
        return
    }
    try {
        $text = [System.IO.File]::ReadAllText($ini, [System.Text.Encoding]::UTF8)
        if ($text -match '(?m)^\s*language\s*=\s*fr\s*$') { Write-Ok '游戏语言配置已是 fr，无需修改'; return }
        Copy-Item -LiteralPath $ini -Destination "$ini.bak" -Force
        if ($text -match '(?m)^\s*language\s*=\s*\S+') {
            $new = [regex]::Replace($text, '(?m)^(\s*language\s*=\s*)\S+', '${1}fr')
        } else {
            Write-Warn2 '配置里没有 language= 项，未自动修改；请进游戏手动选法国国旗'
            return
        }
        [System.IO.File]::WriteAllText($ini, $new, (New-Object System.Text.UTF8Encoding($false)))
        Write-Ok "已把 $ini 的 language 设为 fr（原文件备份为 .bak）"
    } catch {
        Write-Warn2 "写入语言配置失败：$($_.Exception.Message)，请进游戏手动选第二个国旗"
    }
}

# ================================ 主流程 ================================
Write-Host ''
Write-Host '  Path of Exile 1 / 繁体中文一键汉化 (PoeChinese3)' -ForegroundColor Cyan
Write-Host '  来源: poedb.tw/cn/chinese  |  github.com/aianlinb/LibGGPK3' -ForegroundColor DarkGray
Write-Host ''

try {
    # 探测
    Write-Head '1/5 定位游戏数据文件'
    $candidates = @(Find-GameCandidates)
    if ($ListTargets) {
        if ($candidates.Count -eq 0) { Write-Warn2 '未自动探测到 PoE1 客户端，请用 -GamePath 指定或创建 poe1_path.txt' }
        else { foreach ($c in $candidates) { Write-Host ("  {0}  ({1:N1} MB, {2})" -f $c.FullName, ($c.Length / 1MB), $c.LastWriteTime) } }
        exit 0
    }

    $target = Get-ConfiguredTarget
    if ($null -eq $target) {
        if ($candidates.Count -eq 0) {
            throw '找不到 Content.ggpk 或 Bundles2\_.index.bin。请确认游戏已安装完成，或用 -GamePath 指定游戏根目录。'
        }
        $target = $candidates[0]
        if ($candidates.Count -gt 1) {
            Write-Warn2 "探测到多个客户端，默认选用最近修改的那个；如需指定请建 poe1_path.txt 或用 -GamePath"
            foreach ($c in $candidates) { Write-Step "  候选: $($c.FullName)" }
        }
    }
    Write-Ok "已选定: $($target.FullName)"

    $gameRoot = Split-Path -Parent $target.FullName
    if ($target.Extension -ieq '.bin') { $gameRoot = Split-Path -Parent $gameRoot }

    # 进程检查
    Write-Head '2/5 检查游戏进程'
    $running = @(Get-RunningGameProcesses -GameRoot $gameRoot)
    if ($running.Count -gt 0) {
        if (-not $Force) {
            throw "检测到游戏/启动器正在运行：$($running -join ', ')。请等更新完成并完全退出游戏后再运行本脚本（-Force 可跳过检查，不推荐）。"
        }
        Write-Warn2 "已用 -Force 跳过进程检查：$($running -join ', ')"
    } else {
        Write-Ok '游戏未运行，可以安全写入'
    }

    # 工具
    Write-Head '3/5 准备汉化工具'
    if (-not (Update-Tool)) { exit 2 }
    if (-not (Test-Path -LiteralPath $script:ToolPath)) { throw '工具文件不存在' }
    Test-ToolVersion

    if ($DryRun) {
        Write-Head '预演结束'
        Write-Ok "预演完成：将对 $($target.FullName) 执行汉化。未修改任何文件。"
        exit 0
    }

    # 备份 + 执行
    Write-Head '4/5 备份并执行汉化'
    try {
        $di = New-Object System.IO.DriveInfo -ArgumentList (Split-Path -Qualifier $target.FullName)
        $freeGB = $di.AvailableFreeSpace / 1GB
        Write-Step ("所在磁盘剩余空间: {0:N1} GB" -f $freeGB)
        if ($freeGB -lt 1) { throw '游戏所在磁盘剩余空间不足 1 GB，写入 ggpk/index 可能失败，请先清理空间' }
    } catch {
        if ($_.Exception.Message -like '*剩余空间不足*') { throw }
        Write-Warn2 "磁盘空间检查跳过：$($_.Exception.Message)"
    }
    Backup-IndexBin -Target $target
    $result = Invoke-ChinesePatch -Target $target
    if (-not $result.Ok) {
        Write-Err 'PoeChinese3 未返回成功标志（Done!），详见上方输出。常见原因：'
        Write-Step ' · 缺 VCRUNTIME140_1.dll -> 安装 Microsoft Visual C++ Redistributable'
        Write-Step ' · 杀毒软件拦截了写入 -> 加白名单后重跑'
        Write-Step ' · 数据文件正在被游戏更新占用 -> 等更新结束再重跑'
        exit 3
    }
    Write-Ok '汉化写入完成（此操作是幂等的，重复运行不会损坏数据）'

    # 记住路径，便于下次一键运行（内容不变则不写，避免重复运行时产生无谓改动）
    try {
        $want = "# PoE1 汉化目标路径（可改成游戏根目录，或直接写 Content.ggpk / _.index.bin 的完整路径）`n$($target.FullName)`n"
        $old = ''
        if (Test-Path -LiteralPath $script:PathFile) {
            $old = [System.IO.File]::ReadAllText($script:PathFile, [System.Text.Encoding]::UTF8) -replace "`r`n", "`n"
        }
        if ($old -ne $want) {
            [System.IO.File]::WriteAllText($script:PathFile, $want, (New-Object System.Text.UTF8Encoding($false)))
        }
    } catch {}

    if ($SetLanguageFr) { Set-LanguageConfig }

    Write-Head '5/5 完成'
    Write-Host '  进游戏后选择「第二个（法国）国旗」即为繁体中文。' -ForegroundColor White
    Write-Host '  首次使用请到官网市集设定把 Whisper Language 设为 English。' -ForegroundColor White
    Write-Host '  若启动器又开始读条还原文件：重跑本脚本，或先 PackCheck.exe 再重跑；' -ForegroundColor White
    Write-Host '  也可直接用游戏目录下的 Client.exe 启动以绕过启动器覆盖检查。' -ForegroundColor White
    Write-Host '  以后每次游戏更新完成后，重新双击 run_hanhua.cmd 即可。' -ForegroundColor Green
    Write-Host ''
    exit 0
} catch {
    Write-Host ''
    Write-Err $_.Exception.Message
    Write-Host ''
    exit 1
}
