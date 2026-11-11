# Windows PowerShell 5.1 / PowerShell 7 対応。日本語のためUTF-8 BOMで保存する。
[CmdletBinding()]
param(
    [string]$Python指定 = '3.15'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

# ── 1. 引数・ログ・進捗の共通処理 ──────────────────────
function 引数を引用 {
    param([AllowEmptyString()][string]$値)
    if ($値 -ne '' -and $値 -notmatch '[\s"]') { return $値 }
    # .NET Frameworkにも対応するWindowsネイティブ引数の引用。
    $組立 = New-Object System.Text.StringBuilder
    [void]$組立.Append('"')
    $斜線数 = 0
    foreach ($文字 in $値.ToCharArray()) {
        if ($文字 -eq '\') {
            $斜線数++
        } elseif ($文字 -eq '"') {
            [void]$組立.Append(('\' * ($斜線数 * 2 + 1)))
            [void]$組立.Append('"')
            $斜線数 = 0
        } else {
            [void]$組立.Append(('\' * $斜線数))
            [void]$組立.Append($文字)
            $斜線数 = 0
        }
    }
    [void]$組立.Append(('\' * ($斜線数 * 2)))
    [void]$組立.Append('"')
    return $組立.ToString()
}

function ビルドログを表示 {
    param([string]$内容, [ConsoleColor]$色 = [ConsoleColor]::Gray)
    Write-Host $内容 -ForegroundColor $色
    if ($script:ログ書込器) {
        $script:ログ書込器.WriteLine(('{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $内容))
    }
}

function 進捗を表示 {
    param([int]$率, [string]$工程)
    $率 = [Math]::Max($script:進捗率, [Math]::Min(100, $率))
    Write-Progress -Id 1 -Activity 'VBA2SRC EXEビルド' -Status "$工程 : $率%" -PercentComplete $率
    if ($工程 -ne $script:直前工程) {
        ビルドログを表示 ''
        ビルドログを表示 ('  > ' + $工程) Cyan
    }
    $script:進捗率 = $率
    $script:直前工程 = $工程
}

function 結果を表示 {
    param([string]$状態, [string]$説明, [string[]]$詳細 = @())
    $色 = switch ($状態) { '完了' { 'Green' }; '中止' { 'Yellow' }; default { 'Red' } }
    # 状態を色と文言の両方で示し、ログでも読み取れる結果ブロックにする。
    ビルドログを表示 ''
    ビルドログを表示 ('  ' + ('─' * 56)) $色
    ビルドログを表示 ('  VBA2SRC  ビルド' + $状態) $色
    ビルドログを表示 ('  ' + $説明)
    foreach ($行 in $詳細) { ビルドログを表示 ('  ' + $行) }
    ビルドログを表示 ('  ' + ('─' * 56)) $色
    ビルドログを表示 ''
}

# ── 2. 外部工程の実行と出力監視 ────────────────────────────
function 外部コマンドを実行 {
    param(
        [string]$実行パス,
        [string[]]$引数一覧,
        [string]$工程名,
        [int]$開始率,
        [int]$終了率,
        [switch]$PyInstaller工程
    )
    進捗を表示 $開始率 $工程名
    ビルドログを表示 ('実行：{0} {1}' -f $実行パス, (($引数一覧 | ForEach-Object { 引数を引用 $_ }) -join ' ')) DarkGray
    $起動情報 = New-Object System.Diagnostics.ProcessStartInfo
    $起動情報.FileName = $実行パス
    $起動情報.Arguments = ($引数一覧 | ForEach-Object { 引数を引用 $_ }) -join ' '
    $起動情報.WorkingDirectory = $script:ルート
    $起動情報.UseShellExecute = $false
    $起動情報.CreateNoWindow = $true
    $起動情報.RedirectStandardOutput = $true
    $起動情報.RedirectStandardError = $true
    $起動情報.RedirectStandardInput = $true
    $起動情報.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $起動情報.StandardErrorEncoding = New-Object System.Text.UTF8Encoding($false)
    $起動情報.EnvironmentVariables['PYTHONUTF8'] = '1'
    $起動情報.EnvironmentVariables['PYTHONIOENCODING'] = 'utf-8'
    $起動情報.EnvironmentVariables['PYTHONUNBUFFERED'] = '1'
    $プロセス = New-Object System.Diagnostics.Process
    $プロセス.StartInfo = $起動情報
    $計時 = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        [void]$プロセス.Start()
        $プロセス.StandardInput.Close()
        # 両方を非同期に読み取り、stderrが多いビルドでも詰まらせない。
        $読取待ち = @($プロセス.StandardOutput.ReadLineAsync(), $プロセス.StandardError.ReadLineAsync())
        while (-not $プロセス.HasExited -or $読取待ち[0] -or $読取待ち[1]) {
            foreach ($番号 in 0, 1) {
                if ($読取待ち[$番号] -and $読取待ち[$番号].IsCompleted) {
                    $行 = $読取待ち[$番号].GetAwaiter().GetResult()
                    if ($null -eq $行) {
                        $読取待ち[$番号] = $null
                    } else {
                        $表示色 = if ($行 -match '\bERROR\b|Traceback') { 'Red' } elseif ($行 -match '\bWARNING\b|\bWARN\b') { 'Yellow' } else { 'Gray' }
                        ビルドログを表示 $行 $表示色
                        if ($PyInstaller工程) {
                            # PyInstallerには正確な進捗率APIがないため工程を目安にする。
                            $段階率 = switch -Regex ($行) {
                                'Initializing module dependency graph|Analyzing modules for' { 53; break }
                                'Analyzing .*vba2src.py' { 58; break }
                                'Processing module hooks' { 72; break }
                                'Processing .*hook' { 62; break }
                                'Looking for dynamic libraries' { 76; break }
                                'Building PYZ' { 80; break }
                                'Building PKG' { 83; break }
                                'Building EXE|Copying bootloader' { 86; break }
                                'Build complete' { 88; break }
                                default { $script:進捗率 }
                            }
                            進捗を表示 $段階率 $工程名
                        }
                        $読取待ち[$番号] = if ($番号 -eq 0) { $プロセス.StandardOutput.ReadLineAsync() } else { $プロセス.StandardError.ReadLineAsync() }
                    }
                }
            }
            # 出力が止まっていても、終了前は工程の上限に達しない概算を表示する。
            if (-not $PyInstaller工程) {
                $推定率 = $開始率 + [int](($終了率 - $開始率) * 0.8 * (1 - [Math]::Exp(-$計時.Elapsed.TotalSeconds / 25)))
                進捗を表示 $推定率 $工程名
            } else {
                進捗を表示 $script:進捗率 $工程名
            }
            if (-not $プロセス.HasExited) { [void]$プロセス.WaitForExit(80) }
        }
        $プロセス.WaitForExit()
        if ($プロセス.ExitCode -ne 0) {
            throw "$工程名 に失敗しました（終了コード $($プロセス.ExitCode)）。"
        }
        進捗を表示 $終了率 $工程名
    } finally {
        if ($プロセス.Id -and -not $プロセス.HasExited) { $プロセス.Kill() }
        $プロセス.Dispose()
        $計時.Stop()
    }
}

# ── 3. 実行確認 ──────────────────────────────────────────
$script:ルート = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$script:ログ書込器 = $null
try {
    Write-Host (Get-Content -LiteralPath (Join-Path $script:ルート 'vba2src_app/title.txt') -Raw -Encoding UTF8) -ForegroundColor Cyan
} catch {
    結果を表示 '失敗' 'タイトルファイルを読み込めません。' @('リポジトリ全体を配置して再実行してください。')
    exit 1
}
Write-Host '  WINDOWS EXE BUILDER' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  依存確認の後、Windows用EXEを作成します。'
Write-Host '  dist内の同名成果物を更新します。依存の自動導入は行いません。' -ForegroundColor DarkGray
Write-Host ''
$回答 = [string](Read-Host 'ビルドを実行しますか？ 続行する場合は yes を入力')
if (-not $回答.Trim().Equals('yes', [System.StringComparison]::OrdinalIgnoreCase)) {
    結果を表示 '中止' '成果物は更新していません。'
    exit 130
}

# ── 4. 依存確認・ビルド ──────────────────────────────
$script:進捗率 = -1
$script:直前工程 = ''
$script:ログ書込器 = $null
$終了コード = 1
try {
    $ログフォルダ = Join-Path $script:ルート 'build/logs'
    [void][System.IO.Directory]::CreateDirectory($ログフォルダ)
    $ログ保存先 = Join-Path $ログフォルダ ('build_{0}_{1}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), [Guid]::NewGuid().ToString('N').Substring(0, 8))
    $script:ログ書込器 = New-Object System.IO.StreamWriter($ログ保存先, $false, (New-Object System.Text.UTF8Encoding($false)))
    $script:ログ書込器.AutoFlush = $true
    $script:ログ書込器.NewLine = "`n"
    ビルドログを表示 "ビルドログ：$ログ保存先"
    $Pythonコマンド = (Get-Command py -CommandType Application -ErrorAction Stop).Source
    $Python引数 = @("-$Python指定")
    外部コマンドを実行 $Pythonコマンド ($Python引数 + @('scripts/check_dependencies.py')) '依存ライブラリ確認' 0 15
    外部コマンドを実行 $Pythonコマンド ($Python引数 + @('-m', 'pip', 'check')) '依存関係の整合性確認' 15 20
    外部コマンドを実行 $Pythonコマンド ($Python引数 + @('scripts/build_notices.py')) 'ライセンス収集' 20 50
    外部コマンドを実行 $Pythonコマンド ($Python引数 + @('-m', 'PyInstaller', '--clean', '--noconfirm', '--log-level', 'INFO', 'VBA2SRC.spec')) 'PyInstaller EXE作成' 50 88 -PyInstaller工程
    $配布先 = Join-Path $script:ルート 'dist'
    $実行ファイル = Join-Path $配布先 'VBA2SRC.exe'
    外部コマンドを実行 $実行ファイル @('--version') 'EXE起動確認' 88 96
    進捗を表示 97 'ライセンスを配置'
    Copy-Item -LiteralPath (Join-Path $script:ルート 'build/THIRD_PARTY_NOTICES.txt') -Destination $配布先
    Copy-Item -LiteralPath (Join-Path $script:ルート 'LICENSE') -Destination $配布先
    進捗を表示 100 '完了'
    Write-Progress -Id 1 -Activity 'VBA2SRC EXEビルド' -Completed
    結果を表示 '完了' 'EXEの作成と起動確認が完了しました。' @(
        ('EXE：' + $実行ファイル),
        ('ログ：' + $ログ保存先)
    )
    $終了コード = 0
} catch {
    Write-Progress -Id 1 -Activity 'VBA2SRC EXEビルド' -Completed
    $詳細一覧 = @('停止した工程：' + $script:直前工程)
    if ($ログ保存先) { $詳細一覧 += 'ログ：' + $ログ保存先 }
    $詳細一覧 += '詳細ログを確認し、原因を解消して再実行してください。'
    結果を表示 '失敗' $_.Exception.Message $詳細一覧
} finally {
    Write-Progress -Id 1 -Activity 'VBA2SRC EXEビルド' -Completed
    if ($script:ログ書込器) { $script:ログ書込器.Dispose() }
}
exit $終了コード
