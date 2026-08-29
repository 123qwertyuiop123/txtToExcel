$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$outputExe = Join-Path $projectRoot "dist\消费统计转换工具.exe"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "找不到项目虚拟环境：$pythonExe`n请先在项目目录创建 .venv。"
}

Push-Location $projectRoot
try {
    # 一键打包不执行任何安装，只使用项目虚拟环境中已经具备的工具。
    & $pythonExe -c "import openpyxl, PyInstaller"
    if ($LASTEXITCODE -ne 0) {
        throw "项目环境缺少 openpyxl 或 PyInstaller，无法执行免安装打包。"
    }

    # --onefile 生成单个 EXE；--windowed 避免双击时出现命令行黑窗口。
    # --noconfirm 只允许覆盖下方 dist 中的同名 EXE；--clean 只清理 PyInstaller 构建缓存。
    # build 和 dist 固定放在项目目录，不扫描或删除用户选择的消费文件。
    $pyinstallerArgs = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", "消费统计转换工具",
        "--distpath", (Join-Path $projectRoot "dist"),
        "--workpath", (Join-Path $projectRoot "build\work"),
        "--specpath", (Join-Path $projectRoot "build"),
        (Join-Path $projectRoot "main.py")
    )
    & $pythonExe @pyinstallerArgs
    if ($LASTEXITCODE -ne 0) {
        throw "打包失败，退出代码：$LASTEXITCODE"
    }
    if (-not (Test-Path -LiteralPath $outputExe)) {
        throw "打包命令已结束，但没有找到输出文件：$outputExe"
    }

    $file = Get-Item -LiteralPath $outputExe
    Write-Host ""
    Write-Host "打包完成：$($file.FullName)"
    Write-Host "文件大小：$([math]::Round($file.Length / 1MB, 2)) MB"
}
finally {
    Pop-Location
}
