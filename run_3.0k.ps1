[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$FILE_NAME = "3.0k_motor_control-ver3.4"
$ICON_NAME = "caoshen"

$DIST_DIR    = ".\dist\$FILE_NAME"
$INTERNAL_SRC = ".\_internal"
$INTERNAL_DST = "$DIST_DIR\_internal"
$RELEASE_DIR = ".\release"
$ZIP_PATH    = "$RELEASE_DIR\$FILE_NAME.zip"

# 使用 PyInstaller 打包 Python 脚本
..\venv\pyqt5\Scripts\pyinstaller `
    -D `
    -w `
    --noconsole `
    --icon ".\ico\$ICON_NAME.ico" `
    "$FILE_NAME.py"

# 检查 PyInstaller 是否成功
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller failed with exit code $LASTEXITCODE."
    pause
    exit 1
}

Write-Host "PyInstaller build success."

# 确保 _internal 目录存在
if (!(Test-Path $INTERNAL_DST)) {
    New-Item -ItemType Directory -Path $INTERNAL_DST | Out-Null
}

# 拷贝资源文件
Copy-Item "$INTERNAL_SRC\motor_control.ui" -Destination $INTERNAL_DST -Force
Copy-Item "$INTERNAL_SRC\background-image.png" -Destination $INTERNAL_DST -Force
# Copy-Item ".\Multi-stage_setting.xlsx" -Destination $DIST_DIR -Force

# 确保 release 目录存在
if (!(Test-Path $RELEASE_DIR)) {
    New-Item -ItemType Directory -Path $RELEASE_DIR | Out-Null
}

# 压缩为 ZIP
if (Test-Path $ZIP_PATH) {
    Remove-Item $ZIP_PATH -Force
}

Compress-Archive -Path $DIST_DIR -DestinationPath $ZIP_PATH -Force

Write-Host "Package complete:"
Write-Host "  -> $ZIP_PATH"

pause
