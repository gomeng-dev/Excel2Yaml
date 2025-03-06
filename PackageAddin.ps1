# Excel To JSON/YAML 애드인 패키징 스크립트
# 이 스크립트는 애드인을 단일 설치 파일로 패키징합니다.

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projectName = "ExcelToYamlAddin"
$outputFolder = Join-Path $scriptPath "Deploy"
$zipFile = Join-Path $scriptPath "$projectName-Setup.zip"

Write-Host "애드인 패키지 생성 시작..." -ForegroundColor Green

# 필요한 폴더 생성
if(Test-Path $outputFolder) {
    Remove-Item -Path $outputFolder -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $outputFolder | Out-Null

# 릴리스 폴더 확인
$releaseFolder = Join-Path $scriptPath "bin\Release"
if(-not (Test-Path $releaseFolder)) {
    Write-Host "Release 폴더가 없습니다. Release 모드로 빌드를 먼저 수행해주세요." -ForegroundColor Red
    exit 1
}

# 필요한 파일 복사
Write-Host "파일 복사 중..." -ForegroundColor Green
Copy-Item -Path $releaseFolder -Destination $outputFolder -Recurse

# setup.bat 파일 생성
$setupContent = @'
@echo off
echo Excel To JSON/YAML Converter Add-in 설치 중...

REM 설치 디렉토리 생성
set INSTALL_DIR=%APPDATA%\Microsoft\AddIns\ExcelToYamlAddin
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM 파일 복사
xcopy /Y /E "%~dp0Release\*.*" "%INSTALL_DIR%\"

echo 설치가 완료되었습니다!
echo Excel을 열고 파일 > 옵션 > 추가 기능으로 이동하여 애드인이 활성화되었는지 확인하세요.
pause
'@

$setupFile = Join-Path $outputFolder "setup.bat"
$setupContent | Out-File -FilePath $setupFile -Encoding utf8

# README 파일 생성
$readmeContent = @'
# Excel To JSON/YAML Converter Add-in

이 애드인은 Excel 파일을 JSON 또는 YAML 형식으로 쉽게 변환할 수 있게 해줍니다.

## 설치 방법

1. setup.bat 파일을 실행하세요. (관리자 권한 필요)
2. 설치가 완료되면 Excel을 재시작하세요.
3. Excel의 "파일 > 옵션 > 추가 기능"에서 애드인이 활성화되었는지 확인하세요.

## 사용 방법

1. Excel을 실행하면 리본 메뉴에 "Excel To JSON" 탭이 나타납니다.
2. 해당 탭의 버튼을 사용하여 데이터를 JSON 또는 YAML로 변환할 수 있습니다.

## 시스템 요구사항

- Windows 7 이상
- Excel 2013 이상
- .NET Framework 4.8.1
- VSTO 런타임 4.0

## 수동 설치 방법

자동 설치가 작동하지 않을 경우:

1. Release 폴더의 모든 파일을 %APPDATA%\Microsoft\AddIns\ExcelToYamlAddin 폴더에 복사
2. Excel에서 파일 > 옵션 > 추가 기능 > COM 추가 기능 관리 > 찾아보기
3. %APPDATA%\Microsoft\AddIns\ExcelToYamlAddin\ExcelToYamlAddin.vsto 파일 선택
'@

$readmeFile = Join-Path $outputFolder "README.txt"
$readmeContent | Out-File -FilePath $readmeFile -Encoding utf8

# 배포 파일 압축
Write-Host "패키지 압축 중..." -ForegroundColor Green
if(Test-Path $zipFile) {
    Remove-Item -Path $zipFile -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($outputFolder, $zipFile)

Write-Host "패키징 완료! 배포 파일 생성됨: $zipFile" -ForegroundColor Green
Write-Host "이 파일을 팀원들과 공유하세요. 설치하려면 압축을 풀고 setup.bat를 실행하면 됩니다." -ForegroundColor Yellow 