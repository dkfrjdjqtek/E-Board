@echo off
setlocal enabledelayedexpansion

where git >nul 2>nul || (echo [ERROR] Git이 설치되어 있지 않습니다. & pause & exit /b 1)

for /f "delims=" %%i in ('git rev-parse --show-toplevel 2^>nul') do set REPO=%%i
if not defined REPO (
  echo [ERROR] 여기는 Git 저장소가 아닙니다.
  pause & exit /b 1
)
pushd "%REPO%"

for /f "delims=" %%i in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%i

git add -A :/

set HASCHANGES=
for /f "delims=" %%i in ('git status --porcelain') do set HASCHANGES=1

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm\""') do set NOW=%%i
set MSG=Daily update: %NOW%

if defined HASCHANGES (
  git commit -m "%MSG%"
) else (
  echo [INFO] 커밋할 변경 없음. 푸시만 진행합니다.
)

git rev-parse --abbrev-ref --symbolic-full-name @{u} >nul 2>nul
if errorlevel 1 (
  git push -u origin %BRANCH% --force
) else (
  git push origin %BRANCH% --force
)

if errorlevel 1 (
  echo [ERROR] 푸시 실패. 세부 로그 확인 필요.
  popd & pause & exit /b 1
)

git log --oneline -1
echo [OK] Pushed to origin/%BRANCH%.
popd
pause