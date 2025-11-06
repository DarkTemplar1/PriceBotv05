@echo off
setlocal enabledelayedexpansion

REM ==== 1) Parametry ====
set APP_NAME=PriceBot
set ENTRY=main.py
set VENV=.venv
set DIST_DIR=dist
set BUILD_DIR=build

REM ==== 2) Czystość katalogów (opcjonalnie) ====
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"

REM ==== 3) Wirtualne środowisko ====
if not exist "%VENV%\Scripts\python.exe" (
  py -3 -m venv "%VENV%"
)
call "%VENV%\Scripts\activate"

REM ==== 4) Instalacja zależności ====
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

REM ==== 5) Budowanie EXE (onefile, bez konsoli) ====
REM  - --hidden-import: czasem pomaga wykryć tkinter/lxml w nietypowych przypadkach
pyinstaller ^
  --noconfirm ^
  --name "%APP_NAME%" ^
  --onefile ^
  --windowed ^
  --hidden-import=tkinter ^
  --hidden-import=lxml ^
  "%ENTRY%"

REM ==== 6) Kopiowanie dodatkowych plików (jeśli chcesz mieć wzorce/README) ====
REM copy README.txt "%DIST_DIR%\README.txt" >nul 2>&1

REM ==== 7) Podsumowanie ====
if exist "%DIST_DIR%\%APP_NAME%.exe" (
  echo.
  echo [OK] Zbudowano: %DIST_DIR%\%APP_NAME%.exe
  echo Uruchom i w GUI w sekcji "Miejsce tworzenia plikow i folderow" kliknij "Przygotowanie Aplikacji".
) else (
  echo.
  echo [ERR] Nie znaleziono pliku EXE w %DIST_DIR%.
  exit /b 1
)

exit /b 0
