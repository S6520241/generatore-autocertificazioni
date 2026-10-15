@echo off
rem ======================================================================
rem  Crea "Autocertificazioni.exe" a partire da genera_documenti.py
rem  Da eseguire UNA VOLTA SOLA sul TUO computer (serve Python installato).
rem  In questa cartella devono esserci:
rem     genera_documenti.py
rem     template_autocertificazione_it.docx
rem     template_autocertificazione_en.docx
rem     ISTRUZIONI.txt
rem ======================================================================
cd /d "%~dp0"

set PY=py
where py >nul 2>nul || set PY=python
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo Python non trovato. Installalo da https://www.python.org e riprova.
    pause
    exit /b 1
)

for %%F in (genera_documenti.py template_autocertificazione_it.docx template_autocertificazione_en.docx) do (
    if not exist "%%F" (
        echo Manca il file %%F in questa cartella.
        pause
        exit /b 1
    )
)

echo.
echo [1/3] Installazione librerie...
%PY% -m pip install --upgrade pdfplumber python-docx pyinstaller
if errorlevel 1 goto errore

echo.
echo [2/3] Creazione del programma (1-2 minuti)...
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name Autocertificazioni ^
    --add-data "template_autocertificazione_it.docx;." ^
    --add-data "template_autocertificazione_en.docx;." ^
    genera_documenti.py
if errorlevel 1 goto errore

echo.
echo [3/3] Preparazione della cartella da consegnare...
if exist "Da_consegnare" rmdir /s /q "Da_consegnare"
mkdir "Da_consegnare\Autocertificazioni"
copy /y "dist\Autocertificazioni.exe" "Da_consegnare\Autocertificazioni\" >nul
if exist "ISTRUZIONI.txt" copy /y "ISTRUZIONI.txt" "Da_consegnare\Autocertificazioni\" >nul

echo.
echo FATTO. Copia la cartella "Da_consegnare\Autocertificazioni" sul computer dell'utente.
explorer "Da_consegnare"
pause
exit /b 0

:errore
echo.
echo Si e' verificato un errore: leggi i messaggi qui sopra.
pause
exit /b 1
