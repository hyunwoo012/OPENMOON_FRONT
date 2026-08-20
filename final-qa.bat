@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo OPENMOON AI - FINAL QA
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [FAIL] .venv does not exist.
    echo Run setup.bat first.
    exit /b 1
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [FAIL] Could not activate .venv.
    exit /b 1
)

echo [1/7] Python environment
python --version
if errorlevel 1 goto :fail

echo.
echo [2/7] Database migration / initialization
python -m backend.scripts.init_db
if errorlevel 1 goto :fail

echo.
echo [3/7] Python syntax checks
python -m py_compile backend\app\models.py
if errorlevel 1 goto :fail
python -m py_compile backend\app\schemas.py
if errorlevel 1 goto :fail
python -m py_compile backend\app\routers\mails.py
if errorlevel 1 goto :fail
python -m py_compile backend\app\routers\quotations.py
if errorlevel 1 goto :fail
python -m py_compile backend\app\services\llm_service.py
if errorlevel 1 goto :fail
python -m py_compile backend\app\services\quotation_service.py
if errorlevel 1 goto :fail
python -m py_compile backend\app\services\smtp_service.py
if errorlevel 1 goto :fail

echo.
echo [4/7] Regression tests
python -m pytest ^
  backend\tests\test_attachment_vision.py ^
  backend\tests\test_customer_pdf_policy.py ^
  backend\tests\test_customer_pdf_no_excel.py ^
  backend\tests\test_send_to_self_mode.py ^
  backend\tests\test_quotation_email_update.py ^
  backend\tests\test_customer_pdf_send.py ^
  backend\tests\test_phase4d_customer_output.py ^
  backend\tests\test_mail_soft_delete.py ^
  -q
if errorlevel 1 goto :fail

echo.
echo [5/7] Frontend production build
pushd frontend
call npm run build
if errorlevel 1 (
    popd
    goto :fail
)
popd

if not exist "frontend\dist\index.html" (
    echo [FAIL] frontend\dist\index.html is missing.
    goto :fail
)

if not exist "frontend\dist\assets" (
    echo [FAIL] frontend\dist\assets is missing.
    goto :fail
)

echo.
echo [6/7] Required project files
for %%F in (
    "config\product_catalog.json"
    "backend\data\templates\quotation_template.xlsx"
) do (
    if not exist %%F (
        echo [WARN] Missing: %%F
    ) else (
        echo [OK] %%F
    )
)

echo.
echo [7/7] Microsoft Excel COM availability
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$excel=$null; try { $excel=New-Object -ComObject Excel.Application; Write-Host '[OK] Excel COM available'; exit 0 } catch { Write-Host '[WARN] Excel COM unavailable - customer PDF/native Excel QA must be done on another PC'; Write-Host $_.Exception.Message; exit 0 } finally { if ($null -ne $excel) { try { $excel.Quit() } catch {} }; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }"

echo.
echo ============================================================
echo [PASS] Automated final QA completed.
echo ============================================================
echo.
echo Manual QA is still required:
echo   docs\FINAL_QA_CHECKLIST.md
echo.
echo If Excel COM showed WARN, test quotation PDF and live send
echo on a Windows PC with Microsoft Excel installed.
echo.
exit /b 0

:fail
echo.
echo ============================================================
echo [FAIL] FINAL QA stopped because a check failed.
echo ============================================================
exit /b 1
