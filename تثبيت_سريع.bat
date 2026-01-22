@echo off
chcp 65001 >nul
echo.
echo ╔════════════════════════════════════════════════════════╗
echo ║         📦 تثبيت Add-in للمراجعة اللغوية            ║
echo ╚════════════════════════════════════════════════════════╝
echo.
echo هذا البرنامج سيقوم بتثبيت جميع المتطلبات تلقائياً
echo.
echo ════════════════════════════════════════════════════════
echo.

cd /d "%~dp0"

echo [1/3] 📥 تثبيت المكتبات المطلوبة...
echo.
call npm install
if %errorlevel% neq 0 (
    echo.
    echo ❌ خطأ: فشل تثبيت المكتبات
    echo    تأكد من تثبيت Node.js من: https://nodejs.org
    echo.
    pause
    exit /b 1
)

echo.
echo [2/3] 🔧 تثبيت أداة الشهادات...
echo.
call npm install -g office-addin-dev-certs
if %errorlevel% neq 0 (
    echo.
    echo ⚠️  تحذير: فشل تثبيت أداة الشهادات
    echo.
)

echo.
echo [3/3] 🔐 إنشاء الشهادة الموثوقة...
echo.
call office-addin-dev-certs install
if %errorlevel% neq 0 (
    echo.
    echo ⚠️  تحذير: فشل إنشاء الشهادة
    echo.
)

echo.
echo ════════════════════════════════════════════════════════
echo.
echo ✅ تم التثبيت بنجاح!
echo.
echo 📋 الخطوات التالية:
echo    1. اضغط مرتين على: تشغيل_التطبيق.bat
echo    2. افتح Microsoft Word
echo    3. اذهب إلى: إدراج → وظائفي الإضافية
echo    4. اختر: Shared Folder
echo    5. اختر التطبيق من القائمة
echo.
echo 🔑 لا تنسى الحصول على مفتاح Gemini API من:
echo    https://makersuite.google.com/app/apikey
echo.
echo ════════════════════════════════════════════════════════
echo.
pause
