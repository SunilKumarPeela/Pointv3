@echo off
setlocal
if not exist build mkdir build
rc /nologo /fo build\point.res src\point.rc
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /O2 /MT ^
  src\point_fetcher.cpp build\point.res /Fe:build\PointFetcher.exe ^
  /link /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib shell32.lib advapi32.lib winhttp.lib
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /O2 /MT ^
  src\point_core.cpp src\point_compliance.cpp src\point_excel_import.cpp src\point_win32.cpp ^
  build\point.res /Fe:build\Point.exe ^
  /link /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib oleaut32.lib ^
  advapi32.lib crypt32.lib bcrypt.lib
if errorlevel 1 exit /b 1
if not exist build\Inbox mkdir build\Inbox
if not exist build\Inbox\Users.csv copy /Y sample\*.csv build\Inbox\ >nul
if not exist build\point-security.conf copy /Y point-security.conf build\ >nul
echo Built build\Point.exe and build\PointFetcher.exe
