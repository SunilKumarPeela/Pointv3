@echo off
setlocal
if not exist build-debug mkdir build-debug
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /Od /Zi /RTC1 ^
  src\point_core.cpp src\point_compliance.cpp src\point_excel_import.cpp src\point_win32.cpp ^
  /Fe:build-debug\Point-Debug.exe ^
  /link /DEBUG /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib ole32.lib oleaut32.lib ^
  advapi32.lib crypt32.lib bcrypt.lib
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /utf-8 /EHsc /W4 /WX /sdl /guard:cf ^
  /DUNICODE /D_UNICODE /Od /Zi /RTC1 ^
  src\point_fetcher.cpp /Fe:build-debug\PointFetcher-Debug.exe ^
  /link /DEBUG /DYNAMICBASE /NXCOMPAT /GUARD:CF ^
  user32.lib gdi32.lib comctl32.lib shell32.lib advapi32.lib winhttp.lib
if errorlevel 1 exit /b 1
echo Built Point and Point Fetcher debug executables with runtime checks.
