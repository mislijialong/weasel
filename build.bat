@echo off

setlocal

if not exist env.bat copy env.bat.template env.bat

if exist env.bat call env.bat

if not defined WEASEL_ROOT set WEASEL_ROOT=%CD%

set "VSWHERE_EXE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
set VS_INSTALL_VERSION=
set VS_MAJOR=
if exist "%VSWHERE_EXE%" (
  for /f "usebackq delims=" %%i in (`"%VSWHERE_EXE%" -latest -products * -property installationVersion 2^>nul`) do (
    set VS_INSTALL_VERSION=%%i
  )
)
if defined VS_INSTALL_VERSION (
  for /f "tokens=1 delims=." %%i in ("%VS_INSTALL_VERSION%") do set VS_MAJOR=%%i
)

if not defined VERSION_MAJOR set VERSION_MAJOR=0
if not defined VERSION_MINOR set VERSION_MINOR=17
if not defined VERSION_PATCH set VERSION_PATCH=4

if not defined WEASEL_VERSION set WEASEL_VERSION=%VERSION_MAJOR%.%VERSION_MINOR%.%VERSION_PATCH%
if not defined WEASEL_BUILD set WEASEL_BUILD=0

rem use numeric build version for release build
set PRODUCT_VERSION=%WEASEL_VERSION%.%WEASEL_BUILD%
rem for non-release build, try to use git commit hash as product build version
if not defined RELEASE_BUILD (
  rem check if git is installed and available, then get the short commit id of head
  git --version >nul 2>&1
  if not errorlevel 1 (
    for /f "delims=" %%i in ('git tag --sort=-creatordate ^| findstr /r "%WEASEL_VERSION%"') do (
      set LAST_TAG=%%i
      goto found_tag
    )
    :found_tag
    for /f "delims=" %%i in ('git rev-list %LAST_TAG%..HEAD --count') do (
      set WEASEL_BUILD=%%i
    )
    rem get short commmit id of head
    for /F %%i in ('git rev-parse --short HEAD') do (set PRODUCT_VERSION=%WEASEL_VERSION%.%WEASEL_BUILD%.%%i)
  )
)

rem FILE_VERSION is always 4 numbers; same as PRODUCT_VERSION in release build
if not defined FILE_VERSION set FILE_VERSION=%WEASEL_VERSION%.%WEASEL_BUILD%

echo PRODUCT_VERSION=%PRODUCT_VERSION%
echo WEASEL_VERSION=%WEASEL_VERSION%
echo WEASEL_BUILD=%WEASEL_BUILD%
echo WEASEL_ROOT=%WEASEL_ROOT%
echo WEASEL_BUNDLED_RECIPES=%WEASEL_BUNDLED_RECIPES%
echo.

if defined GITHUB_ENV (
	setlocal enabledelayedexpansion
	echo git_ref_name=%PRODUCT_VERSION%>>!GITHUB_ENV!
)

if defined BOOST_ROOT (
  if exist "%BOOST_ROOT%\boost" goto boost_found
)
echo Error: Boost not found! Please set BOOST_ROOT in env.bat.
exit /b 1

:boost_found
echo BOOST_ROOT=%BOOST_ROOT%
echo.

if not defined BJAM_TOOLSET (
  rem the number actually means platform toolset, not %VisualStudioVersion%
  set BJAM_TOOLSET=msvc-14.2
)

if not defined PLATFORM_TOOLSET (
  set PLATFORM_TOOLSET=v142
)

set CMAKE_GENERATOR_SANITIZED=
if defined CMAKE_GENERATOR call set "CMAKE_GENERATOR_SANITIZED=%%CMAKE_GENERATOR:"=%%"
if defined CMAKE_GENERATOR (
  if defined CMAKE_GENERATOR_SANITIZED (
    set "CMAKE_GENERATOR=%CMAKE_GENERATOR_SANITIZED%"
  ) else (
    set CMAKE_GENERATOR=
  )
)

if not defined CMAKE_GENERATOR if defined MSBUILD_PATH call :resolve_cmake_generator_from_msbuild_path

if not defined CMAKE_GENERATOR (
  if "%VS_MAJOR%" == "16" set CMAKE_GENERATOR=Visual Studio 16 2019
  if "%VS_MAJOR%" == "17" set CMAKE_GENERATOR=Visual Studio 17 2022
  if "%VS_MAJOR%" == "18" set CMAKE_GENERATOR=Visual Studio 18 2026
)

if defined VS_INSTALL_VERSION echo VS_INSTALL_VERSION=%VS_INSTALL_VERSION%
if defined CMAKE_GENERATOR echo CMAKE_GENERATOR=%CMAKE_GENERATOR%
if defined PLATFORM_TOOLSET echo PLATFORM_TOOLSET=%PLATFORM_TOOLSET%
if defined BJAM_TOOLSET echo BJAM_TOOLSET=%BJAM_TOOLSET%
echo.

if defined DEVTOOLS_PATH set PATH=%DEVTOOLS_PATH%%PATH%

rem Setup include paths for ATL/WTL (similar to build_and_package.ps1)
set "project_include=%WEASEL_ROOT%\include"
set "project_wtl_include=%WEASEL_ROOT%\include\wtl"

rem Find Windows SDK include path using vswhere or manual search
set sdk_include=
set sdk_10_root=
for /f "delims=" %%i in ('"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.VCLibs.140.00.Desktop -find "Windows Kits\10\Include\10.0.*" 2^>nul') do (
    set sdk_10_root=%%i
)
if defined sdk_10_root (
    for /d %%i in ("%sdk_10_root%.*") do set sdk_include=%%i
)

rem Fallback: try common SDK locations
if not defined sdk_include (
    if exist "D:\Windows Kits\10\Include\10.0.26100.7705" set sdk_include=D:\Windows Kits\10\Include\10.0.26100.7705
    if exist "C:\Windows Kits\10\Include\10.0.26100.7705" set sdk_include=C:\Windows Kits\10\Include\10.0.26100.7705
    if exist "D:\Windows Kits\10\Include\10.0.22621.0" set sdk_include=D:\Windows Kits\10\Include\10.0.22621.0
    if exist "C:\Windows Kits\10\Include\10.0.22621.0" set sdk_include=C:\Windows Kits\10\Include\10.0.22621.0
)

if defined sdk_include (
    echo Using SDK include: %sdk_include%
    set "atl_include=%sdk_include%\atl"
    set "um_include=%sdk_include%\um"
    set "shared_include=%sdk_include%\shared"
) else (
    echo Warning: Windows SDK not found, ATL headers may be missing
    set atl_include=
    set um_include=
    set shared_include=
)

rem Build INCLUDE path
set "build_include=%project_include%;%project_wtl_include%"
if defined atl_include set "build_include=%build_include%;%atl_include%"
if defined um_include set "build_include=%build_include%;%um_include%"
if defined shared_include set "build_include=%build_include%;%shared_include%"
if defined INCLUDE set "build_include=%build_include%;%INCLUDE%"
set INCLUDE=%build_include%

set build_config=Release
set build_option=/t:Build
set build_boost=0
set boost_build_variant=release
set build_data=0
set build_opencc=0
set build_rime=0
set rime_build_variant=release
set build_weasel=0
set build_installer=0
set build_arm64=0

rem parse the command line options
:parse_cmdline_options
  if "%1" == "" goto end_parsing_cmdline_options
  if "%1" == "debug" (
    set build_config=Debug
    set boost_build_variant=debug
    set rime_build_variant=debug
  )
  if "%1" == "release" (
    set build_config=Release
    set boost_build_variant=release
    set rime_build_variant=release
  )
  if "%1" == "rebuild" set build_option=/t:Rebuild
  if "%1" == "boost" set build_boost=1
  if "%1" == "data" set build_data=1
  if "%1" == "opencc" set build_opencc=1
  if "%1" == "rime" set build_rime=1
  if "%1" == "librime" set build_rime=1
  if "%1" == "weasel" set build_weasel=1
  if "%1" == "installer" set build_installer=1
  if "%1" == "arm64" set build_arm64=1
  if "%1" == "all" (
    set build_boost=1
    set build_data=1
    set build_opencc=1
    set build_rime=1
    set build_weasel=1
    set build_installer=1
    set build_arm64=1
  )
  shift
  goto parse_cmdline_options
:end_parsing_cmdline_options

if %build_weasel% == 0 (
if %build_boost% == 0 (
if %build_data% == 0 (
if %build_opencc% == 0 (
if %build_rime% == 0 (
  set build_weasel=1
)))))

rem quit WeaselServer.exe before building
cd /d %WEASEL_ROOT%
if exist output\weaselserver.exe (
  output\weaselserver.exe /q
)

rem build booost
if %build_boost% == 1 (
  call :build_boost
  if errorlevel 1 exit /b 1
  cd /d %WEASEL_ROOT%
)

rem -------------------------------------------------------------------------
rem build librime x64 and Win32
if %build_rime% == 1 (
  if not exist librime\build.bat (
    git submodule update --init --recursive
  )
  if defined LIBRIME_PREDICT_ROOT (
    call :sync_librime_predict
    if errorlevel 1 exit /b 1
  )

  rem install lua plugin before building librime
  set RIME_PLUGINS=hchunhui/librime-lua
  pushd %WEASEL_ROOT%\librime
  call action-install-plugins-windows.bat
  popd
  if errorlevel 1 (
    echo Warning: failed to install rime plugins, continuing...
  )

  cd %WEASEL_ROOT%\librime
  rem clean cache before building
  for %%a in ( build dist lib ^
    deps\glog\build ^
    deps\googletest\build ^
    deps\leveldb\build ^
    deps\marisa-trie\build ^
    deps\opencc\build ^
    deps\yaml-cpp\build ) do (
      if exist %%a rd /s /q %%a
  )

  rem build x64 librime
  set ARCH=x64
  call :build_librime_platform x64 %WEASEL_ROOT%\lib64 %WEASEL_ROOT%\output
  rem build Win32 librime
  set ARCH=Win32
  call :build_librime_platform Win32 %WEASEL_ROOT%\lib %WEASEL_ROOT%\output\Win32
  rem clean the modified file
  rem git checkout .
  rem git submodule foreach git checkout .
)

rem -------------------------------------------------------------------------
if %build_weasel% == 1 (
  if not exist output\data\essay.txt (
    set build_data=1
  )
  if not exist output\data\opencc\TSCharacters.ocd* (
    set build_opencc=1
  )
)
if %build_data% == 1 call :build_data
if %build_opencc% == 1 call :build_opencc_data

if %build_weasel% == 0 goto end

cd /d %WEASEL_ROOT%

set WEASEL_PROJECT_PROPERTIES=BOOST_ROOT^
  PLATFORM_TOOLSET^
  VERSION_MAJOR^
  VERSION_MINOR^
  VERSION_PATCH^
  PRODUCT_VERSION^
  FILE_VERSION

cscript.exe render.js weasel.props %WEASEL_PROJECT_PROPERTIES%

del msbuild*.log

if defined SDKVER set build_sdk_option=/p:WindowsTargetPlatformVersion=%SDKVER%
if not defined SDKVER set build_sdk_option=

if %build_arm64% == 1 (

  msbuild.exe weasel.sln %build_option% /p:Configuration=%build_config% /p:Platform="ARM" /fl6 %build_sdk_option%
  if errorlevel 1 goto error
  msbuild.exe weasel.sln %build_option% /p:Configuration=%build_config% /p:Platform="ARM64" /fl5 %build_sdk_option%
  if errorlevel 1 goto error
)

msbuild.exe weasel.sln %build_option% /p:Configuration=%build_config% /p:Platform="x64" /fl2 %build_sdk_option%
if errorlevel 1 goto error
msbuild.exe weasel.sln %build_option% /p:Configuration=%build_config% /p:Platform="Win32" /fl1 %build_sdk_option%
if errorlevel 1 goto error

if %build_arm64% == 1 (
  pushd arm64x_wrapper
  call build.bat
  if errorlevel 1 goto error
  popd

  copy arm64x_wrapper\weaselARM64X.dll output
  if errorlevel 1 goto error
)

if %build_installer% == 1 (
  "%ProgramFiles(x86)%"\NSIS\Bin\makensis.exe ^
  /DWEASEL_VERSION=%WEASEL_VERSION% ^
  /DWEASEL_BUILD=%WEASEL_BUILD% ^
  /DPRODUCT_VERSION=%PRODUCT_VERSION% ^
  output\install.nsi
  if errorlevel 1 goto error
)

goto end

rem -------------------------------------------------------------------------
rem build boost
:build_boost
  set BJAM_OPTIONS_COMMON=-j%NUMBER_OF_PROCESSORS%^
    --with-filesystem^
    --with-json^
    --with-locale^
    --with-regex^
    --with-serialization^
    --with-system^
    --with-thread^
    define=BOOST_USE_WINAPI_VERSION=0x0603^
    toolset=%BJAM_TOOLSET%^
    link=static^
    runtime-link=static^
    --build-type=complete
  
  set BJAM_OPTIONS_X86=%BJAM_OPTIONS_COMMON%^
    architecture=x86^
    address-model=32
  
  set BJAM_OPTIONS_X64=%BJAM_OPTIONS_COMMON%^
    architecture=x86^
    address-model=64
  
  set BJAM_OPTIONS_ARM32=%BJAM_OPTIONS_COMMON%^
    define=BOOST_USE_WINAPI_VERSION=0x0A00^
    architecture=arm^
    address-model=32
  
  set BJAM_OPTIONS_ARM64=%BJAM_OPTIONS_COMMON%^
    define=BOOST_USE_WINAPI_VERSION=0x0A00^
    architecture=arm^
    address-model=64
  
  cd /d %BOOST_ROOT%
  if not exist b2.exe call bootstrap.bat
  if errorlevel 1 goto error
  b2 %BJAM_OPTIONS_X86% stage %BOOST_COMPILED_LIBS%
  if errorlevel 1 goto error
  b2 %BJAM_OPTIONS_X64% stage %BOOST_COMPILED_LIBS%
  if errorlevel 1 goto error
  
  if %build_arm64% == 1 (
    b2 %BJAM_OPTIONS_ARM32% stage %BOOST_COMPILED_LIBS%
    if errorlevel 1 goto error
    b2 %BJAM_OPTIONS_ARM64% stage %BOOST_COMPILED_LIBS%
    if errorlevel 1 goto error
  )
  exit /b

rem ---------------------------------------------------------------------------
:sync_librime_predict
  if not defined LIBRIME_PREDICT_ROOT exit /b 0
  if not exist "%LIBRIME_PREDICT_ROOT%\CMakeLists.txt" (
    echo Error: LIBRIME_PREDICT_ROOT is invalid: %LIBRIME_PREDICT_ROOT%
    exit /b 1
  )
  set "predict_dst=%WEASEL_ROOT%\librime\plugins\predict"
  if not exist "%predict_dst%" mkdir "%predict_dst%"
  echo Syncing librime-predict from %LIBRIME_PREDICT_ROOT%
  robocopy "%LIBRIME_PREDICT_ROOT%" "%predict_dst%" /E /R:2 /W:1 /NP /NFL /NDL /NJH /NJS /XD .git .github .claude >nul
  if errorlevel 8 (
    echo Error: failed to sync librime-predict.
    exit /b 1
  )
  exit /b 0

rem ---------------------------------------------------------------------------
:resolve_cmake_generator_from_msbuild_path
  setlocal EnableDelayedExpansion
  set "msbuild_path_local=%MSBUILD_PATH%"
  echo(!msbuild_path_local! | findstr /I /C:"2019" >nul
  if not errorlevel 1 (
    endlocal & set CMAKE_GENERATOR=Visual Studio 16 2019
    exit /b 0
  )
  echo(!msbuild_path_local! | findstr /I /C:"2022" >nul
  if not errorlevel 1 (
    endlocal & set CMAKE_GENERATOR=Visual Studio 17 2022
    exit /b 0
  )
  echo(!msbuild_path_local! | findstr /I /C:"2026" >nul
  if not errorlevel 1 (
    endlocal & set CMAKE_GENERATOR=Visual Studio 18 2026
    exit /b 0
  )
  endlocal
  exit /b 0

rem ---------------------------------------------------------------------------
:build_data
  copy %WEASEL_ROOT%\LICENSE.txt output\
  copy %WEASEL_ROOT%\README.md output\README.txt
  copy %WEASEL_ROOT%\plum\rime-install.bat output\
  copy %WEASEL_ROOT%\librime\share\opencc\*.* output\data\opencc\
  copy %WEASEL_ROOT%\third_party\webview2\pkg\build\native\x64\WebView2Loader.dll output\

  set plum_dir=plum
  set rime_dir=output/data
  set WSLENV=plum_dir:rime_dir
  bash plum/rime-install %WEASEL_BUNDLED_RECIPES%
  if errorlevel 1 goto error
  exit /b

rem ---------------------------------------------------------------------------
:build_opencc_data
  if not exist %WEASEL_ROOT%\librime\share\opencc\TSCharacters.ocd2 (
    cd %WEASEL_ROOT%\librime
    call build.bat deps %rime_build_variant%
    if errorlevel 1 goto error
  )
  cd %WEASEL_ROOT%
  if not exist output\data\opencc mkdir output\data\opencc
  copy %WEASEL_ROOT%\librime\share\opencc\*.* output\data\opencc\
  if errorlevel 1 goto error
  exit /b

rem ---------------------------------------------------------------------------
rem %1 : ARCH
rem %2 : push | pop , push to backup when pop to restore
:stash_build
  pushd %WEASEL_ROOT%\librime
  for %%a in ( build dist lib ^
    deps\glog\build ^
    deps\googletest\build ^
    deps\leveldb\build ^
    deps\marisa-trie\build ^
    deps\opencc\build ^
    deps\yaml-cpp\build ) do (
    if "%2"=="push" (
      if exist %%a  move %%a %%a_%1 
    )
    if "%2"=="pop" (
      if exist %%a_%1  move %%a_%1 %%a 
    )
  )
  popd
  exit /b

rem ---------------------------------------------------------------------------
rem %1 : ARCH
rem %2 : target_path of rime.lib, base %WEASEL_ROOT% or abs path
rem %3 : target_path of rime.dll, base %WEASEL_ROOT% or abs path
:build_librime_platform
  rem restore backuped %1 build
  call :stash_build %1 pop

  rem clear stale CMake caches so generator/toolset switches don't fail
  if exist %WEASEL_ROOT%\librime\build\CMakeCache.txt del /f /q %WEASEL_ROOT%\librime\build\CMakeCache.txt
  if exist %WEASEL_ROOT%\librime\build\CMakeFiles rd /s /q %WEASEL_ROOT%\librime\build\CMakeFiles
  for %%d in (deps\glog deps\googletest deps\leveldb deps\marisa-trie deps\opencc deps\yaml-cpp) do (
    if exist %WEASEL_ROOT%\librime\%%d\build\CMakeCache.txt del /f /q %WEASEL_ROOT%\librime\%%d\build\CMakeCache.txt
    if exist %WEASEL_ROOT%\librime\%%d\build\CMakeFiles rd /s /q %WEASEL_ROOT%\librime\%%d\build\CMakeFiles
  )

  cd %WEASEL_ROOT%\librime
  if not exist env.bat (
    copy %WEASEL_ROOT%\env.bat env.bat
  )
  if not exist lib\opencc.lib (
    call build.bat deps %rime_build_variant%
    if errorlevel 1 (
      call :stash_build %1 push
      goto error
    )
  )
  call build.bat %rime_build_variant%
  if errorlevel 1 (
    call :stash_build %1 push
    goto error
  )

  cd %WEASEL_ROOT%\librime
  call :stash_build %1 push

  copy /Y %WEASEL_ROOT%\librime\dist_%1\include\rime_*.h %WEASEL_ROOT%\include\
  if errorlevel 1 goto error
  copy /Y %WEASEL_ROOT%\librime\dist_%1\lib\rime.lib %2\
  if errorlevel 1 goto error
  copy /Y %WEASEL_ROOT%\librime\dist_%1\lib\rime.dll %3\
  if errorlevel 1 goto error

  exit /b
rem ---------------------------------------------------------------------------

:error

cd %WEASEL_ROOT%
echo error building weasel...
exit /b 1

:end
cd %WEASEL_ROOT%
