@echo off
setlocal

rem ============================================================
rem  PPT Sidebar launcher
rem
rem  改 WIDTH / HEIGHT 即可（也就是这个文件存在的意义）。
rem  单位都是毫米，所以在任何屏幕缩放下都保持相同的物理尺寸。
rem
rem    WIDTH   0  = 交给侧栏自动算，默认约 38 mm。
rem                 N  = 强制 N mm 宽              (允许 38..80)
rem                 比 38 mm 还窄会把按钮文字切掉，所以会被拒绝。
rem
rem    HEIGHT  0  = 交给侧栏自动算，默认"贴合内容"：
rem                 刚好装下 5 个按钮，不再顶天立地。
rem                 N  = 强制 N mm 高              (允许 77..300)
rem
rem    注意    HEIGHT 不会低于内容所需的自然高度。再小会把最后一个按钮裁掉，
rem            侧栏没有滚动条。
rem
rem    EDGE    LEFT 或 RIGHT —— 启动时贴哪条边。
rem            也可以拖到另一边来切换。
rem ============================================================

set WIDTH=0
set HEIGHT=0
set EDGE=RIGHT

rem ---- 输出目录（改了项目目录就改这里） ----
set TFM=net10.0-windows

set EXE=%~dp0PowerPointSidebar\bin\Release\%TFM%\PowerPointSidebar.exe

if exist "%EXE%" goto killold

echo.
echo [x] 找不到可执行文件：
echo     %EXE%
echo.
echo 先编译：
echo     dotnet build "%~dp0PowerPointSidebar\PowerPointSidebar.csproj" -c Release
echo.
goto end

:killold
rem 先关掉旧实例。不关的话改了 WIDTH / HEIGHT 看起来"没反应"：
rem 程序的单实例保护只会把老窗口拉出来，新参数根本读不到。
taskkill /IM PowerPointSidebar.exe /F >nul 2>&1

echo.
echo 启动 PPT 侧栏 ...
echo.
echo   WIDTH  = %WIDTH% mm    (0 = 自动)
echo   HEIGHT = %HEIGHT% mm    (0 = 贴合内容)
echo   EDGE   = %EDGE%
echo.
echo   * 启动后缩进系统托盘 —— 没有窗口，没有任务栏图标。
echo   * 在 PowerPoint 里按 F5 开始放映：侧栏自动弹出并贴边。
echo   * 放映结束（按 Esc / 放完）：自动收托盘。
echo   * 拖标题栏可以移动 —— 拖到屏幕 40 px 内自动吸附贴边。
echo   * 托盘图标：双击 = 显示/隐藏；右键 = 退出 / 颜色 / 尺寸。
echo.

rem 只有非 0 才传。0 表示沿用上次（默认自动），这样在侧栏里手拖的宽度
rem 不会被脚本每次都覆盖回去。
set ARGS=
if not "%WIDTH%"=="0"  set ARGS=%ARGS% --width %WIDTH%
if not "%HEIGHT%"=="0" set ARGS=%ARGS% --height %HEIGHT%
if not "%EDGE%"==""    set ARGS=%ARGS% --edge %EDGE%

start "" "%EXE%" %ARGS%

:end
endlocal