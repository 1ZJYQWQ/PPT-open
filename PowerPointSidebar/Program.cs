using System;
using System.Text;
using System.Windows.Forms;

namespace PowerPointSidebar;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        // ★ 高 DPI 必须在**任何分支之前**设置。
        //   否则诊断分支看到的是 DPI 虚拟化之后的坐标，
        //   和真实运行时的坐标对不上（本项目就踩过：自检报 1280 宽，实际是 2560）。
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);

        // 无头自检模式：不创建窗口，只检查 COM 挂载情况，把结果打到 stdout
        if (args is { Length: > 0 } && args[0] is "--selftest" or "-selftest")
        {
            return SelfTest.Run();
        }

        // 几何诊断：真的建一次窗体、应用贴边，把屏幕/DPI/最终尺寸打出来
        if (args is { Length: > 0 } && args[0] is "--where" or "-where")
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ApplyPanelOverrides(args);
            using var probe = new MainForm();
            // 支持「--where --width 38 --height 77」直接对数，不必先改设置文件
            probe.ApplySizeOverrides(DoubleArg(args, "--width"), DoubleArg(args, "--height"));
            _ = probe.Handle;              // 强制创建句柄，让 DPI/尺寸落到实处
            probe.DockForDiagnostics();
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine(probe.DescribeGeometry());
            return 0;
        }

        if (args is { Length: > 0 } && args[0] is "--help" or "-h" or "-?")
        {
            Console.WriteLine("PowerPoint 演示侧边栏");
            Console.WriteLine();
            Console.WriteLine("  PowerPointSidebar.exe              启动（缩进托盘，放映开始时自动弹出）");
            Console.WriteLine("  PowerPointSidebar.exe --show       启动并立即显示侧栏（不等到放映）");
            Console.WriteLine("  PowerPointSidebar.exe --selftest   无头自检：检查与 PowerPoint 的 COM 连接");
            Console.WriteLine("  PowerPointSidebar.exe --where      打印屏幕/DPI/侧栏实际尺寸（排查大小问题）");
            Console.WriteLine("  PowerPointSidebar.exe --kill       结束已在运行的侧栏进程后退出");
            Console.WriteLine("  PowerPointSidebar.exe --help       显示本帮助");
            Console.WriteLine();
            Console.WriteLine("尺寸参数（单位是【毫米】—— 不是像素。同一串像素在不同 ppi 的");
            Console.WriteLine("屏幕上大小能差 3 倍，只有毫米才能保证\"看起来一样大\"）：");
            Console.WriteLine("  --width  N    侧栏宽度，毫米。0 = 用默认值（38 mm），默认 0");
            Console.WriteLine("  --height N    侧栏高度，毫米。0 = 贴合内容（5 个按钮 ≈ 77 mm），默认 0");
            Console.WriteLine("  --edge  L/R   初始贴在屏幕哪条边，LEFT 或 RIGHT，默认 RIGHT");
            Console.WriteLine();
            Console.WriteLine("  例：PowerPointSidebar.exe --width 45 --height 90 --edge LEFT");
            Console.WriteLine();
            Console.WriteLine("  「启动侧栏.cmd」顶部已经预置了这三个变量，改那里就行。");
            Console.WriteLine("  命令行指定的值只在本次运行生效，不会写进设置文件。");
            Console.WriteLine();
            Console.WriteLine("屏幕物理尺寸（决定毫米换算）：");
            Console.WriteLine("  程序会自动读显示器 EDID 拿真实物理尺寸。读不到时（某些显示器/投影");
            Console.WriteLine("  不填这个字段）可以用下面的参数手动告诉它：");
            Console.WriteLine("  --panel N        屏幕对角英寸数，按 16:9 反算。例：--panel 86");
            Console.WriteLine("  --mm-per-dip X   直接指定每逻辑像素多少毫米。例：--mm-per-dip 0.494");
            Console.WriteLine("  跑 --where 可以看到它读到了什么、算出的 ppi 和每控件毫米数。");
            Console.WriteLine();
            Console.WriteLine("笔色设置：放映中点「笔」右侧的小色块即可换色，");
            Console.WriteLine("         颜色保存在 %APPDATA%\\PowerPointSidebar\\settings.txt");
            return 0;
        }

        if (args is { Length: > 0 } && args[0] is "--kill" or "-kill")
        {
            int killed = 0;
            foreach (var p in System.Diagnostics.Process.GetProcessesByName("PowerPointSidebar"))
            {
                try
                {
                    if (p.Id == Environment.ProcessId) continue;   // 不杀自己
                    p.Kill();
                    killed++;
                }
                catch { /* 忽略 */ }
            }
            Console.WriteLine($"已结束 {killed} 个侧栏进程。");
            return 0;
        }

        // ---------------------------------------------------------------
        // 解析 --show / --width N / --height N / --edge L|R
        //   ★ 宽度和高度单位是【毫米】；-1 表示「本次没指定，沿用设置文件」。
        //     用毫米而不是像素：同一串像素在不同 ppi 的屏幕上大小差 3 倍以上。
        // ---------------------------------------------------------------
        bool forceShow = false;
        double cliWidth = -1, cliHeight = -1;
        string cliEdge = "";

        ApplyPanelOverrides(args);   // --panel / --mm-per-dip 要在建窗体前生效

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a is "--show" or "-show") { forceShow = true; continue; }

            if ((a is "--width" or "-width") && i + 1 < args.Length
                && double.TryParse(args[i + 1], System.Globalization.NumberStyles.Float,
                                   System.Globalization.CultureInfo.InvariantCulture, out double wv))
            {
                cliWidth = Math.Max(0, wv);   // 负数当 0（自动）处理
                i++;
                continue;
            }
            if ((a is "--height" or "-height") && i + 1 < args.Length
                && double.TryParse(args[i + 1], System.Globalization.NumberStyles.Float,
                                   System.Globalization.CultureInfo.InvariantCulture, out double hv))
            {
                cliHeight = Math.Max(0, hv);
                i++;
                continue;
            }
            if ((a is "--edge" or "-edge") && i + 1 < args.Length)
            {
                cliEdge = args[++i];
                continue;
            }
        }

        // ---------------------------------------------------------------
        // 单实例保护
        //   重复双击启动脚本不要开出第二个侧栏。
        //   已经有一个在跑时，就发信号让那个把侧栏显示出来，本次直接退出。
        //
        //   刻意**不**在 .cmd 里用 tasklist|find 判断 —— 那些外部命令会被
        //   Git Bash 的同名工具顶掉（实测 `find` 被顶成 GNU find，直接报错），
        //   判断就会失效。放在程序里用命名内核对象判断才可靠。
        // ---------------------------------------------------------------
        const string MutexName = "PowerPointSidebar.SingleInstance.v1";
        const string ShowEventName = "PowerPointSidebar.ShowRequest.v1";

        bool isFirstInstance;
        using var singleInstance = new Mutex(initiallyOwned: true, MutexName, out isFirstInstance);

        if (!isFirstInstance)
        {
            try
            {
                using var req = EventWaitHandle.OpenExisting(ShowEventName);
                req.Set();                       // 通知已有实例：把侧栏调出来
            }
            catch { /* 已有实例还没开始监听，忽略 */ }
            return 0;
        }

        // 高 DPI 已在 Main 开头设置过（见上方注释）
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // 全局异常兜底，避免侧栏闪退时整个进程无声消失
        Application.ThreadException += (_, e) =>
            MessageBox.Show(
                "侧边栏发生未处理异常：\n\n" + e.Exception.Message,
                "PowerPoint 侧边栏", MessageBoxButtons.OK, MessageBoxIcon.Error);

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            var ex = e.ExceptionObject as Exception;
            MessageBox.Show(
                "侧边栏发生未捕获异常：\n\n" + (ex?.Message ?? "(未知错误)"),
                "PowerPoint 侧边栏", MessageBoxButtons.OK, MessageBoxIcon.Error);
        };

        // ★ 用 ApplicationContext 承载，不用 Application.Run(form)：
        //   后者会强制把窗体 Show() 一次，启动瞬间会闪一个窗口；
        //   而侧栏平时应该安静地待在托盘里等放映开始。
        var form = new MainForm();
        // ★ 命令行尺寸必须在 Start() 之前套用 —— Start() 里就会 ApplyDock() 定尺寸
        form.ApplySizeOverrides(cliWidth, cliHeight);
        if (cliEdge.Length > 0) form.ApplyEdgeOverride(cliEdge);

        var ctx = new ApplicationContext(form);
        form.Start();
        if (forceShow) form.ShowNow();

        // 监听"显示"请求：重复启动本程序时把已在运行的侧栏调出来
        var showRequest = new EventWaitHandle(false, EventResetMode.AutoReset, ShowEventName);
        var waiter = new Thread(() =>
        {
            while (true)
            {
                showRequest.WaitOne();
                try { form.BeginInvoke(new Action(() => form.ShowNow())); }
                catch { /* 窗体已关闭 */ }
            }
        })
        { IsBackground = true, Name = "show-request-listener" };
        waiter.Start();

        Application.Run(ctx);

        showRequest.Dispose();
        return 0;
    }

    /// <summary>
    /// 从参数数组里取「--name N」的 N（毫米，负数会被夹成 0）。
    /// 取不到就返回 -1，表示"本次没指定"。
    /// </summary>
    private static double DoubleArg(string[] args, string name)
    {
        for (int i = 0; i < args.Length - 1; i++)
            if (args[i].Equals(name, StringComparison.OrdinalIgnoreCase)
                && double.TryParse(args[i + 1], System.Globalization.NumberStyles.Float,
                                   System.Globalization.CultureInfo.InvariantCulture, out double v))
                return Math.Max(0, v);
        return -1;
    }

    /// <summary>
    /// --panel N（屏幕对角英寸数） / --mm-per-dip X（每逻辑像素毫米数）。
    /// 用来在 EDID 读不到物理尺寸时手动兜底。
    /// </summary>
    private static void ApplyPanelOverrides(string[] args)
    {
        double panel = DoubleArg(args, "--panel");
        double mmpd = DoubleArg(args, "--mm-per-dip");
        if (panel > 0) ScreenMetrics.ForcedDiagonalInch = panel;
        if (mmpd > 0) ScreenMetrics.ForcedMmPerDip = mmpd;
    }
}

/// <summary>
/// 无头自检：验证 COM 挂载链路是否可用。不创建任何窗口，不启动 PowerPoint。
/// 用法：PowerPointSidebar.exe --selftest
/// </summary>
internal static class SelfTest
{
    public static int Run()
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== PowerPoint 侧边栏 自检 ===");
        sb.AppendLine();

        int exitCode = 0;

        // ---- 屏幕 / DPI：用来核对「自动侧栏宽度」公式的单位是否一致 ----
        try
        {
            var scr = Screen.PrimaryScreen!;
            float dpi;
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) dpi = g.DpiX;

            sb.AppendLine($"0) 主屏 Bounds      : {scr.Bounds.Width} x {scr.Bounds.Height}");
            sb.AppendLine($"   主屏 WorkingArea : {scr.WorkingArea.Width} x {scr.WorkingArea.Height}");
            sb.AppendLine($"   系统 DPI         : {dpi:0.##}  (约 {dpi / 96.0 * 100:0}% 缩放)");
            sb.AppendLine($"   => 由此算出的自动宽度 : {AutoWidthProbe(scr.WorkingArea.Width)}");
            sb.AppendLine();
        }
        catch (Exception ex)
        {
            sb.AppendLine("0) 屏幕信息读取失败: " + ex.Message);
            sb.AppendLine();
        }

        try
        {
            var procs = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
            sb.AppendLine($"1) PowerPoint 进程 (POWERPNT) 数量 : {procs.Length}");

            using var ctrl = new PptController();

            bool attached = ctrl.EnsureAttached();
            sb.AppendLine($"2) COM 挂载 PowerPoint.Application  : {(attached ? "成功" : "未挂载")}");

            if (!attached)
            {
                sb.AppendLine($"   原因 : {ctrl.LastError}");
                sb.AppendLine();
                sb.AppendLine("=> 结论：当前无法连接。请先打开 PowerPoint 再重试。");
                exitCode = 2;
            }
            else
            {
                bool show = ctrl.IsSlideshowRunning;
                sb.AppendLine($"3) 是否正在放映幻灯片           : {(show ? "是" : "否")}");

                var hwnd = ctrl.GetMainWindowHandle();
                sb.AppendLine($"4) PowerPoint 主窗口句柄         : 0x{hwnd.ToInt64():X}");

                sb.AppendLine();
                if (show)
                {
                    sb.AppendLine("=> 结论：一切就绪。侧栏的六个按钮现在都可以用。");
                }
                else
                {
                    sb.AppendLine("=> 结论：已连上 PowerPoint，但当前没有在放映。");
                    sb.AppendLine("        请在 PowerPoint 里按 F5 开始放映，按钮才会生效。");
                }
            }
        }
        catch (Exception ex)
        {
            sb.AppendLine();
            sb.AppendLine("!! 自检过程中抛出异常：");
            sb.AppendLine("   " + ex.GetType().Name + ": " + ex.Message);
            if (ex.InnerException is not null)
                sb.AppendLine("   inner: " + ex.InnerException.Message);
            exitCode = 1;
        }

        sb.AppendLine();
        sb.AppendLine("=== 自检结束 ===");

        var text = sb.ToString();
        Console.Out.Write(text);
        Console.Out.Flush();

        // WinExe 从资源管理器启动时没有控制台，额外写一份到 exe 同目录便于排查
        try
        {
            var dir = AppContext.BaseDirectory;
            File.WriteAllText(Path.Combine(dir, "selftest.log"), text, new UTF8Encoding(false));
        }
        catch { /* 写不了就算了，stdout 已有输出 */ }

        return exitCode;
    }

    /// <summary>
    /// 复刻 MainForm.AutoWidthFor 的算法，供自检打印用。
    /// 注意：这里依赖 Screen.WorkingArea 的单位与窗体 SetBounds 的单位一致，
    /// 自检正是用来核对这一点的。
    /// </summary>
    private static int AutoWidthProbe(int workWidth)
    {
        int w = (int)Math.Round(workWidth * 0.06);
        return Math.Clamp(w, 62, 140);
    }
}