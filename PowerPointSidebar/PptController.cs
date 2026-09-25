using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PowerPointSidebar;

/// <summary>
/// 通过 COM 后期绑定（不依赖 Office PIA）操控 PowerPoint。
///
/// 设计原则：**绝不主动启动 PowerPoint**。
///   - Office 是单实例 COM 服务器，若 PowerPoint 已在运行，Activator.CreateInstance
///     会直接挂到现有实例上；若没运行则会把它启动起来 —— 我们不希望发生后者。
///   - 因此先用进程名探测 POWERPNT 是否存在，只在已运行时才去挂载。
///   - PowerPoint 关闭后 COM 引用会失效，捕获异常并自动解绑，下次轮询时重试。
/// </summary>
public sealed class PptController : IDisposable
{
    private object? _app;   // PowerPoint.Application COM 对象
    private bool _disposed;

    // -------------------------------------------------------------------
    // 官方枚举 PpSlideShowPointerType（值务必对齐，见 learn.microsoft.com）
    //
    //   ★ 这里踩过一个坑：橡皮是 5，不是 3！
    //     3 = ppSlideShowPointerAlwaysHidden（指针永久隐藏）
    //     写成 3 的话，点「橡皮」实际是把指针藏起来，看不到任何效果。
    //   ★ 旁证：刚起放映时读到的 PointerType 是 4 = ppSlideShowPointerAutoArrow，
    //     与本表一致，说明此表可信。
    // -------------------------------------------------------------------
    private const int PpPointerNone         = 0;
    private const int PpPointerArrow        = 1;
    private const int PpPointerPen          = 2;
    private const int PpPointerAlwaysHidden = 3;
    private const int PpPointerAutoArrow    = 4;
    private const int PpPointerEraser       = 5;

    /// <summary>最近一次失败原因（UI 可用它做提示）。</summary>
    public string? LastError { get; private set; }

    /// <summary>当前是否已挂上 PowerPoint。</summary>
    public bool IsAttached => _app is not null;

    // -------------------------------------------------------------------
    // 挂载 / 解绑
    // -------------------------------------------------------------------

    /// <summary>尝试挂载到已在运行的 PowerPoint；未运行则返回 false（不会启动它）。</summary>
    public bool EnsureAttached()
    {
        if (_disposed) return false;
        if (_app is not null) return true;

        if (!IsPowerPointProcessRunning())
        {
            LastError = "PowerPoint 未运行。";
            return false;
        }

        try
        {
            var t = Type.GetTypeFromProgID("PowerPoint.Application");
            if (t is null)
            {
                LastError = "找不到 PowerPoint.Application ProgID，请确认已安装 PowerPoint。";
                return false;
            }
            _app = Activator.CreateInstance(t);
            if (_app is null)
            {
                LastError = "无法挂载到 PowerPoint 实例。";
                return false;
            }
            LastError = null;
            return true;
        }
        catch (Exception ex)
        {
            LastError = "挂载 PowerPoint 失败：" + ex.Message;
            _app = null;
            return false;
        }
    }

    private void Detach()
    {
        if (_app is null) return;
        try { Marshal.ReleaseComObject(_app); } catch { /* 忽略 */ }
        _app = null;
    }

    private static bool IsPowerPointProcessRunning()
    {
        try
        {
            // PowerPoint.exe 的进程名是 POWERPNT
            var procs = Process.GetProcessesByName("POWERPNT");
            return procs.Length > 0;
        }
        catch
        {
            return false;
        }
    }

    // -------------------------------------------------------------------
    // 状态
    // -------------------------------------------------------------------

    /// <summary>当前是否正在放映幻灯片（至少存在一个 SlideShowWindow）。</summary>
    public bool IsSlideshowRunning
    {
        get
        {
            if (!EnsureAttached()) return false;
            try
            {
                var windows = GetProp(_app!, "SlideShowWindows");
                if (windows is null) return false;
                int count = Convert.ToInt32(GetProp(windows, "Count") ?? 0);
                return count > 0;
            }
            catch
            {
                // PowerPoint 可能刚被关闭，引用失效
                Detach();
                return false;
            }
        }
    }

    // -------------------------------------------------------------------
    // 动作
    // -------------------------------------------------------------------

    public void NextSlide() => ActOnView(view => InvokeMethod(view, "Next"));

    public void PreviousSlide() => ActOnView(view => InvokeMethod(view, "Previous"));

    /// <summary>笔：切到 PowerPoint 自带笔，并使用指定颜色。</summary>
    public void UsePen(Color color) => ActOnView(view =>
    {
        SetPointerColor(view, color);
        SetProp(view, "PointerType", PpPointerPen);
    });

    /// <summary>
    /// 荧光笔：PowerPoint 的 COM 接口只暴露「一种笔」，没有真正的半透明高亮工具，
    /// 也无法通过 COM 调整笔头粗细（官方对象模型无 PenWidth/MarkerSize）。
    /// 所以这里退而求其次：用同一种笔，只是颜色取用户设定的「荧光色」。
    /// </summary>
    public void UseHighlighter(Color color) => ActOnView(view =>
    {
        SetPointerColor(view, color);
        SetProp(view, "PointerType", PpPointerPen);
    });

    /// <summary>
    /// 橡皮：切到 PowerPoint 自带橡皮指针。
    ///
    /// ★ 实测（本机 Office 16 x64）：
    ///   `PointerType = 5`（ppSlideShowPointerEraser）**只在幻灯片上已经存在墨迹时才生效**。
    ///   页面空白、一笔都没画时，PowerPoint 会**静默忽略**这次赋值（读回退到 4 = AutoArrow）。
    ///   0/1/2/3/4 则任何时候都能写入。
    ///   所以这里写入后回读确认，返回 false 表示没切成功（通常是本页还没笔迹）。
    /// </summary>
    public bool UseEraser()
    {
        bool ok = true;
        ActOnView(view =>
        {
            SetProp(view, "PointerType", PpPointerEraser);
            try
            {
                ok = Convert.ToInt32(GetProp(view, "PointerType") ?? -1) == PpPointerEraser;
            }
            catch
            {
                ok = true;   // 回读不了就当成功，别误报
            }
        });
        return ok;
    }

    /// <summary>清空当前页所有笔迹。</summary>
    public void ClearCurrentSlideInk() => ActOnView(view => InvokeMethod(view, "EraseDrawing"));

    /// <summary>退出当前幻灯片放映，回到普通编辑视图。</summary>
    public void ExitSlideshow() => ActOnView(view => InvokeMethod(view, "Exit"));

    /// <summary>
    /// 设置放映笔的颜色。
    ///
    /// ★ 官方文档里 SlideShowView.PointerColor 是**只读的 ColorFormat 对象**，
    ///   正确写法是给它的子属性 .RGB 赋值，而不是把整数直接塞给 PointerColor。
    ///   但没有 PIA 时后期绑定的实际行为不一定与文档一致，
    ///   所以这里先走 ColorFormat.RGB，失败再退回直接赋值。
    /// </summary>
    private static void SetPointerColor(object view, Color color)
    {
        int bgr = ToBgr(color);

        try
        {
            var cf = GetProp(view, "PointerColor");
            if (cf is not null)
            {
                SetProp(cf, "RGB", bgr);
                return;
            }
        }
        catch { /* 落到下面的直接赋值 */ }

        SetProp(view, "PointerColor", bgr);
    }

    /// <summary>
    /// 读回当前的指针状态，供自检/排错用。
    /// pointerType 为 ppSlideShowPointerType 的值；pointerRgb 取不到时为 null。
    /// </summary>
    public bool TryGetPointerState(out int pointerType, out int? pointerRgb)
    {
        pointerType = -1;
        pointerRgb = null;
        if (!EnsureAttached()) return false;

        object? view = null;
        try
        {
            view = GetActiveViewOrThrow();
            pointerType = Convert.ToInt32(GetProp(view, "PointerType") ?? -1);

            var cf = GetProp(view, "PointerColor");
            if (cf is not null)
            {
                try { pointerRgb = Convert.ToInt32(GetProp(cf, "RGB") ?? 0); }
                catch { pointerRgb = null; }
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (view is not null)
            {
                try { Marshal.ReleaseComObject(view); } catch { /* 忽略 */ }
            }
        }
    }

    // -------------------------------------------------------------------
    // 窗口
    // -------------------------------------------------------------------

    /// <summary>
    /// PowerPoint 主窗口句柄；不可用时返回 IntPtr.Zero。
    ///
    /// 注意：PowerPoint 只开着「开始屏幕」而没打开任何演示文稿时，
    /// Application.HWND 会返回 0。这时退回用 Win32 枚举 POWERPNT 进程的顶层窗口，
    /// 取面积最大的那一个（即真正的主框架窗口）。
    /// </summary>
    public IntPtr GetMainWindowHandle()
    {
        if (_app is not null)
        {
            try
            {
                var hwnd = GetProp(_app, "HWND");
                if (hwnd is not null)
                {
                    int v = Convert.ToInt32(hwnd);
                    if (v != 0) return new IntPtr(v);
                }
            }
            catch { /* 落到 Win32 兜底 */ }
        }
        return FindPowerPointWindowByProcess();
    }

    // ---- Win32 窗口枚举兜底 ----

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }

    private const uint GW_OWNER = 4;

    private static IntPtr FindPowerPointWindowByProcess()
    {
        try
        {
            var procs = Process.GetProcessesByName("POWERPNT");
            if (procs.Length == 0) return IntPtr.Zero;

            var pids = new System.Collections.Generic.HashSet<uint>();
            foreach (var p in procs)
            {
                try { pids.Add((uint)p.Id); } catch { /* 忽略 */ }
            }

            IntPtr best = IntPtr.Zero;
            long bestArea = -1;

            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                if (GetWindow(hWnd, GW_OWNER) != IntPtr.Zero) return true; // 跳过有 owner 的弹窗

                GetWindowThreadProcessId(hWnd, out uint pid);
                if (!pids.Contains(pid)) return true;

                if (!GetWindowRect(hWnd, out var rc)) return true;
                long area = (long)(rc.Right - rc.Left) * (rc.Bottom - rc.Top);
                if (area > bestArea)
                {
                    bestArea = area;
                    best = hWnd;
                }
                return true;
            }, IntPtr.Zero);

            return best;
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    /// <summary>
    /// 把前台焦点还给 PowerPoint。
    /// 单击本侧栏按钮的瞬间焦点会转移到侧栏，动作完成后必须交回，
    /// 否则放映期间键盘翻页 / 翻页笔会失灵。
    /// </summary>
    public void ActivateMainWindow()
    {
        var hwnd = GetMainWindowHandle();
        if (hwnd != IntPtr.Zero)
        {
            try { SetForegroundWindow(hwnd); } catch { /* 忽略 */ }
        }
    }

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    // -------------------------------------------------------------------
    // 内部工具
    // -------------------------------------------------------------------

    private void ActOnView(Action<object> action)
    {
        if (!EnsureAttached())
            throw new InvalidOperationException(
                LastError ?? "未连接到 PowerPoint。请先打开 PowerPoint 并开始放映（F5）。");

        object? view = null;
        try
        {
            view = GetActiveViewOrThrow();
            action(view);
            LastError = null;
        }
        catch (Exception)
        {
            Detach();          // 引用可能已失效，下次重挂
            throw;
        }
        finally
        {
            if (view is not null)
            {
                try { Marshal.ReleaseComObject(view); } catch { /* 忽略 */ }
            }
        }
    }

    private object GetActiveViewOrThrow()
    {
        var windows = GetProp(_app!, "SlideShowWindows")
            ?? throw new InvalidOperationException("当前没有幻灯片放映窗口。请先在 PowerPoint 中按 F5 开始放映。");

        int count = Convert.ToInt32(GetProp(windows, "Count") ?? 0);
        if (count == 0)
            throw new InvalidOperationException("当前没有幻灯片放映窗口。请先在 PowerPoint 中按 F5 开始放映。");

        var ssw = GetCollectionItem(windows, 1)
            ?? throw new InvalidOperationException("无法获取 SlideShowWindow。");

        var view = GetProp(ssw, "View")
            ?? throw new InvalidOperationException("无法获取 SlideShowView。");

        return view;
    }

    /// <summary>
    /// 从 Office 集合里取第 index 个元素。
    ///
    /// ★ 这里有个实测踩过的坑：Office 各个集合的 Item 成员，
    ///   有的必须按「方法」调用、有的按「属性」调用，不能一律当属性。
    ///   例如 SlideShowWindows.Item 用 GetProperty 会直接抛
    ///   DISP_E_MEMBERNOTFOUND (0x80020003)，必须用 InvokeMethod。
    ///   所以先按方法试，失败再按属性试。
    /// </summary>
    private static object? GetCollectionItem(object collection, int index)
    {
        try
        {
            var byMethod = InvokeMethod(collection, "Item", index);
            if (byMethod is not null) return byMethod;
        }
        catch { /* 落到属性方式 */ }

        return GetProp(collection, "Item", index);
    }

    private static object? GetProp(object target, string name, params object[] args)
    {
        return target.GetType().InvokeMember(
            name,
            BindingFlags.GetProperty | BindingFlags.GetField,
            null,
            target,
            args.Length == 0 ? null : args);
    }

    private static void SetProp(object target, string name, object value)
    {
        target.GetType().InvokeMember(
            name, BindingFlags.SetProperty, null, target, new object[] { value });
    }

    private static object? InvokeMethod(object target, string name, params object[] args)
    {
        return target.GetType().InvokeMember(
            name, BindingFlags.InvokeMethod, null, target, args.Length == 0 ? null : args);
    }

    private static int ToBgr(Color c) => (c.B << 16) | (c.G << 8) | c.R;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Detach();
    }
}