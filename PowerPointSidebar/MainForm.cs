using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

#pragma warning disable WFO1000 // WinForms 分析器：自定义控件属性无需 designer 序列化

namespace PowerPointSidebar;

/// <summary>
/// PowerPoint 演示侧边栏。
///
/// 核心行为（按用户要求）：
///   · 放映开始 → 自动弹出，并贴到屏幕左/右边缘
///   · 放映结束 → 自动收进系统托盘
///   · 全程 **不出现任务栏条目**，也不会把任务栏带出来
///   · 全程 **不抢 PowerPoint 的键盘焦点**（否则放映时翻页笔/方向键会失灵）
///   · 始终盖在全屏放映之上
///
/// 实现要点：
///   WS_EX_NOACTIVATE  —— 点击侧栏不激活它，焦点始终留在 PowerPoint
///   WS_EX_TOOLWINDOW  —— 不出现在任务栏 / Alt+Tab
///   定时 SetWindowPos(HWND_TOPMOST) —— PowerPoint 全屏放映本身也是置顶窗口，
///                        会被它压下去，所以要周期性重新断言置顶
///   因为 NOACTIVATE 会让模态 MessageBox 失效，所有提示改为侧栏内嵌提示条
/// </summary>
public class MainForm : Form
{
    // ================= 配色（浅色卡片） =================
    private static readonly Color BgColor       = Color.FromArgb(247, 248, 251);       // 卡片底：近白浅灰
    private static readonly Color TitleBarColor = Color.FromArgb(238, 240, 245);       // 标题栏略深
    private static readonly Color SeparatorCol  = Color.FromArgb(222, 225, 232);       // 分隔线
    private static readonly Color TextColor     = Color.FromArgb(48, 50, 58);          // 深灰：图标/文字
    private static readonly Color DimTextColor  = Color.FromArgb(140, 143, 152);
    private static readonly Color OverlayColor  = Color.FromArgb(246, 250, 251, 253);  // 浅色覆盖层
    private static readonly Color CardBorder    = Color.FromArgb(214, 218, 226);       // 卡片描边

    private static readonly Color AccentNav    = Color.FromArgb(45, 127, 249);         // 蓝
    private static readonly Color AccentEraser = Color.FromArgb(120, 126, 138);        // 灰
    private static readonly Color AccentExit   = Color.FromArgb(229, 72, 77);          // 红

    // 调色板预设色（5 列 × 3 行）。第一行是常用亮色，第二行是深色/冷色，第三行是中性与荧光色。
    private static readonly Color[] PaletteColors =
    {
        Color.FromArgb(255, 60, 60),    // 红（默认笔色）
        Color.FromArgb(255, 140, 30),   // 橙
        Color.FromArgb(255, 215, 0),    // 金黄
        Color.FromArgb(255, 235, 0),    // 亮黄 / 荧光（默认荧光笔色）
        Color.FromArgb(80, 220, 100),   // 绿
        Color.FromArgb(60, 200, 200),   // 青
        Color.FromArgb(60, 140, 255),   // 蓝
        Color.FromArgb(30, 70, 180),    // 深蓝
        Color.FromArgb(160, 90, 220),   // 紫
        Color.FromArgb(255, 105, 170),  // 粉
        Color.FromArgb(150, 160, 175),  // 灰
        Color.FromArgb(20, 20, 20),     // 近黑
        Color.FromArgb(255, 255, 255),  // 白
        Color.FromArgb(140, 90, 50),    // 棕
        Color.FromArgb(170, 20, 40),    // 深红
    };

    // ================= 尺寸：全部以【毫米】为目标 =================
    //
    // ★★ 为什么不用像素：触控目标必须按毫米定（手指是物理尺寸固定的东西），
    //   而"1 个像素 = 几毫米"完全取决于屏幕 ppi：
    //     65" 4K 大屏（100% 缩放）→ 0.37 ~ 0.49 mm/像素
    //     13" 笔记本（200% 缩放）→ 0.13 mm/像素        ← 差 3 倍多
    //   同一份"11 毫米"的按钮，在这些屏上必须用完全不同的像素数才能一样大。
    //
    //   换算链：毫米  →  Mm2Px()  →  物理像素
    //           除数来自 ScreenMetrics（读显示器 EDID 得到真实物理尺寸，
    //           读不到会兜底到名义 96 DPI，并允许命令行 --panel / --mm-per-dip 覆盖）。

    /// <summary>标题栏高度。同时是拖动把手，10 mm 够手指/细笔抓。</summary>
    private const double MmTitleBar = 10;
    /// <summary>单个按钮高度。10 mm &gt; 触控绝对下限 9 mm。</summary>
    private const double MmButton = 10;
    /// <summary>右上角 ✕ 关闭按钮见方。</summary>
    private const double MmCloseBtn = 9;
    /// <summary>朝向幻灯片那一侧的宽度抓握区宽度。</summary>
    private const double MmResizeGrip = 9;
    /// <summary>贴边时离屏幕边缘的距离。放映时全屏独占、不触发边缘手势，故严丝合缝贴 0。</summary>
    private const double MmEdgeMargin = 0;
    /// <summary>默认侧栏宽度（毫米）。</summary>
    private const double MmWidth = 30;
    private const double MmWidthMin = 28;
    private const double MmWidthMax = 70;
    /// <summary>上下留白（毫米）。</summary>
    private const double MmPadding = 2;
    /// <summary>内侧圆角半径（毫米）。贴屏那侧保持直角，不破坏贴边观感。</summary>
    private const double MmCorner = 3;

    /// <summary>目标屏的物理尺寸信息（每逻辑像素多少毫米等）。ApplyDock 里刷新。</summary>
    private ScreenMetrics.Info _metrics = new();

    /// <summary>毫米 → 物理像素。</summary>
    private int Mm2Px(double mm) => (int)Math.Round(mm * _metrics.PhysicalPxPerMm);

    /// <summary>毫米 → 逻辑像素（给 MinimumSize 这类按逻辑像素算的场合用）。</summary>
    private int Mm2Dip(double mm) => (int)Math.Round(mm / Math.Max(1e-9, _metrics.MmPerDip));

    /// <summary>逻辑像素 → 物理像素。</summary>
    private int L(double logical) => (int)Math.Round(logical * DpiScale);

    private int TitleBarH   => Mm2Px(MmTitleBar);
    private int ResizeGripW => Mm2Px(MmResizeGrip);
    private int MinBtnH     => Mm2Px(MmButton);
    private int DockMargin  => Mm2Px(MmEdgeMargin);
    private int CornerR     => Mm2Px(MmCorner);

    // ================= Win32 =================
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

    private static readonly IntPtr HWND_TOPMOST = new(-1);
    private const uint SWP_NOSIZE     = 0x0001;
    private const uint SWP_NOMOVE     = 0x0002;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    // ================= 依赖 =================
    private readonly PptController _ppt = new();

    // ================= 设置（笔色等，持久化到 %APPDATA%） =================
    private readonly SettingsStore _settings = SettingsStore.Load();
    private Color _penColor;
    private SidebarButton? _penBtn;

    // ================= 托盘 =================
    private NotifyIcon? _tray;
    private Icon? _trayIcon;
    private ContextMenuStrip? _menu;

    // ================= 布局 =================
    private readonly List<Control> _stack = new();
    private bool _uiBuilt;
    private readonly List<SidebarButton> _navButtons = new();

    // ================= 停靠 =================
    /// <summary>
    /// 侧栏只能贴在屏幕【左边缘】或【右边缘】。
    ///
    /// ★ 刻意没有 Float（自由悬浮）状态 —— 按用户要求：可以拖拽，但松手后必须贴边，
    ///   不允许像以前那样停在屏幕中间。拖动只用来「换边」和「选垂直位置」。
    /// </summary>
    public enum DockSide { Left, Right }

    private DockSide _dock = DockSide.Right;

    /// <summary>当前生效的宽度（物理像素）。由 ApplyDock() 算出来。</summary>
    private int _dockWidth;

    /// <summary>
    /// 贴边时的垂直位置（物理像素）。null = 还没拖过，ApplyDock() 会垂直居中。
    /// 拖过一次之后就一直沿用用户选的位置。
    /// </summary>
    private int? _dockTop;

    /// <summary>命令行 --width 覆盖，单位**毫米**。&lt;0 未指定；0 = 强制自动；&gt;0 = 固定毫米数。</summary>
    private double _overrideWidthMm = -1;

    /// <summary>命令行 --height 覆盖，单位**毫米**。&lt;0 未指定；0 = 强制自动；&gt;0 = 固定毫米数。</summary>
    private double _overrideHeightMm = -1;

    // ================= 状态 =================
    private bool _isInSlideshow;
    private bool _userHidden;                 // 用户从托盘手动隐藏（本次放映内不再自动弹）
    private readonly System.Windows.Forms.Timer _pollTimer;
    private readonly System.Windows.Forms.Timer _topMostTimer;

    // ================= 拖动 / 缩放 =================
    private bool _dragging;
    private Point _dragStartMouse;
    private Point _dragStartForm;
    private bool _resizing;
    private int _resizeStartWidth;
    private int _resizeStartMouseX;

    // ================= 悬停 =================
    private bool _hoverClose;
    private bool _hoverGrip;

    public MainForm()
    {
        Text = "PPT 侧栏";
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;          // ★ 不占任务栏
        TopMost = true;
        BackColor = BgColor;
        DoubleBuffered = true;
        // 注意：MinimumSize/MaximumSize 是**逻辑像素**且会被 DPI 缩放，
        // 只是外层保险；真正的裁剪由 ApplyDock 按毫米把关。
        MinimumSize = new Size(30, 120);
        MaximumSize = new Size(600, 4000);
        StartPosition = FormStartPosition.Manual;

        Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold);

        // 读上次保存的笔色（宽度/高度会在 ApplyDock 里按毫米重新算）
        _penColor = _settings.PenColor;

        BuildTray();

        // 轮询放映状态 —— 用来驱动「放映开始弹出 / 放映结束收托盘」
        _pollTimer = new System.Windows.Forms.Timer { Interval = 600 };
        _pollTimer.Tick += (_, _) => OnPollTick();
        _pollTimer.Start();

        // 周期性重新断言置顶，压过 PowerPoint 的全屏放映窗口
        _topMostTimer = new System.Windows.Forms.Timer { Interval = 700 };
        _topMostTimer.Tick += (_, _) => { if (Visible) ReassertTopMost(); };
        _topMostTimer.Start();

        FormClosing += OnFormClosing;
    }

    // ===================================================================
    // 关键：不抢焦点 + 不进任务栏
    //   WS_EX_NOACTIVATE (0x08000000)：点击侧栏不激活，焦点留在 PowerPoint
    //   WS_EX_TOOLWINDOW (0x00000080)：不进任务栏、不进 Alt+Tab
    // ===================================================================
    protected override bool ShowWithoutActivation => true;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
            cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
            return cp;
        }
    }

    private void ReassertTopMost()
    {
        if (!IsHandleCreated || !Visible) return;
        // SWP_NOACTIVATE：过程中也不夺取焦点
        SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }

    // ===================================================================
    // 放映状态轮询：驱动自动弹出 / 自动收纳
    // ===================================================================
    private void OnPollTick()
    {
        SyncSlideshowState();
    }

    /// <summary>
    /// 放映状态同步 —— **弹出 / 收托盘 的唯一入口**。
    ///
    /// ★ 踩过的坑：以前 Guard() 的 finally 里只把 _isInSlideshow 改了、调用 Invalidate()，
    ///   却没有真正执行「收托盘」；而轮询定时器一比对发现标志已经同步过，就判定"状态没变"直接返回。
    ///   结果：点侧栏上的「结束放映」退出后，侧栏**永远不收回去**（按 ESC 退出却能收，
    ///   因为那条路没走 Guard）。现在所有路径都走这一个入口，不会再漏。
    /// </summary>
    private void SyncSlideshowState()
    {
        bool now = _ppt.IsSlideshowRunning;

        if (now == _isInSlideshow)
        {
            if (now && Visible) ReassertTopMost();   // 放映中顺便维持置顶
            return;
        }

        _isInSlideshow = now;

        if (now)
        {
            // ── 放映开始：弹出 + 贴边
            _userHidden = false;
            ApplyDock();
            if (!Visible) Show();
            ReassertTopMost();
            Invalidate();
        }
        else
        {
            // ── 放映结束：收进托盘
            HideToTray();
        }
    }

    /// <summary>
    /// 收进托盘。刻意**不弹气泡通知** —— 侧栏自己消失就是最清楚的信号，
    /// 再弹一个系统通知纯属噪音（用户明确反馈过"不必要的提示太多"）。
    /// </summary>
    private void HideToTray()
    {
        if (!Visible) return;
        Hide();
    }

    // ===================================================================
    // 启动
    //   注意：用 ApplicationContext 承载本窗体，不用 Application.Run(form)，
    //   因为后者会强制 Show() 一次，导致启动瞬间闪一个窗口。
    //   这里由 Program 显式调用 Start() 来决定一开始显不显示。
    // ===================================================================
    public void Start()
    {
        if (_uiBuilt) return;
        _uiBuilt = true;

        _ = Handle;      // 强制创建窗口句柄（不显示），后面的 SetWindowPos 等才有得用
        BuildUi();

        // 启动时：如果 PowerPoint 已经在放映就直接弹出；否则先缩进托盘等放映
        if (_ppt.IsSlideshowRunning)
        {
            _isInSlideshow = true;
            ApplyDock();
            Show();
            ReassertTopMost();
        }
        // 不在放映时：什么都不显示，安静待在托盘里等 F5。
        // （以前这里会弹一个"已就绪"气泡，属于噪音，已去掉）
    }

    /// <summary>强制立即显示（供 --show 命令行参数使用，方便在没放映时也能看到侧栏）。</summary>
    public void ShowNow()
    {
        _userHidden = false;
        ApplyDock();
        Show();
        ReassertTopMost();
        Invalidate();
    }

    // ===================================================================
    // 托盘图标（替代任务栏图标，作为稳定的关闭入口）
    // ===================================================================
    private void BuildTray()
    {
        _trayIcon = CreateTrayIcon();

        _menu = new ContextMenuStrip { ShowImageMargin = false };
        _menu.Items.Add("显示 / 隐藏侧栏", null, (_, _) => ToggleVisibleByUser());
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add("贴到屏幕左边缘", null, (_, _) => SetDock(DockSide.Left));
        _menu.Items.Add("贴到屏幕右边缘", null, (_, _) => SetDock(DockSide.Right));
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add("宽度：恢复默认", null, (_, _) =>
        {
            _settings.WidthMm = 0;          // 0 = 用默认毫米数
            _overrideWidthMm = -1;          // ★ 同时撤掉命令行的固定值，否则"改了没反应"
            _settings.Save();
            ApplyDock();
            ShowIfNotUserHidden();
        });
        _menu.Items.Add("高度：贴合内容", null, (_, _) =>
        {
            _settings.HeightMm = 0;         // 0 = 按内容自然高度
            _overrideHeightMm = -1;
            _settings.Save();
            ApplyDock();
            ShowIfNotUserHidden();
        });
        _menu.Items.Add("垂直位置：重新居中", null, (_, _) =>
        {
            _dockTop = null;
            ApplyDock();
            ShowIfNotUserHidden();
        });
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add("笔的颜色…", null, (_, _) => OpenPalette());
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add("当前状态 / PID", null, (_, _) => ShowStatus());
        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add("退出侧栏（关闭本程序）", null, (_, _) => Close());

        _tray = new NotifyIcon
        {
            Icon = _trayIcon,
            Text = "PPT 侧栏",
            Visible = true,
            ContextMenuStrip = _menu,
        };
        _tray.DoubleClick += (_, _) => ToggleVisibleByUser();
    }

    /// <summary>换边（托盘菜单和标题栏黄点都走这里）。侧栏永远贴边，没有悬浮状态。</summary>
    private void SetDock(DockSide side)
    {
        _dock = side;
        ApplyDock();
        ShowIfNotUserHidden();
    }

    private void ToggleVisibleByUser()
    {
        if (Visible)
        {
            _userHidden = true;
            Hide();
        }
        else
        {
            ShowIfNotUserHidden(resetFlag: true);
        }
    }

    private void ShowIfNotUserHidden(bool resetFlag = false)
    {
        if (resetFlag) _userHidden = false;
        else if (_userHidden) return;   // 用户本次放映内手动藏起来了，就不打扰他

        ApplyDock();
        if (!Visible) Show();
        ReassertTopMost();
        Invalidate();
    }

    private static Icon CreateTrayIcon()
    {
        using var bmp = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            using var bg = new SolidBrush(Color.FromArgb(46, 46, 54));
            using var bgPen = new Pen(Color.FromArgb(120, 130, 150), 1.6f);
            var r = new Rectangle(3, 5, 26, 22);
            g.FillRectangle(bg, r);
            g.DrawRectangle(bgPen, r);

            // 一条红色的"笔迹"
            using var ink = new Pen(Color.FromArgb(240, 80, 80), 2.6f)
            {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
            };
            g.DrawLine(ink, 8, 21, 23, 10);

            // 右侧一条竖条，暗示"侧边栏"
            using var bar = new SolidBrush(Color.FromArgb(80, 160, 255));
            g.FillRectangle(bar, 24, 5, 5, 22);
        }

        var h = bmp.GetHicon();
        try
        {
            // Clone 一份，避免依赖临时 HICON
            using var tmp = Icon.FromHandle(h);
            return (Icon)tmp.Clone();
        }
        finally
        {
            DestroyIcon(h);
        }
    }

    // ===================================================================
    // 贴边
    // ===================================================================
    private Screen TargetScreen()
    {
        var hwnd = _ppt.GetMainWindowHandle();
        if (hwnd != IntPtr.Zero)
        {
            try { return Screen.FromHandle(hwnd); } catch { /* 退回主屏 */ }
        }
        return Screen.PrimaryScreen ?? Screen.AllScreens[0];
    }

    /// <summary>当前窗体的 DPI 缩放倍数（1.0 = 96 DPI）。</summary>
    private double DpiScale => (DeviceDpi > 0 ? DeviceDpi : 96) / 96.0;

    /// <summary>
    /// 内容的【自然高度】，毫米 —— 刚好装下标题栏和五个纯图标按钮（无分隔条）。
    ///   10 + 2×2 + 5×10 = 64 mm
    ///
    /// 下限就是它：比它矮会把最后一个按钮裁掉（本窗体没有滚动条）。
    /// 想更高可以在「启动侧栏.cmd」的 HEIGHT 里按**毫米**指定。
    /// </summary>
    private double NaturalHeightMm()
    {
        int btnCount = _navButtons.Count > 0 ? _navButtons.Count : 5;
        return MmTitleBar + MmPadding * 2 + btnCount * MmButton;
    }

    /// <summary>解析出要用的宽度（毫米）。命令行 &gt; 用户手拖的设置 &gt; 默认值。</summary>
    private double ResolveWidthMm() => _overrideWidthMm switch
    {
        0 => MmWidth,                                          // --width 0 = 强制回到默认
        > 0 => _overrideWidthMm,                               // --width N = 强制 N 毫米
        _ => _settings.WidthMm > 0 ? _settings.WidthMm : MmWidth,
    };

    /// <summary>解析出要用的高度（毫米）。命令行 &gt; 设置 &gt; 贴合内容。</summary>
    private double ResolveHeightMm(double naturalMm) => _overrideHeightMm switch
    {
        0 => naturalMm,
        > 0 => _overrideHeightMm,
        _ => _settings.HeightMm > 0 ? _settings.HeightMm : naturalMm,
    };

    private void ApplyDock()
    {
        var target = TargetScreen();
        var wa = target.WorkingArea;             // 物理像素

        // ★★ 先按目标屏刷新「每逻辑像素多少毫米」—— 下面所有换算都依赖它。
        //    用目标屏而不是窗体自己的 DPI：教室常见"笔记本 + 大屏"两块屏缩放不同，
        //    用窗体的 DPI 会在第一帧算错（以前就是这个 bug）。
        _metrics = ScreenMetrics.For(target, DeviceDpi);

        // ---------------- 宽度（毫米）----------------
        double mmW = Math.Clamp(ResolveWidthMm(), MmWidthMin, MmWidthMax);
        int w = Mm2Px(mmW);
        _dockWidth = w;

        // ---------------- 高度（毫米）----------------
        double naturalMm = NaturalHeightMm();
        double workMm = wa.Height / Math.Max(1e-9, _metrics.PhysicalPxPerMm);
        double maxMm = Math.Max(naturalMm, workMm - MmEdgeMargin * 2);
        double mmH = Math.Clamp(ResolveHeightMm(naturalMm), naturalMm, maxMm);
        int h = Mm2Px(mmH);

        // ---------------- 垂直位置 ----------------
        // 拖过就沿用用户选的位置，没拖过就垂直居中。
        // 居中很关键：大屏画面底边通常离地 0.8~0.9 m，
        // 贴顶会让侧栏跑到 1.8 m 以上，个子矮的老师够不着。
        int minTop = wa.Top + DockMargin;
        int maxTop = Math.Max(minTop, wa.Bottom - h - DockMargin);
        int top = _dockTop is int ft ? Math.Clamp(ft, minTop, maxTop) : wa.Top + (wa.Height - h) / 2;

        // 只能贴左右两边，没有悬浮状态
        if (_dock == DockSide.Left)
            SetBounds(wa.Left + DockMargin, top, w, h);
        else
            SetBounds(wa.Right - w - DockMargin, top, w, h);

        if (_uiBuilt) LayoutStack();
        ApplyRoundedRegion();
    }

    /// <summary>命令行覆盖（--width / --height，单位**毫米**；0 = 自动，负数 = 不改动）。</summary>
    public void ApplySizeOverrides(double widthMm, double heightMm)
    {
        if (widthMm >= 0) _overrideWidthMm = widthMm;
        if (heightMm >= 0) _overrideHeightMm = heightMm;
    }

    /// <summary>命令行 --edge：指定先贴哪条边。</summary>
    public void ApplyEdgeOverride(string edge)
    {
        if (edge.Equals("LEFT", StringComparison.OrdinalIgnoreCase)) _dock = DockSide.Left;
        else if (edge.Equals("RIGHT", StringComparison.OrdinalIgnoreCase)) _dock = DockSide.Right;
    }

    // ---- 供 --where 诊断用 ----

    /// <summary>应用一次贴边（不显示窗口），配合 DescribeGeometry 用于排查尺寸问题。</summary>
    public void DockForDiagnostics() => ApplyDock();

    /// <summary>打印屏幕 / DPI / 最终尺寸，用来核对宽度公式的单位是否一致。</summary>
    public string DescribeGeometry()
    {
        var scr = TargetScreen();
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("=== 侧栏几何诊断 ===");
        sb.AppendLine($"目标屏幕      : {scr.DeviceName}");
        sb.AppendLine($"  Bounds      : {scr.Bounds.Width} x {scr.Bounds.Height}");
        sb.AppendLine($"  WorkingArea : {scr.WorkingArea.Width} x {scr.WorkingArea.Height}");
        sb.AppendLine($"窗体 DPI      : {DeviceDpi}   (缩放 {DpiScale * 100:0}%)");
        sb.AppendLine();
        sb.AppendLine("--- 屏幕物理尺寸（决定「1 像素 = 几毫米」）---");
        sb.AppendLine($"来源          : {_metrics.Source}");
        sb.AppendLine($"该屏 DPI      : {_metrics.Dpi}");
        sb.AppendLine($"可视宽度      : {_metrics.PhysicalWidthMm:0.#} mm   分辨率 {_metrics.ResolutionWidth} px"
                      + $"   →  {_metrics.PhysicalWidthMm / Math.Max(1, _metrics.ResolutionWidth):0.####} mm / 物理像素");
        sb.AppendLine($"像素密度      : {_metrics.Ppi:0.#} ppi   对角约 {ScreenMetrics.DiagonalInch:0.#} 英寸");
        sb.AppendLine($"★ 每逻辑像素  : {_metrics.MmPerDip:0.####} mm"
                      + $"   (1 毫米 = {_metrics.DipPerMm:0.##} 逻辑像素 = {_metrics.PhysicalPxPerMm:0.##} 物理像素)");
        sb.AppendLine();
        sb.AppendLine("--- 尺寸目标（毫米 → 物理像素）---");
        sb.AppendLine($"宽度          : {ResolveWidthMm():0.#} mm"
                      + $"   [设置={_settings.WidthMm}  --width={(_overrideWidthMm < 0 ? "未指定" : _overrideWidthMm.ToString("0.#"))}]");
        sb.AppendLine($"高度          : {ResolveHeightMm(NaturalHeightMm()):0.#} mm"
                      + $"   [设置={_settings.HeightMm}  --height={(_overrideHeightMm < 0 ? "未指定" : _overrideHeightMm.ToString("0.#"))}]"
                      + $"   内容自然高 {NaturalHeightMm():0.#} mm");
        sb.AppendLine($"贴边          : {_dock}   把手在内侧{(GripOnInnerLeft ? "左" : "右")}边");
        sb.AppendLine();
        sb.AppendLine("--- 各控件（触控标准：≥9 mm 舒服，6 mm 绝对下限）---");
        void Mm2(string name, double mm)
        {
            int px = Mm2Px(mm);
            sb.AppendLine($"  {name,-10} {mm,5:0.#} mm  →  {px,4} px"
                          + (mm < 6 ? "   <<< 低于绝对下限" : mm < 9 ? "   (够用但偏小)" : ""));
        }
        Mm2("标题栏", MmTitleBar);
        Mm2("单个按钮", MmButton);
        Mm2("✕ 关闭", MmCloseBtn);
        Mm2("宽度抓握区", MmResizeGrip);
        Mm2("离屏边距", MmEdgeMargin);
        var cr = CloseButtonRect();
        sb.AppendLine($"  整条标题栏可拖 : x0..{cr.Left - Mm2Px(2)} 与 x{cr.Right + Mm2Px(2)}..{_dockWidth}"
                      + $"  (✕ 占 x{cr.Left}..{cr.Right})");

        sb.AppendLine();
        sb.AppendLine("--- 纯图标布局 ---");
        sb.AppendLine("  按钮只画图标（名称靠悬停 ToolTip），所以不存在文字切字问题。");
        sb.AppendLine($"  单按钮可用宽  : {_dockWidth} px");
        sb.AppendLine("  图标源图      : 128 px（4× 超采样，放大到 ~96 px 仍清晰）");
        sb.AppendLine();
        sb.AppendLine($"最终 Bounds   : X={Bounds.X} Y={Bounds.Y}  {Bounds.Width} x {Bounds.Height}  (物理像素)");
        sb.AppendLine($"换算成逻辑    : {Bounds.Width / DpiScale:0} x {Bounds.Height / DpiScale:0}");
        sb.AppendLine($"实际毫米      : {Bounds.Width / Math.Max(1e-9, _metrics.PhysicalPxPerMm):0.#} x "
                      + $"{Bounds.Height / Math.Max(1e-9, _metrics.PhysicalPxPerMm):0.#} mm");
        sb.AppendLine($"宽度占屏幕    : {Bounds.Width * 100.0 / Math.Max(1, scr.Bounds.Width):0.0}%");
        sb.AppendLine($"高度占屏幕    : {Bounds.Height * 100.0 / Math.Max(1, scr.Bounds.Height):0.0}%");
        return sb.ToString();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _pollTimer?.Stop();
            _topMostTimer?.Stop();
            if (_tray is not null) { _tray.Visible = false; _tray.Dispose(); _tray = null; }
            _trayIcon?.Dispose();
            _trayIcon = null;
            _menu?.Dispose();
            _menu = null;
            _ppt.Dispose();
        }
        base.Dispose(disposing);
    }

    // ===================================================================
    // 构建按钮
    // ===================================================================
    private void BuildUi()
    {
        SuspendLayout();
        _stack.Clear();
        _navButtons.Clear();
        _penBtn = null;

        // 纯图标布局：按钮上没有文字，名称靠悬停 ToolTip 给出。
        // ToolTip 不抢键盘焦点，放映全屏时也能正常弹出。
        _tip ??= new ToolTip
        {
            InitialDelay = 350,
            ReshowDelay = 150,
            AutoPopDelay = 4000,
            ShowAlways = true,
        };

        AddButton("上一页", IconFactory.Previous(TextColor), AccentNav, OnPrevClicked);
        AddButton("下一页", IconFactory.Next(TextColor), AccentNav, OnNextClicked);

        _penBtn = AddButton("笔", IconFactory.Pen(TextColor), _penColor, OnPenClicked);
        _penBtn.ChipColor = _penColor;
        _penBtn.OnChipClick = () => OpenPalette();

        // ★ 「荧光笔」按钮已删除：PowerPoint 的 COM 只暴露一种笔，那个按钮
        //   实质只是"把笔调成黄色"，而颜色本来就能用「笔」右侧色条改，完全冗余。
        AddButton("橡皮", IconFactory.Eraser(TextColor), AccentEraser, OnEraserClicked, OnClearInkClicked);

        // 结束放映用红色图标，和"停止"语义一致，也是唯一有破坏性的按钮
        AddButton("结束放映", IconFactory.Exit(AccentExit), AccentExit, OnExitClicked);

        BuildOverlay();
        LayoutStack();
        ResumeLayout(true);
    }

    private SidebarButton AddButton(string label, Bitmap icon, Color accent, Action onClick,
        Action? onDoubleClick = null)
    {
        var btn = new SidebarButton
        {
            Label = label,
            IconImage = icon,
            AccentColor = accent,
        };
        btn.Click += (_, _) => onClick();
        if (onDoubleClick is not null) btn.DoubleClick += (_, _) => onDoubleClick();

        _tip?.SetToolTip(btn, label);   // 纯图标：名称靠悬停提示补上

        _navButtons.Add(btn);
        _stack.Add(btn);
        Controls.Add(btn);
        return btn;
    }

    /// <summary>
    /// 布局：按钮从标题栏下方开始，把剩余高度**平均分给各个按钮**，
    /// 这样贴满屏幕高度时不会在底部留一大块空白。
    /// </summary>
    private void LayoutStack()
    {
        int w = Math.Max(10, ClientSize.Width);
        int fixedH = 2;
        int sepTotal = 0;
        foreach (var c in _stack)
        {
            if (c is SidebarButton) fixedH += MinBtnH;
            else sepTotal += c.Height;
        }

        int avail = ClientSize.Height - TitleBarH - 4 - sepTotal;
        int btnH = _navButtons.Count > 0
            ? Math.Max(MinBtnH, avail / _navButtons.Count)
            : MinBtnH;

        int y = TitleBarH + 2;
        foreach (var c in _stack)
        {
            int h = c is SidebarButton ? btnH : c.Height;
            c.SetBounds(0, y, w, h);
            y += h;
        }

        _overlay?.SetBounds(0, TitleBarH, w, Math.Max(0, ClientSize.Height - TitleBarH));
        LayoutOverlayChildren();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        ApplyRoundedRegion();     // 尺寸一变就重算圆角
        if (_uiBuilt)
        {
            // 注意：这里**不要**改 _dockWidth。
            // 手动拖动已经通过 _settings.ManualWidth 记录了，宽度归属由 ApplyDock 统一判定。
            LayoutStack();
            Invalidate();
        }
    }

    // ===================================================================
    // 绘制：标题栏 + 关闭按钮 + 状态点
    // ===================================================================
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        // 标题栏底色 + 与按钮区之间一条细分隔线
        using (var tb = new SolidBrush(TitleBarColor))
            g.FillRectangle(tb, 0, 0, Width, TitleBarH);
        using (var sepPen = new Pen(SeparatorCol, 1f))
            g.DrawLine(sepPen, 0, TitleBarH - 0.5f, Width, TitleBarH - 0.5f);

        // 右上角 ✕：收进托盘（不是退出程序）。hover 时变红圆，明确是"收掉这一侧栏"。
        var closeRect = CloseButtonRect();
        if (_hoverClose)
        {
            using var cbg = new SolidBrush(AccentExit);
            using var cpath = RoundedRectPath(closeRect, Math.Max(2f, closeRect.Width * 0.28f));
            g.FillPath(cbg, cpath);
        }
        using (var xPen = new Pen(_hoverClose ? Color.White : Color.FromArgb(96, 100, 110),
                                  Math.Max(1.5f, L(0.9))))
        {
            int pad = Math.Max(3, (int)Math.Round(closeRect.Width * 0.28));
            g.DrawLine(xPen, closeRect.Left + pad, closeRect.Top + pad,
                             closeRect.Right - pad, closeRect.Bottom - pad);
            g.DrawLine(xPen, closeRect.Right - pad, closeRect.Top + pad,
                             closeRect.Left + pad, closeRect.Bottom - pad);
        }

        // 标题栏不再画「PPT 侧栏」文字，也不放圆点：整条都留给拖动把手。

        // ---- 拖动手感提示：✕ 以外的整条标题栏都能拖，中间画一排"抓手"短竖线 ----
        {
            int gripLeft = Mm2Px(2);
            int gripRight = closeRect.Left - Mm2Px(2);
            int gap = gripRight - gripLeft;
            if (gap >= Mm2Px(4))
            {
                using var dragPen = new Pen(Color.FromArgb(150, 155, 165), Math.Max(1.2f, L(0.8)));
                float dy = TitleBarH / 2f;
                float cx0 = (gripLeft + gripRight) / 2f;
                int step = Math.Max(3, Math.Min(Mm2Px(1.6), gap / 6));
                int half = Math.Max(2, Mm2Px(1.4));
                for (int i = -1; i <= 1; i++)
                    g.DrawLine(dragPen, cx0 + i * step, dy - half, cx0 + i * step, dy + half);
            }
        }

        // 宽度抓握区：画在【朝向幻灯片的内侧】、标题栏以下，hover 点亮。
        bool hot = _hoverGrip;
        using (var gripPen = new Pen(
                   hot ? AccentNav : Color.FromArgb(202, 206, 214),
                   hot ? Math.Max(2f, L(2)) : Math.Max(1f, L(1))))
        {
            float gx = GripOnInnerLeft ? ResizeGripW / 2f : Width - ResizeGripW / 2f;
            g.DrawLine(gripPen, gx, TitleBarH, gx, Height - L(4));
        }

        // 整块卡片描边（浅底在浅色幻灯片上也能看清边界）
        using (var borderPen = new Pen(CardBorder, 1f))
            g.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);
    }

    /// <summary>
    /// 右上角 ✕ 的位置：按毫米目标取 10 mm 见方（关闭属于"代价大"的操作，
    /// 微软的建议是至少 9 mm），贴右侧、上下居中。
    /// </summary>
    private Rectangle CloseButtonRect()
    {
        int size = Math.Min(Mm2Px(MmCloseBtn), Math.Max(8, TitleBarH - Mm2Px(2)));
        int x = Width - size - Mm2Px(2);
        int y = Math.Max(0, (TitleBarH - size) / 2);
        return new Rectangle(x, y, size, size);
    }

    /// <summary>圆角矩形路径（半径自动夹到不超过短边一半）。</summary>
    private static GraphicsPath RoundedRectPath(Rectangle r, float radius)
    {
        float d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
        var p = new GraphicsPath();
        if (d <= 1f) { p.AddRectangle(r); return p; }
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    /// <summary>
    /// 把窗体裁成圆角：**只圆内侧两角**（朝幻灯片那侧），贴屏那侧保持直角 ——
    /// 既有圆角卡片观感，又不会在贴边处留缺口（贴边严丝合缝是首要诉求）。
    /// </summary>
    private void ApplyRoundedRegion()
    {
        int r = CornerR;
        if (r <= 0 || Width <= 2 || Height <= 2) { Region = null; return; }

        int w = Width, h = Height, d = r * 2;
        using var path = new GraphicsPath();
        if (GripOnInnerLeft)
        {
            // 内侧是左边 → 左两角圆、右两角直角
            path.AddArc(0, 0, d, d, 180, 90);
            path.AddLine(r, 0, w, 0);
            path.AddLine(w, 0, w, h);
            path.AddLine(w, h, r, h);
            path.AddArc(0, h - d, d, d, 90, 90);
        }
        else
        {
            // 内侧是右边 → 右两角圆、左两角直角
            path.AddLine(0, 0, w - r, 0);
            path.AddArc(w - d, 0, d, d, 270, 90);
            path.AddArc(w - d, h - d, d, d, 0, 90);
            path.AddLine(w - r, h, 0, h);
            path.AddLine(0, h, 0, 0);
        }
        path.CloseFigure();

        var old = Region;
        Region = new Region(path);
        old?.Dispose();
    }

    // ===================================================================
    // 鼠标
    // ===================================================================
    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);

        // ★ 左右键都算。触控被"提升"成鼠标后，手指按住不放约 0.5 秒会变成**右键**
        //   （"按住 = 右键"是 Windows 的标准手势）。只认左键的话，
        //   老师按得稍微久一点就完全没反应 —— 这是触控大屏上最容易踩的坑。
        if (e.Button is not (MouseButtons.Left or MouseButtons.Right)) return;

        if (e.Button == MouseButtons.Left && CloseButtonRect().Contains(e.Location))
        {
            _userHidden = true;
            HideToTray();          // ★ 收托盘，不退出程序
            return;
        }

        // 宽度抓握区：只在标题栏【以下】生效。
        // 否则它会和标题栏的「拖动把手」抢同一块区域。
        if (e.Y > TitleBarH && IsInGripZone(e.X))
        {
            _resizing = true;
            _resizeStartWidth = Width;
            _resizeStartMouseX = Cursor.Position.X;
            Cursor = Cursors.SizeWE;
            return;
        }

        // 标题栏：✕ 以外**整条**都能拖。
        // （原来这里塞了红/黄两个圆点，把标题栏占满、只剩十几像素的拖动空档，
        //   手指根本拖不动。红点和 ✕ 功能重复、黄点的"换边"拖动 + 托盘菜单都能做，
        //   所以两个点全删了 —— 目标小到 3 mm 的控件本来就不该给触控用。）
        if (e.Y <= TitleBarH)
        {
            _dragging = true;
            _dragStartMouse = Cursor.Position;
            _dragStartForm = Location;
            Cursor = Cursors.SizeAll;
        }
    }

    /// <summary>
    /// 宽度抓握区应该在【朝向幻灯片的那一侧】（内侧），而不是固定右边缘。
    ///
    /// 原因：贴右边缘时侧栏的右边缘被钉在屏幕最右边，往右没有空间了 ——
    /// 把手放在右边就只能把侧栏拖窄、永远拖不宽。所以贴右边时把手放左边。
    /// </summary>
    private bool GripOnInnerLeft => _dock == DockSide.Right;

    private bool IsInGripZone(int x) =>
        GripOnInnerLeft ? x <= ResizeGripW : x >= Width - ResizeGripW;

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);

        if (_resizing)
        {
            int minPhys = Mm2Px(MmWidthMin);
            int maxPhys = Mm2Px(MmWidthMax);
            int dx = Cursor.Position.X - _resizeStartMouseX;      // 物理像素

            // 把手在【内侧左边】时（贴右边缘），往左拖才是变宽 —— 方向要反过来
            if (GripOnInnerLeft) dx = -dx;

            int newW = Math.Clamp(_resizeStartWidth + dx, minPhys, maxPhys);

            _dockWidth = newW;
            // 存**毫米**，这样换台屏幕（不同 ppi）依然得到同样的物理宽度。
            // 命令行 --width 生效时不写盘，免得命令行的值被"记住"成用户偏好。
            if (_overrideWidthMm < 0)
                _settings.WidthMm = Math.Round(newW / Math.Max(1e-9, _metrics.PhysicalPxPerMm), 1);

            // 外侧那条边固定不动：贴右边固定 Right，贴左边固定 Left
            if (_dock == DockSide.Right) SetBounds(Right - newW, Top, newW, Height);
            else SetBounds(Left, Top, newW, Height);

            LayoutStack();
            return;
        }

        if (_dragging)
        {
            var pos = Cursor.Position;
            var wa = TargetScreen().WorkingArea;
            int nx = _dragStartForm.X + (pos.X - _dragStartMouse.X);
            int ny = _dragStartForm.Y + (pos.Y - _dragStartMouse.Y);

            // 拖动过程中把窗口限制在工作区内，别让它被拖到看不见的地方。
            // 左右各允许越界半个宽度，这样能顺畅地拖到对面的边。
            nx = Math.Clamp(nx, wa.Left - Width / 2, Math.Max(wa.Left - Width / 2, wa.Right - Width / 2));
            ny = Math.Clamp(ny, wa.Top, Math.Max(wa.Top, wa.Bottom - Height));

            Location = new Point(nx, ny);
            return;
        }

        bool overClose = CloseButtonRect().Contains(e.Location);
        // 抓握区只在标题栏以下算，标题栏那块留给拖动
        bool overGrip = e.Y > TitleBarH && IsInGripZone(e.X);
        if (overClose != _hoverClose || overGrip != _hoverGrip)
        {
            _hoverClose = overClose;
            _hoverGrip = overGrip;
            Invalidate();
        }

        if (overGrip) Cursor = Cursors.SizeWE;
        else if (e.Y <= TitleBarH) Cursor = Cursors.SizeAll;
        else Cursor = Cursors.Default;
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (!_dragging && !_resizing) return;

        bool wasDragging = _dragging;
        bool wasResizing = _resizing;
        _dragging = false;
        _resizing = false;
        Cursor = Cursors.Default;

        // 用户调过宽度就落盘记住（拖动过程中不写，避免频繁写盘）
        if (wasResizing && _overrideWidthMm < 0) _settings.Save();

        if (wasDragging)
        {
            // ★ 只能贴边：松手后一定吸到最近的那条屏幕边，**不允许停在中间**。
            //   判据是"窗口中心落在屏幕左半边还是右半边"——
            //   这样轻轻动一下不会突然换边，拖过半屏才换。
            var wa = TargetScreen().WorkingArea;
            int centerX = Left + Width / 2;
            _dock = centerX < wa.Left + wa.Width / 2 ? DockSide.Left : DockSide.Right;
            _dockTop = Top;          // 记住用户挑的垂直位置
            ApplyDock();
        }
        ReassertTopMost();
        Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        if (_hoverClose || _hoverGrip)
        {
            _hoverClose = false;
            _hoverGrip = false;
            Invalidate();
        }
    }

    private void OnFormClosing(object? sender, CancelEventArgs e)
    {
        _pollTimer.Stop();
        _topMostTimer.Stop();
        if (_tray is not null) { _tray.Visible = false; _tray.Dispose(); }
        _trayIcon?.Dispose();
        _menu?.Dispose();
        _ppt.Dispose();
    }

    // ===================================================================
    // 业务动作
    // ===================================================================
    private void OnPrevClicked() => Guard("上一页失败", () => _ppt.PreviousSlide());
    private void OnNextClicked() => Guard("下一页失败", () => _ppt.NextSlide());
    private void OnPenClicked() => Guard("切换笔失败", () => _ppt.UsePen(_penColor));
    // 荧光笔按钮已删除（见 BuildUi 里的说明），所以这里没有 OnHighlighterClicked。
    private void OnClearInkClicked() => Guard("清空笔迹失败", () => _ppt.ClearCurrentSlideInk());

    /// <summary>
    /// 橡皮要单独处理：PowerPoint 在本页没有墨迹时会静默忽略橡皮指令，
    /// 这时必须给用户一句明确解释，否则又变成"点了没反应"。
    /// </summary>
    private void OnEraserClicked()
    {
        try
        {
            // 只有真的没切上才提示一句 —— 这种情况不说用户会以为点击失灵
            if (!_ppt.UseEraser())
                ShowMessage("本页还没有笔迹，橡皮没切上。先画两笔再擦。");
        }
        catch (Exception ex)
        {
            ShowMessage("切换橡皮失败\n\n" + ExplainError(ex));
        }
        finally
        {
            _ppt.ActivateMainWindow();
            SyncSlideshowState();   // ★ 同上：统一走一个入口
        }
    }

    /// <summary>应用新的笔色：更新按钮色块与强调色，并持久化。</summary>
    private void ApplyColor(Color color)
    {
        _penColor = color;
        _settings.PenColor = color;
        if (_penBtn is not null)
        {
            _penBtn.ChipColor = color;
            _penBtn.AccentColor = color;
            _penBtn.Invalidate();
        }
        _settings.Save();

        // 如果此刻正在放映，立刻把新颜色套用到 PowerPoint 的笔上
        if (_isInSlideshow)
        {
            try { _ppt.UsePen(color); }
            catch { /* 套用失败不影响设置保存 */ }
        }
    }

    /// <summary>
    /// 结束放映。**直接执行，不再弹确认框** —— 点错了再按一下 F5 就回来了，
    /// 而每次都要多点一次"确定"在放映现场很打断（用户明确要求去掉这个提示）。
    /// </summary>
    private void OnExitClicked() => Guard("结束放映失败", () => _ppt.ExitSlideshow());

    private void Guard(string title, Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            ShowMessage(title + "\n\n" + ExplainError(ex));
        }
        finally
        {
            // NOACTIVATE 已经不抢焦点了，这里再兜一层，确保键盘回到 PowerPoint
            _ppt.ActivateMainWindow();
            // ★ 必须走统一入口：点「结束放映」也要能自动收托盘
            SyncSlideshowState();
        }
    }

    private static string ExplainError(Exception ex)
    {
        var inner = ex;
        while (inner.InnerException is not null) inner = inner.InnerException;

        if (inner is COMException com)
        {
            return com.HResult switch
            {
                unchecked((int)0x80020003) =>
                    "PowerPoint 拒绝了这次调用（找不到成员）。\n最常见原因是当前没有在放映幻灯片。",
                unchecked((int)0x80010108) =>
                    "与 PowerPoint 的连接已断开（PowerPoint 可能已被关闭）。\n重新打开 PowerPoint 后会自动重连。",
                unchecked((int)0x800401E3) =>
                    "操作不可用（没有正在运行的 PowerPoint 实例）。",
                _ => $"PowerPoint 返回错误 0x{com.HResult:X8}\n{com.Message}"
            };
        }

        return inner.Message;
    }

    private void ShowStatus()
    {
        int pid = Environment.ProcessId;
        var hwnd = _ppt.GetMainWindowHandle();
        string text =
            $"侧栏 PID：{pid}\n" +
            $"贴边位置：{(_dock == DockSide.Left ? "屏幕左侧" : "屏幕右侧")}\n" +
            $"侧栏尺寸：{Bounds.Width / Math.Max(1e-9, _metrics.PhysicalPxPerMm):0.#} × "
                       + $"{Bounds.Height / Math.Max(1e-9, _metrics.PhysicalPxPerMm):0.#} mm\n" +
            $"PowerPoint 窗口：0x{hwnd.ToInt64():X}\n" +
            $"正在放映：{(_isInSlideshow ? "是" : "否")}\n\n" +
            "关闭入口：\n" +
            "  · 点侧栏右上角 ✕（收进托盘）\n" +
            "  · 托盘图标右键 → 退出侧栏\n" +
            "  · 双击「关闭侧栏.cmd」";

        ShowIfNotUserHidden(resetFlag: true);
        ShowMessage(text);
    }

    // ===================================================================
    // 内嵌提示条
    //   WS_EX_NOACTIVATE 会让模态 MessageBox 无法正常激活（点不动按钮），
    //   所以所有提示/确认都改成侧栏内的覆盖层。
    // ===================================================================
    private Panel? _overlay;
    private Label? _overlayText;
    private Button? _overlayOk;          // 只需要一个「知道了」按钮：提示会自己消失
    private System.Windows.Forms.Timer? _overlayAutoHide;

    /// <summary>悬停提示（纯图标布局下用来显示按钮名称）。必须作为字段存活，否则会被 GC。</summary>
    private ToolTip? _tip;

    private void BuildOverlay()
    {
        _overlayText = new Label
        {
            ForeColor = TextColor,
            BackColor = Color.Transparent,
            Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular),
            AutoSize = false,
        };

        _overlayOk = MakeOverlayButton("知道了", AccentNav);
        _overlayOk.Click += (_, _) => HideOverlay();

        _overlay = new Panel { BackColor = OverlayColor, Visible = false };
        _overlay.Controls.Add(_overlayText);
        _overlay.Controls.Add(_overlayOk);
        Controls.Add(_overlay);
        _overlay.BringToFront();

        BuildPalette();

        _overlayAutoHide = new System.Windows.Forms.Timer();
        _overlayAutoHide.Tick += (_, _) => HideOverlay();
    }

    // -------------------------------------------------------------------
    // 调色板
    //   为什么不做任意 RGB / 系统取色器：
    //     1. 侧栏带 WS_EX_NOACTIVATE，覆盖层里的输入框拿不到键盘焦点，
    //        hex 输入这类交互直接不可用；
    //     2. Windows 系统取色器是需要激活的模态对话框，在全屏放映上方弹出不可靠。
    //   15 个预设色对"放映时标注"这个场景已经够用。
    // -------------------------------------------------------------------
    private Panel? _palettePanel;
    private Label? _paletteTitle;
    private Button[] _swatches = Array.Empty<Button>();
    private Button? _paletteReset;
    private Button? _paletteClose;

    // 调色板的间距/尺寸全部按**毫米**，否则在 200% 缩放的屏上只有一半大、点不中
    private const int PaletteCols = 5;
    private const double MmPaletteGap = 1.5;
    private const double MmPalettePad = 3;
    private const double MmPaletteTitle = 9;
    private const double MmPaletteBtn = 9;
    private const double MmSwatch = 8;     // 色块最小边长（触控）

    private void BuildPalette()
    {
        if (_overlay is null) return;

        _paletteTitle = new Label
        {
            ForeColor = TextColor,
            BackColor = Color.Transparent,
            Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
            AutoSize = false,
            TextAlign = ContentAlignment.MiddleCenter,
        };

        _palettePanel = new Panel { BackColor = Color.Transparent, Visible = false };
        _palettePanel.Controls.Add(_paletteTitle);

        _swatches = new Button[PaletteColors.Length];
        for (int i = 0; i < PaletteColors.Length; i++)
        {
            var c = PaletteColors[i];
            var b = new Button
            {
                FlatStyle = FlatStyle.Flat,
                BackColor = c,
                TabStop = false,
                Cursor = Cursors.Hand,
            };
            // 悬停/按下都不改底色 —— 色块必须始终如实显示自己的颜色
            b.FlatAppearance.MouseOverBackColor = c;
            b.FlatAppearance.MouseDownBackColor = c;
            b.FlatAppearance.BorderColor = Color.FromArgb(202, 206, 214);   // 浅底上用浅灰边
            b.FlatAppearance.BorderSize = 1;

            int idx = i;
            b.Click += (_, _) =>
            {
                ApplyColor(PaletteColors[idx]);
                HideOverlay();
            };

            _swatches[i] = b;
            _palettePanel.Controls.Add(b);
        }

        _paletteReset = MakeOverlayButton("恢复默认", Color.FromArgb(120, 126, 138));
        _paletteReset.Click += (_, _) =>
        {
            ApplyColor(SettingsStore.DefaultPenColor);
            HideOverlay();
        };

        _paletteClose = MakeOverlayButton("关闭", AccentNav);
        _paletteClose.Click += (_, _) => HideOverlay();

        _palettePanel.Controls.Add(_paletteReset);
        _palettePanel.Controls.Add(_paletteClose);

        _overlay.Controls.Add(_palettePanel);
    }

    /// <summary>打开调色板（覆盖层进入调色板模式）。只给笔换色 —— 荧光笔按钮已删除。</summary>
    private void OpenPalette()
    {
        if (_overlay is null || _palettePanel is null) return;

        if (_paletteTitle is not null) _paletteTitle.Text = "笔的颜色";

        // 调色板模式：藏掉消息模式的控件
        if (_overlayText is not null) _overlayText.Visible = false;
        if (_overlayOk is not null) _overlayOk.Visible = false;
        _palettePanel.Visible = true;

        HighlightSelectedSwatch(_penColor);

        _overlay.SetBounds(0, TitleBarH, ClientSize.Width, Math.Max(0, ClientSize.Height - TitleBarH));
        _palettePanel.SetBounds(0, 0, _overlay.ClientSize.Width, _overlay.ClientSize.Height);
        LayoutPalette();

        _overlay.Visible = true;
        _overlay.BringToFront();
        _overlayAutoHide?.Stop();
        ReassertTopMost();
    }

    /// <summary>把当前生效的颜色那块色块描成白粗边，一眼看出选的是哪个。</summary>
    private void HighlightSelectedSwatch(Color current)
    {
        foreach (var b in _swatches)
        {
            bool selected = b.BackColor.ToArgb() == current.ToArgb();
            b.FlatAppearance.BorderColor = selected ? AccentNav : Color.FromArgb(202, 206, 214);
            b.FlatAppearance.BorderSize = selected ? 3 : 1;
        }
    }

    /// <summary>
    /// 排列调色板。★ 列数是**算出来的**，不是写死的 5 列：
    /// 侧栏只有 38 mm 宽，5 列会让每个色块只剩 5 mm —— 手指点不中。
    /// 这里先按"每块至少 6 mm"选列数，塞不下就减列增行。
    /// </summary>
    private void LayoutPalette()
    {
        if (_overlay is null || _palettePanel is null || _paletteTitle is null) return;

        int w = _overlay.ClientSize.Width;
        int h = _overlay.ClientSize.Height;

        int pad = Mm2Px(MmPalettePad);
        int gap = Mm2Px(MmPaletteGap);
        int titleH = Mm2Px(MmPaletteTitle);
        int btnH = Mm2Px(MmPaletteBtn);
        int minSw = Math.Max(12, Mm2Px(5));      // 色块绝对最小边长

        int availW = Math.Max(30, w - pad * 2);
        int availH = Math.Max(40, h - pad * 2 - titleH - btnH - Mm2Px(4));

        // 按"每块至少 6 mm"定列数，再据此定行数
        int want = Math.Max(1, Mm2Px(6));
        int cols = Math.Clamp((availW + gap) / (want + gap), 2, PaletteCols);
        int rows = (_swatches.Length + cols - 1) / cols;

        int sw = Math.Max(minSw, (availW - gap * (cols - 1)) / cols);
        int sh = Math.Min(sw, Math.Max(minSw, (availH - gap * (rows - 1)) / rows));

        int gridH = rows * sh + (rows - 1) * gap;
        int blockH = titleH + Mm2Px(2) + gridH + Mm2Px(3) + btnH;

        int top = Math.Max(pad, (h - blockH) / 2);
        _paletteTitle.SetBounds(pad, top, Math.Max(10, w - pad * 2), titleH);

        int gridTop = top + titleH + Mm2Px(2);
        for (int i = 0; i < _swatches.Length; i++)
        {
            int r = i / cols, c = i % cols;
            _swatches[i].SetBounds(pad + c * (sw + gap), gridTop + r * (sh + gap), sw, sh);
        }

        int btnY = gridTop + gridH + Mm2Px(3);
        int btnW = Math.Max(30, (w - pad * 3) / 2);
        _paletteReset?.SetBounds(pad, btnY, btnW, btnH);
        _paletteClose?.SetBounds(pad * 2 + btnW, btnY, btnW, btnH);
    }

    /// <summary>覆盖层里的按钮。高度按毫米定 —— 原来写死 30 物理像素，200% 缩放下只有 4 mm。</summary>
    private Button MakeOverlayButton(string text, Color back) => new()
    {
        Text = text,
        FlatStyle = FlatStyle.Flat,
        BackColor = back,
        ForeColor = Color.White,
        Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
        Height = Mm2Px(MmPaletteBtn),
        TabStop = false,
    };

    private void LayoutOverlayChildren()
    {
        if (_overlay is null || _overlayText is null || _overlayOk is null) return;
        int w = _overlay.ClientSize.Width;
        int h = _overlay.ClientSize.Height;

        int pad = Mm2Px(MmPalettePad);
        int btnH = Mm2Px(MmPaletteBtn);

        _overlayText.SetBounds(pad, pad, Math.Max(10, w - pad * 2), Math.Max(10, h - pad * 2 - btnH - Mm2Px(3)));

        // 只有一个按钮，横贯整宽（原来是一左一右两个按钮）
        _overlayOk.SetBounds(pad, h - pad - btnH, Math.Max(30, w - pad * 2), btnH);

        // 调色板跟着覆盖层一起重排（侧栏改宽度时会走到这里）
        if (_palettePanel is not null)
        {
            _palettePanel.SetBounds(0, 0, w, h);
            LayoutPalette();
        }
    }

    /// <summary>
    /// 覆盖层提示条 —— 只用于「真的出错了」和调色板。
    /// 4.5 秒后自动消失，所以不需要用户点任何东西。
    /// </summary>
    private void ShowMessage(string text)
    {
        if (_overlay is null || _overlayText is null) return;
        _overlayText.Text = text;
        ShowOverlay(autoHideMs: 4500);
    }

    private void ShowOverlay(int autoHideMs)
    {
        if (_overlay is null || _overlayOk is null) return;

        // 退出调色板模式：消息模式和调色板模式共用同一个覆盖层
        if (_palettePanel is not null) _palettePanel.Visible = false;

        _overlay.SetBounds(0, TitleBarH, ClientSize.Width, Math.Max(0, ClientSize.Height - TitleBarH));
        LayoutOverlayChildren();
        _overlayOk.Visible = true;

        _overlay.Visible = true;
        _overlay.BringToFront();

        _overlayAutoHide?.Stop();
        if (autoHideMs > 0)
        {
            _overlayAutoHide!.Interval = autoHideMs;
            _overlayAutoHide.Start();
        }
    }

    private void HideOverlay()
    {
        _overlayAutoHide?.Stop();
        if (_palettePanel is not null) _palettePanel.Visible = false;
        if (_overlay is not null) _overlay.Visible = false;
    }
}

// =======================================================================
// 自绘按钮
// =======================================================================
[ToolboxItem(false)]
public class SidebarButton : Panel
{
    // ---- 浅色卡片配色：白底 / hover 浅蓝 / 按下更深蓝 ----
    private static readonly Color BtnBg      = Color.FromArgb(252, 252, 253);
    private static readonly Color BtnHover   = Color.FromArgb(232, 240, 254);
    private static readonly Color BtnPressed = Color.FromArgb(214, 228, 252);

    private bool _hover, _pressed;
    private bool _chipHover;
    private bool _suppressClick;   // 点了色块就把这次 Click 拦掉

    public string Label { get; set; } = "";
    public Image? IconImage { get; set; }
    public Color AccentColor { get; set; } = Color.White;

    /// <summary>非 null 时在按钮右侧画一个色块，点它会走 OnChipClick 而不是正常的 Click。</summary>
    public Color? ChipColor { get; set; }

    /// <summary>点色条时触发。</summary>
    public Action? OnChipClick { get; set; }

    /// <summary>
    /// 换色色条宽度：按按钮高度取比例。
    ///
    /// ★ 从"右侧小方块"改成"右侧竖条"：纯图标布局下图标居中，色条靠右不抢位；
    ///   竖条的点击面积也比小方块大得多（触控友好）。
    ///   尺寸按**高度比例**算 —— 高度是按毫米定的，所以跨屏物理一致。
    /// </summary>
    private int ChipWidth() => Math.Max(6, (int)Math.Round(Height * 0.30));

    private int ChipInset() => Math.Max(1, (int)Math.Round(Height * 0.16));

    public Rectangle ChipRect()
    {
        if (ChipColor is null) return Rectangle.Empty;
        int w = ChipWidth();
        int inset = ChipInset();
        return new Rectangle(Width - w - inset, inset, w, Math.Max(4, Height - inset * 2));
    }

    public SidebarButton()
    {
        SetStyle(
            ControlStyles.UserPaint |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw |
            ControlStyles.StandardClick |
            ControlStyles.StandardDoubleClick,
            true);
        Cursor = Cursors.Hand;
        BackColor = BtnBg;
        Margin = new Padding(0);
        Padding = new Padding(0);
    }

    protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hover = false;
        _pressed = false;
        _chipHover = false;
        Cursor = Cursors.Hand;
        Invalidate();
        base.OnMouseLeave(e);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        bool overChip = ChipColor is not null && ChipRect().Contains(e.Location);
        if (overChip != _chipHover)
        {
            _chipHover = overChip;
            Cursor = overChip ? Cursors.Hand : Cursors.Hand;
            Invalidate();
        }
        base.OnMouseMove(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        _suppressClick = false;
        // ★ 左右键都算。触控被提升成鼠标后，长按 ~0.5 s 会变成**右键**
        //   （Windows 标准手势）。只认左键的话，老师按得稍久就没反应 —— 触控屏必踩。
        if (e.Button is MouseButtons.Left or MouseButtons.Right) { _pressed = true; Invalidate(); }
        base.OnMouseDown(e);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        _pressed = false;

        // 点在色块上 → 走 OnChipClick，并把随之而来的 Click 拦掉
        // （右键也算 —— 长按笔按钮右半边仍能换色，不该被吞掉）
        if (e.Button is MouseButtons.Left or MouseButtons.Right
            && ChipColor is not null && ChipRect().Contains(e.Location))
        {
            _suppressClick = true;
            Invalidate();
            var cb = OnChipClick;
            cb?.Invoke();
        }

        base.OnMouseUp(e);
    }

    protected override void OnClick(EventArgs e)
    {
        if (_suppressClick)
        {
            _suppressClick = false;
            return;   // 这次点击是冲色块来的，不要触发按钮本身的动作
        }
        base.OnClick(e);
    }

    // ===================================================================
    // 绘制：纯图标 + 圆角卡片
    //   ★ 这一版去掉了按钮文字（宽度只有 30 mm，塞不下"图标 + 文字"，
    //     文字会让按钮又高又挤）。名称改由悬停 ToolTip 给出。
    // ===================================================================
    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;

        // ---- 圆角背景 ----
        var rect = new Rectangle(0, 0, Math.Max(1, Width - 1), Math.Max(1, Height - 1));
        int radius = Math.Max(3, (int)Math.Round(Height * 0.24));
        using var path = RoundedPath(rect, radius);

        var bg = _pressed ? BtnPressed : (_hover ? BtnHover : BtnBg);
        using (var bgBrush = new SolidBrush(bg))
            g.FillPath(bgBrush, path);

        // hover / 按下：用强调色描边，反馈清晰
        if (_hover || _pressed)
        {
            using var hl = new Pen(AccentColor, _pressed ? 2.4f : 1.6f);
            g.DrawPath(hl, path);
        }

        // ---- 图标：纯图标居中，只避开右侧色条 ----
        int chipReserve = ChipColor is not null ? ChipWidth() + ChipInset() : 0;
        int contentW = Math.Max(12, Width - chipReserve);
        int iconBox = Math.Clamp(Math.Min(contentW - 2, Height - 4), 14, 96);
        var iconRect = new Rectangle((contentW - iconBox) / 2, (Height - iconBox) / 2, iconBox, iconBox);
        if (IconImage is not null)
            g.DrawImage(IconImage, iconRect);

        // ---- 右侧色条（点它换颜色）----
        if (ChipColor is Color chip)
        {
            var cr = ChipRect();
            float crRadius = Math.Max(2, cr.Width / 2f);
            using (var chipPath = RoundedPath(cr, crRadius))
            using (var chipBrush = new SolidBrush(chip))
                g.FillPath(chipBrush, chipPath);

            bool light = (chip.R * 299 + chip.G * 587 + chip.B * 114) / 1000 > 140;
            var edge = _chipHover
                ? AccentColor
                : (light ? Color.FromArgb(178, 184, 194) : Color.FromArgb(150, 150, 158));
            using var chipPen = new Pen(edge, _chipHover ? 2.2f : 1.2f);
            using var chipPath2 = RoundedPath(cr, crRadius);
            g.DrawPath(chipPen, chipPath2);
        }
    }

    /// <summary>圆角矩形路径（半径会自动夹到不超过短边的一半）。</summary>
    private static GraphicsPath RoundedPath(Rectangle r, float radius)
    {
        float d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
        var p = new GraphicsPath();
        if (d <= 1f)
        {
            p.AddRectangle(r);
            return p;
        }
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}