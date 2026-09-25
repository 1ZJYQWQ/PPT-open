using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PowerPointSidebar;

/// <summary>
/// 屏幕物理尺寸探测 —— 把【毫米】换算成像素的唯一依据。
///
/// 为什么必须做这件事：
///   这个侧栏的触控目标要按**毫米**定（手指是物理尺寸固定的东西），
///   但 Windows 只给"逻辑像素"，而"1 逻辑像素 = 几毫米"完全取决于屏幕 ppi：
///     65" 4K 大屏   → 0.49 mm / 逻辑像素（100% 缩放）
///     86" 4K 大屏   → 0.49 mm
///     13" 笔记本    → 0.13 mm
///   同一份"10 毫米"的目标，在这些屏上差 3 倍以上。所以不能用固定像素值。
///
/// 物理尺寸的唯一可靠来源是显示器 **EDID**（第 21 / 22 字节 = 水平 / 垂直可视尺寸，单位厘米），
/// 官方文档明说某些显示器会返回 0（比如投影，尺寸本来就不确定），所以：
///   1. 读不到 → 退回名义 96 DPI（1 逻辑像素 = 1/96 英寸 ≈ 0.2646 mm）
///   2. 仍然允许命令行覆盖：--panel 86（按 16:9 反算）或 --mm-per-dip 0.49
///
/// 注意：**不要**用 Win32 的 GetDeviceCaps(HORZSIZE)。它多数情况下只是拿标称 DPI
/// 反推出来的名义值（等于假设屏幕就是 96 DPI），永远算不出真实尺寸。
/// </summary>
public static class ScreenMetrics
{
    /// <summary>名义 1/96 英寸。EDID 读不到时的兜底值。</summary>
    public const double NominalMmPerDip = 25.4 / 96.0;   // ≈ 0.2646

    /// <summary>命令行强制指定：每逻辑像素多少毫米。&lt;= 0 表示未指定。</summary>
    public static double ForcedMmPerDip { get; set; }

    /// <summary>命令行强制指定：屏幕对角英寸数（按 16:9 反算物理宽度）。&lt;= 0 表示未指定。</summary>
    public static double ForcedDiagonalInch { get; set; }

    public sealed class Info
    {
        /// <summary>每逻辑像素多少毫米 —— 这才是程序要用的数。</summary>
        public double MmPerDip = NominalMmPerDip;

        /// <summary>可视区物理宽度（毫米），由 EDID 得到。</summary>
        public double PhysicalWidthMm;

        /// <summary>该屏的物理像素宽度。</summary>
        public int ResolutionWidth;

        /// <summary>该屏的有效 DPI（100% = 96）。</summary>
        public int Dpi = 96;

        /// <summary>每英寸像素数。</summary>
        public double Ppi;

        /// <summary>是否来自 EDID（false 表示用的是兜底值）。</summary>
        public bool FromEdid;

        /// <summary>给人看的说明，--where 里会打印。</summary>
        public string Source = "";

        /// <summary>1 毫米等于多少逻辑像素。</summary>
        public double DipPerMm => MmPerDip > 0 ? 1.0 / MmPerDip : 1.0 / NominalMmPerDip;

        /// <summary>1 毫米等于多少物理像素。</summary>
        public double PhysicalPxPerMm => DipPerMm * (Dpi / 96.0);
    }

    private static readonly Dictionary<string, Info> Cache = new();

    // ================= Win32 =================
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct DISPLAY_DEVICE
    {
        public int cb;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]  public string DeviceName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string DeviceString;
        public int StateFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string DeviceID;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string DeviceKey;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X, Y; }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "EnumDisplayDevicesW")]
    private static extern bool EnumDisplayDevices(string? device, uint devNum,
        ref DISPLAY_DEVICE dd, uint flags);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromPoint(POINT pt, uint flags);

    [DllImport("shcore.dll")]
    private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

    private const uint EDD_GET_DEVICE_ID = 0x00000001;
    private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
    private const int MDT_EFFECTIVE_DPI = 0;

    /// <summary>
    /// 取某块屏幕的物理尺寸信息。
    /// <paramref name="fallbackDpi"/> 是拿不到每屏 DPI 时的退路（一般传窗体自己的 DeviceDpi）。
    /// </summary>
    public static Info For(Screen screen, int fallbackDpi)
    {
        string key = $"{screen.DeviceName}|{screen.Bounds.Width}x{screen.Bounds.Height}|{fallbackDpi}"
                     + $"|{ForcedMmPerDip:0.#####}|{ForcedDiagonalInch:0.##}";
        if (Cache.TryGetValue(key, out var hit)) return hit;

        var info = new Info
        {
            ResolutionWidth = screen.Bounds.Width,   // PerMonitorV2 下这里是物理像素
            Dpi = fallbackDpi > 0 ? fallbackDpi : 96,
        };

        // ---- 每屏真实 DPI（比窗体自己的 DPI 准，多屏缩放不同时尤其重要）----
        try
        {
            var center = new POINT
            {
                X = screen.Bounds.Left + screen.Bounds.Width / 2,
                Y = screen.Bounds.Top + screen.Bounds.Height / 2,
            };
            IntPtr mon = MonitorFromPoint(center, MONITOR_DEFAULTTONEAREST);
            if (mon != IntPtr.Zero
                && GetDpiForMonitor(mon, MDT_EFFECTIVE_DPI, out uint dx, out _) == 0
                && dx > 0)
            {
                info.Dpi = (int)dx;
            }
        }
        catch { /* 拿不到就用 fallbackDpi */ }

        // ---- 1) 命令行覆盖优先 ----
        if (ForcedMmPerDip > 0)
        {
            info.MmPerDip = ForcedMmPerDip;
            info.PhysicalWidthMm = ForcedMmPerDip * (info.ResolutionWidth * 96.0 / info.Dpi);
            info.Source = $"命令行 --mm-per-dip {ForcedMmPerDip:0.####}";
            info.FromEdid = false;
            Finish(info, screen);
            Cache[key] = info;
            return info;
        }

        if (ForcedDiagonalInch > 0)
        {
            // 16:9 → 宽 = 对角 × 16/√(16²+9²)
            info.PhysicalWidthMm = ForcedDiagonalInch * 25.4 * (16.0 / Math.Sqrt(16.0 * 16.0 + 9.0 * 9.0));
            info.MmPerDip = info.PhysicalWidthMm / (info.ResolutionWidth * 96.0 / info.Dpi);
            info.Source = $"命令行 --panel {ForcedDiagonalInch:0.#}\" (按 16:9 反算)";
            info.FromEdid = false;
            Finish(info, screen);
            Cache[key] = info;
            return info;
        }

        // ---- 2) 读 EDID ----
        if (TryReadEdidWidthMm(screen, out double widthMm, out string why))
        {
            info.PhysicalWidthMm = widthMm;
            info.MmPerDip = widthMm / (info.ResolutionWidth * 96.0 / info.Dpi);
            info.FromEdid = true;
            info.Source = "EDID";
            Finish(info, screen);
            Cache[key] = info;
            return info;
        }

        // ---- 3) 兜底：名义 96 DPI ----
        info.MmPerDip = NominalMmPerDip;
        info.PhysicalWidthMm = NominalMmPerDip * (info.ResolutionWidth * 96.0 / info.Dpi);
        info.FromEdid = false;
        info.Source = "兜底（EDID 读不到：" + why + "）";
        Finish(info, screen);
        Cache[key] = info;
        return info;
    }

    private static void Finish(Info info, Screen screen)
    {
        if (info.MmPerDip <= 0) info.MmPerDip = NominalMmPerDip;
        double widthInch = info.PhysicalWidthMm / 25.4;
        double heightMm = info.PhysicalWidthMm * screen.Bounds.Height / Math.Max(1, screen.Bounds.Width);
        double diagInch = Math.Sqrt(widthInch * widthInch + Math.Pow(heightMm / 25.4, 2));
        info.Ppi = widthInch > 0 ? info.ResolutionWidth / widthInch : 0;
        DiagonalInch = diagInch;
    }

    /// <summary>最近一次算出来的屏幕对角英寸数（仅供诊断显示）。</summary>
    public static double DiagonalInch { get; private set; }

    /// <summary>
    /// 从注册表里的 EDID 二进制块取可视区宽度（毫米）。
    /// EDID 第 21 字节 = 水平可视尺寸(cm)，第 22 字节 = 垂直可视尺寸(cm)；
    /// 官方文档说明这两个字段是"四舍五入到厘米"的，而且可能为 0（未定义）。
    /// </summary>
    private static bool TryReadEdidWidthMm(Screen screen, out double widthMm, out string why)
    {
        widthMm = 0;
        why = "未知原因";
        try
        {
            // \\.\DISPLAY1 → 找到它的监视器设备 ID（形如 MONITOR\BOE0A1F\{GUID}\0001）
            string deviceId = "";
            var dd = new DISPLAY_DEVICE { cb = Marshal.SizeOf<DISPLAY_DEVICE>() };
            for (uint i = 0; EnumDisplayDevices(null, i, ref dd, 0); i++)
            {
                if (!string.Equals(dd.DeviceName, screen.DeviceName, StringComparison.OrdinalIgnoreCase))
                {
                    dd.cb = Marshal.SizeOf<DISPLAY_DEVICE>();
                    continue;
                }

                var mon = new DISPLAY_DEVICE { cb = Marshal.SizeOf<DISPLAY_DEVICE>() };
                if (EnumDisplayDevices(dd.DeviceName, 0, ref mon, EDD_GET_DEVICE_ID))
                    deviceId = mon.DeviceID ?? "";
                break;
            }

            if (string.IsNullOrEmpty(deviceId))
            {
                why = "拿不到监视器设备 ID";
                return false;
            }

            // MONITOR\BOE0A1F\{GUID}\0001
            //   → HKLM\SYSTEM\CurrentControlSet\Enum\MONITOR\BOE0A1F\{GUID}\0001\Device Parameters
            string path = deviceId.StartsWith("MONITOR\\", StringComparison.OrdinalIgnoreCase)
                ? deviceId[8..]
                : deviceId;
            string regPath = @"SYSTEM\CurrentControlSet\Enum\MONITOR\" + path + @"\Device Parameters";

            using var key = Registry.LocalMachine.OpenSubKey(regPath);
            if (key?.GetValue("EDID") is not byte[] edid)
            {
                why = "注册表里没有 EDID";
                return false;
            }
            if (edid.Length < 23)
            {
                why = $"EDID 长度不足（{edid.Length} 字节）";
                return false;
            }

            int hCm = edid[21];
            int vCm = edid[22];
            if (hCm <= 0 || vCm <= 0)
            {
                why = "EDID 里没有填可视尺寸（显示器和投影常见）";
                return false;
            }

            widthMm = hCm * 10.0;
            why = "";
            return true;
        }
        catch (Exception ex)
        {
            why = ex.GetType().Name + ": " + ex.Message;
            return false;
        }
    }
}
