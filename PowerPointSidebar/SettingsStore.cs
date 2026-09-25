using System;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace PowerPointSidebar;

/// <summary>
/// 极简设置存储。
///
/// 位置：%APPDATA%\PowerPointSidebar\settings.txt
/// 格式：每行 key=值，以 # 开头的整行是注释。
///       刻意用最简单的文本格式 —— 人类可读、可直接手改、零依赖。
///
/// 为什么不用 bin\ 目录：重新构建会把 bin 清掉，设置就丢了。
/// 为什么不用注册表：避免权限问题和卸载残留。
///
/// ★ 尺寸存的是**毫米**，不是像素。
///   像素在不同 ppi 的屏幕上代表的物理大小能差 3 倍以上（65" 4K 大屏 0.37 mm/px，
///   13" 笔记本 0.13 mm/px），所以用户"我拖到了这么宽"这个意图只能用毫米表达。
///   旧的 manualWidth（像素）会被忽略 —— 键名换了，不会把旧值当毫米误读。
///
/// 读写失败一律静默降级（用默认值 / 放弃保存）—— 设置存不下来不该让侧栏崩掉。
/// </summary>
public sealed class SettingsStore
{
    private const string KeyPen = "penColor";
    private const string KeyWidth = "widthMm";
    private const string KeyHeight = "heightMm";

    /// <summary>默认笔色：红。</summary>
    public static readonly Color DefaultPenColor = Color.FromArgb(255, 60, 60);

    public static string FilePath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "PowerPointSidebar",
        "settings.txt");

    public Color PenColor { get; set; } = DefaultPenColor;

    /// <summary>
    /// 用户手动拖过的侧栏宽度，**毫米**。0 = 用程序默认的毫米数（见 MmWidth）。
    /// </summary>
    public double WidthMm { get; set; }

    /// <summary>
    /// 侧栏高度，**毫米**。0 = 按内容自然高度（刚好装下 5 个按钮）。
    /// 只接受比自然高度更大的值：更小会把最后一个按钮裁掉（本窗体没有滚动条）。
    /// </summary>
    public double HeightMm { get; set; }

    public static SettingsStore Load()
    {
        var s = new SettingsStore();
        try
        {
            if (!File.Exists(FilePath)) return s;

            foreach (var raw in File.ReadAllLines(FilePath))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line[0] == '#') continue;   // 注释 / 空行

                int eq = line.IndexOf('=');
                if (eq <= 0) continue;

                string key = line[..eq].Trim();
                string val = line[(eq + 1)..].Trim();

                if (key.Equals(KeyWidth, StringComparison.OrdinalIgnoreCase))
                {
                    if (TryParseDouble(val, out double w) && w > 0) s.WidthMm = w;
                    continue;
                }

                if (key.Equals(KeyHeight, StringComparison.OrdinalIgnoreCase))
                {
                    if (TryParseDouble(val, out double h) && h > 0) s.HeightMm = h;
                    continue;
                }

                if (!TryParseHex(val, out var c)) continue;

                if (key.Equals(KeyPen, StringComparison.OrdinalIgnoreCase))
                    s.PenColor = c;
            }
        }
        catch { /* 读失败就用默认值 */ }

        return s;
    }

    public void Save()
    {
        try
        {
            string? dir = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            File.WriteAllLines(FilePath, new[]
            {
                "# PPT 侧栏设置 —— 可以直接改这个文件",
                "# 颜色格式 #RRGGBB",
                "",
                "# 尺寸单位是【毫米】（不是像素）。因为同一串像素在不同屏幕上",
                "# 代表的物理大小能差 3 倍以上，毫米才能保证大小一致。",
                "# 想改更方便：改「启动侧栏.cmd」顶部的 WIDTH / HEIGHT（也是毫米）",
                $"# {KeyWidth} ：0 = 用默认宽度（38 mm）",
                $"# {KeyHeight}：0 = 贴合内容（5 个按钮，约 77 mm）",
                $"{KeyWidth}={Num(WidthMm)}",
                $"{KeyHeight}={Num(HeightMm)}",
                "",
                $"{KeyPen}={ToHex(PenColor)}",
            });
        }
        catch { /* 存不下来不算致命 */ }
    }

    private static string Num(double v) =>
        v.ToString("0.#", CultureInfo.InvariantCulture);

    private static bool TryParseDouble(string s, out double v) =>
        double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v);

    private static string ToHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

    private static bool TryParseHex(string s, out Color color)
    {
        color = Color.Empty;
        if (string.IsNullOrEmpty(s)) return false;

        if (s[0] == '#') s = s[1..];
        if (s.Length != 6) return false;
        if (!int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int v))
            return false;

        color = Color.FromArgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
        return true;
    }
}
