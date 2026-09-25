using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PowerPointSidebar;

/// <summary>
/// 用 GDI+ 直接画图标，避免引入外部图片资源。
///
/// ★ 关键：所有图标都按 32×32 的【逻辑坐标系】作画，最后整体缩放到
///   SourceSize（128）输出。这样图标在放大到 100+ 物理像素时依然锐利 ——
///   直接拉伸 32px 位图会糊成一团，大屏上尤其明显。
/// </summary>
internal static class IconFactory
{
    /// <summary>作画用的逻辑坐标系边长（写图标时所有坐标按 0~32 来）。</summary>
    private const float Art = 32f;

    /// <summary>输出位图边长（4× 超采样）。调用方再缩放到目标尺寸。</summary>
    private const int SourceSize = 128;

    public static Bitmap Previous(Color color) => Arrow(color, mirror: false);  // ← 指向左
    public static Bitmap Next(Color color)     => Arrow(color, mirror: true);    // → 指向右

    public static Bitmap Pen(Color color)
    {
        using var g = Begin(out var bmp);

        // 笔身（对角线）
        using var body = MakePen(color, 3.2f, LineCap.Round);
        g.DrawLine(body, 8f, 24f, 23f, 9f);

        // 笔尖（实心三角）
        using var tip = new SolidBrush(color);
        g.FillPolygon(tip, new[]
        {
            new PointF(23f, 9f),
            new PointF(28f, 14f),
            new PointF(20.5f, 16.5f),
        });

        // 笔尾小方块，让笔看起来更"实"
        using var tail = new SolidBrush(color);
        g.FillRectangle(tail, 5.5f, 22.5f, 5f, 5f);

        return bmp;
    }

    public static Bitmap Eraser(Color color)
    {
        using var g = Begin(out var bmp);

        // 橡皮主体（两段式：上白下彩，经典橡皮造型）
        using var top = new SolidBrush(Color.FromArgb(255, 245, 210, 222));
        using var band = new SolidBrush(Color.FromArgb(255, 120, 126, 142));
        using var outline = new Pen(color, 1.8f) { LineJoin = LineJoin.Round };

        var rect = new RectangleF(6f, 8f, 20f, 17f);
        g.FillRectangle(top, rect);
        g.FillRectangle(band, new RectangleF(6f, 18f, 20f, 7f));
        g.DrawRectangle(outline, rect.X, rect.Y, rect.Width, rect.Height);
        g.DrawLine(outline, 6f, 18f, 26f, 18f);

        return bmp;
    }

    /// <summary>
    /// 「结束放映」：圆角实心方块 = 通用的「停止」语义。
    /// （原来画的是 ✕，容易被误读成"关闭侧栏"，和真正的关闭按钮撞语义。）
    /// </summary>
    public static Bitmap Exit(Color color)
    {
        using var g = Begin(out var bmp);
        var r = new RectangleF(8f, 8f, 16f, 16f);
        using var path = Rounded(r, 3.2f);
        using var brush = new SolidBrush(color);
        g.FillPath(brush, path);
        return bmp;
    }

    // ---------- 内部工具 ----------

    private static Bitmap Arrow(Color color, bool mirror)
    {
        using var g = Begin(out var bmp);
        using var pen = MakePen(color, 3.6f, LineCap.Round);

        // 尖端（tip）与开口（back）：向上/向下两条线在 tip 汇成尖角
        float tip  = mirror ? 21f : 11f;
        float back = mirror ? 11f : 21f;
        g.DrawLine(pen, back, 7f, tip, 16f);
        g.DrawLine(pen, tip, 16f, back, 25f);

        return bmp;
    }

    /// <summary>建好画布并把坐标系缩放到 32 逻辑单位。</summary>
    private static Graphics Begin(out Bitmap bmp)
    {
        bmp = new Bitmap(SourceSize, SourceSize);
        bmp.SetResolution(96, 96);
        var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        float s = SourceSize / Art;
        g.ScaleTransform(s, s);
        return g;
    }

    private static GraphicsPath Rounded(RectangleF r, float radius)
    {
        float d = radius * 2;
        var p = new GraphicsPath();
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    private static Pen MakePen(Color color, float width, LineCap cap) =>
        new(color, width)
        {
            StartCap = cap,
            EndCap = cap,
            LineJoin = LineJoin.Round,
        };
}
