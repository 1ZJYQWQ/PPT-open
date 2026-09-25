using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PowerPointSidebar;

/// <summary>
/// 用 GDI+ 直接画图标，避免引入外部图片资源。
/// 每个图标都是 32×32 的 PNG，调用方再 SetToBitmap 拉伸到目标大小。
/// </summary>
internal static class IconFactory
{
    private const int SourceSize = 32;

    public static Bitmap Previous(Color color) => Arrow(color, mirror: false);  // ← 指向左
    public static Bitmap Next(Color color)     => Arrow(color, mirror: true);    // → 指向右

    public static Bitmap Pen(Color color)
    {
        var bmp = NewBitmap();
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = MakePen(color, 3f, LineCap.Round);

        // 笔身（对角线）
        g.DrawLine(pen, 6f, 26f, 24f, 8f);
        // 笔尖三角
        using var tipBrush = new SolidBrush(color);
        g.FillPolygon(tipBrush, new[]
        {
            new PointF(24f, 8f),
            new PointF(28f, 12f),
            new PointF(22f, 14f),
        });
        return bmp;
    }

    public static Bitmap Highlighter(Color color)
    {
        var bmp = NewBitmap();
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = MakePen(color, 6f, LineCap.Square);

        // 倾斜的荧光笔：粗斜杠代表笔头下的高亮带
        g.DrawLine(pen, 6f, 26f, 26f, 6f);
        // 外框
        using var outline = new Pen(Color.FromArgb(180, 50, 50, 50), 1.4f);
        g.DrawRectangle(outline, 5f, 5f, 22f, 22f);
        using var cap = new SolidBrush(color);
        g.FillRectangle(cap, 5f, 5f, 22f, 22f);
        // 高亮斜条
        using var stripe = new SolidBrush(Color.FromArgb(255, 255, 200, 0));
        g.FillRectangle(stripe, 8f, 22f, 18f, 4f);
        return bmp;
    }

    public static Bitmap Eraser(Color color)
    {
        var bmp = NewBitmap();
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        // 主体（粉色橡皮经典配色；颜色用 control 提供，方便整体改主题）
        using var body = new SolidBrush(Color.FromArgb(255, 220, 160, 180));
        using var band = new SolidBrush(Color.FromArgb(255, 110, 110, 130));
        using var outline = new Pen(color, 1.6f);
        // 矩形橡皮
        var rect = new RectangleF(5f, 8f, 22f, 18f);
        g.FillRectangle(body, rect);
        g.DrawRectangle(outline, rect.X, rect.Y, rect.Width, rect.Height);
        // 底部金属带
        g.FillRectangle(band, 5f, 19f, 22f, 7f);
        g.DrawRectangle(outline, 5f, 19f, 22f, 7f);
        return bmp;
    }

    public static Bitmap Exit(Color color)
    {
        var bmp = NewBitmap();
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = MakePen(color, 3.6f, LineCap.Round);
        g.DrawLine(pen, 10f, 10f, 22f, 22f);
        g.DrawLine(pen, 22f, 10f, 10f, 22f);
        return bmp;
    }

    // ---------- 内部工具 ----------

    private static Bitmap Arrow(Color color, bool mirror)
    {
        var bmp = NewBitmap();
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = MakePen(color, 3.4f, LineCap.Round);
        // 一个三角形（V 形）
        var left = mirror ? 8f : 24f;
        var right = mirror ? 24f : 8f;
        var midY = 16f;
        g.DrawLine(pen, left, 8f, right, midY);
        g.DrawLine(pen, right, midY, left, 24f);
        return bmp;
    }

    private static Bitmap NewBitmap()
    {
        var bmp = new Bitmap(SourceSize, SourceSize);
        bmp.SetResolution(96, 96);
        return bmp;
    }

    private static Pen MakePen(Color color, float width, LineCap cap) =>
        new(color, width)
        {
            StartCap = cap,
            EndCap = cap,
            LineJoin = LineJoin.Round,
        };
}