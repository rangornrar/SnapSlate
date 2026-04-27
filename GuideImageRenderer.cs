using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using DrawingColor = System.Drawing.Color;
using FoundationPoint = Windows.Foundation.Point;
using Windows.Foundation;

namespace SnapSlate;

public static class GuideImageRenderer
{
    private const int DefaultWidth = 1360;
    private const int DefaultHeight = 820;
    private const int StickerSize = 60;

    public static void SaveDocumentImage(ScreenshotDocument document, IReadOnlyDictionary<string, GradientPaletteDefinition> palettes, string outputPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        using var bitmap = RenderDocument(document, palettes);
        bitmap.Save(outputPath, ImageFormat.Png);
    }

    private static Bitmap RenderDocument(ScreenshotDocument document, IReadOnlyDictionary<string, GradientPaletteDefinition> palettes)
    {
        var sourceImage = LoadSourceImage(document.ImageBytes);
        var crop = GetExportCrop(document, sourceImage);
        var width = Math.Max(1, (int)Math.Round(crop.Width > 0 ? crop.Width : sourceImage?.Width ?? DefaultWidth));
        var height = Math.Max(1, (int)Math.Round(crop.Height > 0 ? crop.Height : sourceImage?.Height ?? DefaultHeight));

        var bitmap = new Bitmap(width, height, PixelFormat.Format32bppPArgb);
        bitmap.SetResolution(96, 96);

        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        if (sourceImage is not null)
        {
            if (crop.Width > 0 && crop.Height > 0)
            {
                graphics.DrawImage(
                    sourceImage,
                    new RectangleF(0, 0, width, height),
                    new RectangleF((float)crop.X, (float)crop.Y, (float)crop.Width, (float)crop.Height),
                    GraphicsUnit.Pixel);
            }
            else
            {
                graphics.DrawImage(sourceImage, new RectangleF(0, 0, width, height));
            }
        }
        else
        {
            using var background = new LinearGradientBrush(
                new RectangleF(0, 0, width, height),
                Color.FromArgb(255, 248, 250, 252),
                Color.FromArgb(255, 232, 238, 246),
                LinearGradientMode.ForwardDiagonal);
            graphics.FillRectangle(background, 0, 0, width, height);

            using var titleFont = new Font("Segoe UI Semibold", 28, FontStyle.Bold, GraphicsUnit.Pixel);
            using var bodyFont = new Font("Segoe UI", 16, FontStyle.Regular, GraphicsUnit.Pixel);
            using var titleBrush = new SolidBrush(Color.FromArgb(255, 32, 41, 60));
            using var bodyBrush = new SolidBrush(Color.FromArgb(255, 84, 99, 126));
            graphics.DrawString(document.BaseTitle, titleFont, titleBrush, new PointF(36, 28));
            graphics.DrawString(document.StepNote?.Trim().Length > 0 ? document.StepNote : "SnapSlate", bodyFont, bodyBrush, new PointF(36, 76));
        }

        var cropOffsetX = crop.Width > 0 ? crop.X : 0;
        var cropOffsetY = crop.Height > 0 ? crop.Y : 0;

        foreach (var annotation in document.Annotations)
        {
            DrawAnnotation(graphics, annotation, palettes, cropOffsetX, cropOffsetY);
        }

        sourceImage?.Dispose();
        return bitmap;
    }

    private static Image? LoadSourceImage(byte[]? bytes)
    {
        if (bytes is not { Length: > 0 })
        {
            return null;
        }

        using var stream = new MemoryStream(bytes, writable: false);
        try
        {
            using var image = Image.FromStream(stream);
            return new Bitmap(image);
        }
        catch
        {
            return null;
        }
    }

    private static Rect GetExportCrop(ScreenshotDocument document, Image? sourceImage)
    {
        var crop = document.AppliedCropRect ?? document.PendingCropRect;
        if (crop is null || crop.Value.Width < 1 || crop.Value.Height < 1)
        {
            if (sourceImage is null)
            {
                return new Rect(0, 0, DefaultWidth, DefaultHeight);
            }

            return new Rect(0, 0, sourceImage.Width, sourceImage.Height);
        }

        var x = Math.Max(0, crop.Value.X);
        var y = Math.Max(0, crop.Value.Y);
        var width = Math.Max(1, crop.Value.Width);
        var height = Math.Max(1, crop.Value.Height);

        if (sourceImage is not null)
        {
            width = Math.Min(width, sourceImage.Width - x);
            height = Math.Min(height, sourceImage.Height - y);
        }

        return new Rect(x, y, Math.Max(1, width), Math.Max(1, height));
    }

    private static void DrawAnnotation(Graphics graphics, AnnotationModel annotation, IReadOnlyDictionary<string, GradientPaletteDefinition> palettes, double offsetX, double offsetY)
    {
        var palette = palettes.TryGetValue(annotation.PaletteKey, out var found) ? found : palettes.Values.First();
        var shade = GetPaletteShade(palette, annotation.PaletteShadeIndex);
        var highlight = Blend(shade, palette.EndColor, 0.25);
        var opacity = ClampAlpha(annotation.Opacity);
        var strokeColor = WithAlpha(shade, opacity);
        var fillColor = WithAlpha(highlight, Math.Max(60, (int)Math.Round(opacity * 0.22)));

        switch (annotation.Kind)
        {
            case AnnotationKind.Text:
                DrawText(graphics, annotation, strokeColor, fillColor, offsetX, offsetY);
                break;
            case AnnotationKind.Sticker:
                DrawSticker(graphics, annotation, strokeColor, highlight, offsetX, offsetY);
                break;
            case AnnotationKind.Rectangle:
                DrawRoundedShape(graphics, annotation.Bounds, strokeColor, fillColor, annotation.StrokeThickness, 24, offsetX, offsetY);
                break;
            case AnnotationKind.Ellipse:
                DrawEllipse(graphics, annotation.Bounds, strokeColor, fillColor, annotation.StrokeThickness, offsetX, offsetY);
                break;
            case AnnotationKind.Focus:
                DrawRoundedShape(graphics, annotation.Bounds, strokeColor, fillColor, Math.Max(3, annotation.StrokeThickness), 24, offsetX, offsetY);
                break;
            case AnnotationKind.Mask:
                DrawMask(graphics, annotation.Bounds, opacity, offsetX, offsetY);
                break;
            case AnnotationKind.ArrowStraight:
                DrawArrow(graphics, annotation.StartPoint, annotation.EndPoint, strokeColor, annotation.StrokeThickness, false, offsetX, offsetY);
                break;
            case AnnotationKind.ArrowCurved:
                DrawArrow(graphics, annotation.StartPoint, annotation.EndPoint, strokeColor, annotation.StrokeThickness, true, offsetX, offsetY);
                break;
        }
    }

    private static void DrawText(Graphics graphics, AnnotationModel annotation, Color strokeColor, Color fillColor, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(annotation.Bounds.X - offsetX),
            (float)(annotation.Bounds.Y - offsetY),
            (float)Math.Max(220, annotation.Bounds.Width),
            (float)Math.Max(60, annotation.Bounds.Height));

        using var path = CreateRoundedRectanglePath(rect, 14);
        using var background = new SolidBrush(Color.FromArgb(Math.Max(220, (int)fillColor.A), 255, 255, 255));
        using var pen = new Pen(strokeColor, 2);
        using var font = new Font("Segoe UI", (float)Math.Max(16, annotation.FontSize), annotation.IsBold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
        using var textBrush = new SolidBrush(strokeColor);
        using var format = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

        graphics.FillPath(background, path);
        graphics.DrawPath(pen, path);
        graphics.DrawString(annotation.Text, font, textBrush, rect.InflateCopy(-12, -8), format);
    }

    private static void DrawSticker(Graphics graphics, AnnotationModel annotation, Color strokeColor, Color highlight, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(annotation.Bounds.X - offsetX),
            (float)(annotation.Bounds.Y - offsetY),
            (float)Math.Max(StickerSize, annotation.Bounds.Width),
            (float)Math.Max(StickerSize, annotation.Bounds.Height));

        using var brush = new LinearGradientBrush(rect, strokeColor, WithAlpha(highlight, strokeColor.A), LinearGradientMode.ForwardDiagonal);
        using var outline = new Pen(Color.FromArgb(Math.Min(255, strokeColor.A + 25), 255, 255, 255), 3);
        using var font = new Font("Segoe UI Semibold", 24, FontStyle.Bold, GraphicsUnit.Pixel);
        using var textBrush = new SolidBrush(Color.White);
        using var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

        graphics.FillEllipse(brush, rect);
        graphics.DrawEllipse(outline, rect);
        graphics.DrawString(annotation.Text, font, textBrush, rect, format);
    }

    private static void DrawRoundedShape(Graphics graphics, Rect bounds, Color strokeColor, Color fillColor, double thickness, int radius, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(bounds.X - offsetX),
            (float)(bounds.Y - offsetY),
            (float)Math.Max(1, bounds.Width),
            (float)Math.Max(1, bounds.Height));

        using var path = CreateRoundedRectanglePath(rect, radius);
        using var fill = new SolidBrush(fillColor);
        using var pen = new Pen(strokeColor, (float)Math.Max(1, thickness));

        graphics.FillPath(fill, path);
        graphics.DrawPath(pen, path);
    }

    private static void DrawEllipse(Graphics graphics, Rect bounds, Color strokeColor, Color fillColor, double thickness, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(bounds.X - offsetX),
            (float)(bounds.Y - offsetY),
            (float)Math.Max(1, bounds.Width),
            (float)Math.Max(1, bounds.Height));

        using var fill = new SolidBrush(fillColor);
        using var pen = new Pen(strokeColor, (float)Math.Max(1, thickness));

        graphics.FillEllipse(fill, rect);
        graphics.DrawEllipse(pen, rect);
    }

    private static void DrawMask(Graphics graphics, Rect bounds, int opacity, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(bounds.X - offsetX),
            (float)(bounds.Y - offsetY),
            (float)Math.Max(1, bounds.Width),
            (float)Math.Max(1, bounds.Height));

        using var path = CreateRoundedRectanglePath(rect, 20);
        using var fill = new SolidBrush(Color.FromArgb(Math.Max(150, opacity), 18, 18, 18));
        using var pen = new Pen(Color.FromArgb(255, 18, 18, 18), 2);
        graphics.FillPath(fill, path);
        graphics.DrawPath(pen, path);
    }

    private static void DrawArrow(Graphics graphics, FoundationPoint start, FoundationPoint end, Color strokeColor, double thickness, bool curved, double offsetX, double offsetY)
    {
        var startPoint = new PointF((float)(start.X - offsetX), (float)(start.Y - offsetY));
        var endPoint = new PointF((float)(end.X - offsetX), (float)(end.Y - offsetY));

        using var pen = new Pen(strokeColor, (float)Math.Max(1, thickness))
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };

        if (curved)
        {
            var control1 = new PointF(startPoint.X + ((endPoint.X - startPoint.X) * 0.3f), startPoint.Y - 80);
            var control2 = new PointF(startPoint.X + ((endPoint.X - startPoint.X) * 0.82f), endPoint.Y - 24);
            graphics.DrawBezier(pen, startPoint, control1, control2, endPoint);
            DrawArrowHead(graphics, strokeColor, thickness, control2, endPoint);
            return;
        }

        graphics.DrawLine(pen, startPoint, endPoint);
        DrawArrowHead(graphics, strokeColor, thickness, startPoint, endPoint);
    }

    private static void DrawArrowHead(Graphics graphics, Color strokeColor, double thickness, PointF from, PointF to)
    {
        var angle = Math.Atan2(to.Y - from.Y, to.X - from.X);
        var length = Math.Max(18, thickness * 4);
        var spread = Math.PI / 7;
        var left = new PointF(
            to.X - (float)(length * Math.Cos(angle - spread)),
            to.Y - (float)(length * Math.Sin(angle - spread)));
        var right = new PointF(
            to.X - (float)(length * Math.Cos(angle + spread)),
            to.Y - (float)(length * Math.Sin(angle + spread)));

        using var pen = new Pen(strokeColor, (float)Math.Max(1, thickness));
        graphics.DrawLine(pen, to, left);
        graphics.DrawLine(pen, to, right);
    }

    private static DrawingColor GetPaletteShade(GradientPaletteDefinition palette, int shadeIndex)
    {
        var index = Math.Clamp(shadeIndex, 0, Math.Max(0, palette.Shades.Count - 1));
        var color = palette.Shades.Count > 0 ? palette.Shades[index] : palette.StartColor;
        return DrawingColor.FromArgb(color.A, color.R, color.G, color.B);
    }

    private static DrawingColor Blend(DrawingColor start, Windows.UI.Color end, double factor)
    {
        var endColor = DrawingColor.FromArgb(end.A, end.R, end.G, end.B);
        return DrawingColor.FromArgb(
            255,
            (byte)Math.Round(start.R + ((endColor.R - start.R) * factor)),
            (byte)Math.Round(start.G + ((endColor.G - start.G) * factor)),
            (byte)Math.Round(start.B + ((endColor.B - start.B) * factor)));
    }

    private static DrawingColor WithAlpha(DrawingColor color, int alpha)
    {
        return DrawingColor.FromArgb(Math.Clamp(alpha, 0, 255), color.R, color.G, color.B);
    }

    private static int ClampAlpha(double opacity)
    {
        return Math.Clamp((int)Math.Round(opacity * 255), 0, 255);
    }

    private static GraphicsPath CreateRoundedRectanglePath(RectangleF rect, float radius)
    {
        var path = new GraphicsPath();
        var diameter = radius * 2;
        var arc = new RectangleF(rect.Location, new SizeF(diameter, diameter));

        path.AddArc(arc, 180, 90);
        arc.X = rect.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = rect.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = rect.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        return path;
    }

    private static RectangleF InflateCopy(this RectangleF rect, float horizontal, float vertical)
    {
        rect.Inflate(horizontal, vertical);
        return rect;
    }
}
