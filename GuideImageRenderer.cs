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
        using var bitmap = RenderDocument(document, palettes);
        AtomicFile.Write(outputPath, temporary => bitmap.Save(temporary, ImageFormat.Png), keepBackup: false);
    }

    public static byte[] RenderPng(ScreenshotDocument document, IReadOnlyDictionary<string, GradientPaletteDefinition> palettes)
    {
        using var bitmap = RenderDocument(document, palettes);
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private static Bitmap RenderDocument(ScreenshotDocument document, IReadOnlyDictionary<string, GradientPaletteDefinition> palettes)
    {
        using var sourceImage = LoadSourceImage(document.ImageBytes);
        var width = sourceImage?.Width ?? (document.SourcePixelWidth > 0 ? document.SourcePixelWidth : DefaultWidth);
        var height = sourceImage?.Height ?? (document.SourcePixelHeight > 0 ? document.SourcePixelHeight : DefaultHeight);

        var bitmap = new Bitmap(width, height, PixelFormat.Format32bppPArgb);
        bitmap.SetResolution(96, 96);

        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            if (sourceImage is not null)
            {
                graphics.DrawImage(sourceImage, new RectangleF(0, 0, width, height));
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

            foreach (var annotation in document.Annotations.Where(item => item.Kind != AnnotationKind.Mask)
                .Concat(document.Annotations.Where(item => item.Kind == AnnotationKind.Mask)))
            {
                DrawAnnotation(graphics, annotation, palettes, 0, 0);
            }

        }

        if (document.AppliedCropRect is Rect requestedCrop)
        {
            var crop = CropGeometry.Clamp(requestedCrop, width, height);
            var cropped = bitmap.Clone(new System.Drawing.Rectangle((int)crop.X, (int)crop.Y, (int)crop.Width, (int)crop.Height), PixelFormat.Format32bppPArgb);
            bitmap.Dispose();
            return cropped;
        }
        return bitmap;
    }

    private static Image? LoadSourceImage(byte[]? bytes)
    {
        if (bytes is not { Length: > 0 })
        {
            return null;
        }

        ImageSafety.Validate(bytes);
        using var stream = new MemoryStream(bytes, writable: false);
        using var image = Image.FromStream(stream);
        return new Bitmap(image);
    }

    private static void DrawAnnotation(Graphics graphics, AnnotationModel annotation, IReadOnlyDictionary<string, GradientPaletteDefinition> palettes, double offsetX, double offsetY)
    {
        var palette = palettes.TryGetValue(annotation.PaletteKey, out var found) ? found : palettes.Values.First();
        var shade = GetPaletteShade(palette, annotation.PaletteShadeIndex);
        var highlight = Blend(shade, palette.EndColor, 0.25);
        var opacity = ClampAlpha(annotation.Opacity);
        var strokeColor = WithAlpha(shade, opacity);
        var fillColor = WithAlpha(highlight, (int)Math.Round(opacity * (annotation.Kind == AnnotationKind.Focus ? 72 : annotation.Kind == AnnotationKind.Ellipse ? 30 : 36) / 255.0));

        switch (annotation.Kind)
        {
            case AnnotationKind.Text:
                DrawText(graphics, annotation, strokeColor, WithAlpha(highlight, opacity), offsetX, offsetY);
                break;
            case AnnotationKind.Sticker:
                DrawSticker(graphics, annotation, strokeColor, highlight, offsetX, offsetY);
                break;
            case AnnotationKind.Rectangle:
                DrawRoundedShape(graphics, annotation.Bounds, strokeColor, WithAlpha(highlight, opacity), fillColor, annotation.StrokeThickness, 24, offsetX, offsetY);
                break;
            case AnnotationKind.Ellipse:
                DrawEllipse(graphics, annotation.Bounds, strokeColor, WithAlpha(highlight, opacity), fillColor, annotation.StrokeThickness, offsetX, offsetY);
                break;
            case AnnotationKind.Focus:
                DrawRoundedShape(graphics, annotation.Bounds, strokeColor, WithAlpha(highlight, opacity), fillColor, Math.Max(3, annotation.StrokeThickness), 24, offsetX, offsetY);
                break;
            case AnnotationKind.Mask:
                DrawMask(graphics, annotation.Bounds, offsetX, offsetY);
                break;
            case AnnotationKind.ArrowStraight:
                DrawArrow(graphics, annotation.StartPoint, annotation.EndPoint, strokeColor, annotation.StrokeThickness, false, offsetX, offsetY);
                break;
            case AnnotationKind.ArrowCurved:
                DrawArrow(graphics, annotation.StartPoint, annotation.EndPoint, strokeColor, annotation.StrokeThickness, true, offsetX, offsetY);
                break;
        }
    }

    private static void DrawText(Graphics graphics, AnnotationModel annotation, Color strokeColor, Color highlightColor, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(annotation.Bounds.X - offsetX),
            (float)(annotation.Bounds.Y - offsetY),
            (float)Math.Max(220, annotation.Bounds.Width),
            (float)Math.Max(60, annotation.Bounds.Height));

        using var path = CreateRoundedRectanglePath(rect, 14);
        using var background = new SolidBrush(Color.FromArgb((int)Math.Round(235 * annotation.Opacity), 255, 255, 255));
        using var gradient = new LinearGradientBrush(rect, strokeColor, highlightColor, LinearGradientMode.ForwardDiagonal);
        using var pen = new Pen(gradient, 2);
        using var font = new Font("Segoe UI", (float)Math.Max(16, annotation.FontSize), annotation.IsBold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
        using var textBrush = new LinearGradientBrush(rect, strokeColor, highlightColor, LinearGradientMode.ForwardDiagonal);
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
        using var outline = new Pen(Color.FromArgb(strokeColor.A, 255, 255, 255), 3);
        using var font = new Font("Segoe UI Semibold", 24, FontStyle.Bold, GraphicsUnit.Pixel);
        using var textBrush = new SolidBrush(Color.FromArgb(strokeColor.A, 255, 255, 255));
        using var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

        graphics.FillEllipse(brush, rect);
        graphics.DrawEllipse(outline, rect);
        graphics.DrawString(annotation.Text, font, textBrush, rect, format);
    }

    private static void DrawRoundedShape(Graphics graphics, Rect bounds, Color strokeColor, Color highlightColor, Color fillColor, double thickness, int radius, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(bounds.X - offsetX),
            (float)(bounds.Y - offsetY),
            (float)Math.Max(1, bounds.Width),
            (float)Math.Max(1, bounds.Height));

        using var path = CreateRoundedRectanglePath(rect, radius);
        using var fill = new SolidBrush(fillColor);
        using var gradient = new LinearGradientBrush(rect, strokeColor, highlightColor, LinearGradientMode.ForwardDiagonal);
        using var pen = new Pen(gradient, (float)Math.Max(1, thickness));

        graphics.FillPath(fill, path);
        graphics.DrawPath(pen, path);
    }

    private static void DrawEllipse(Graphics graphics, Rect bounds, Color strokeColor, Color highlightColor, Color fillColor, double thickness, double offsetX, double offsetY)
    {
        var rect = new RectangleF(
            (float)(bounds.X - offsetX),
            (float)(bounds.Y - offsetY),
            (float)Math.Max(1, bounds.Width),
            (float)Math.Max(1, bounds.Height));

        using var fill = new SolidBrush(fillColor);
        using var gradient = new LinearGradientBrush(rect, strokeColor, highlightColor, LinearGradientMode.ForwardDiagonal);
        using var pen = new Pen(gradient, (float)Math.Max(1, thickness));

        graphics.FillEllipse(fill, rect);
        graphics.DrawEllipse(pen, rect);
    }

    private static void DrawMask(Graphics graphics, Rect bounds, double offsetX, double offsetY)
    {
        // Cover every pixel touching the selection, without antialiased/rounded edges.
        var left = (int)Math.Floor(bounds.X - offsetX);
        var top = (int)Math.Floor(bounds.Y - offsetY);
        var right = (int)Math.Ceiling(bounds.Right - offsetX);
        var bottom = (int)Math.Ceiling(bounds.Bottom - offsetY);
        var smoothing = graphics.SmoothingMode;
        graphics.SmoothingMode = SmoothingMode.None;
        using var fill = new SolidBrush(Color.FromArgb(255, 18, 18, 18));
        graphics.FillRectangle(fill, left, top, Math.Max(1, right - left), Math.Max(1, bottom - top));
        graphics.SmoothingMode = smoothing;
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
