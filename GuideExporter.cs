using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using Windows.Graphics.Imaging;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace SnapSlate;

public static class GuideExporter
{
    public static Task ExportPdfAsync(GuideExportManifest manifest, string outputPath)
    {
        return Task.Run(() =>
        {
            QuestPDF.Settings.License = LicenseType.Community;
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            QuestPDF.Fluent.Document.Create(document =>
            {
                document.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(32);
                    page.DefaultTextStyle(style => style.FontFamily("Segoe UI").FontSize(10).FontColor(Colors.Grey.Darken3));

                    page.Header().PaddingBottom(8).Row(row =>
                    {
                        row.RelativeItem().Text(manifest.Title).SemiBold().FontSize(12);
                        row.AutoItem().Text(manifest.TemplateName).FontSize(9).FontColor(Colors.Grey.Darken1);
                    });

                    page.Content().Column(column =>
                    {
                        column.Spacing(10);
                        column.Item().Text(manifest.Title).FontSize(24).SemiBold().FontColor(Colors.Blue.Darken3);
                        column.Item().Text($"{manifest.TemplateName} · {manifest.Audience} · {manifest.Version}").FontSize(11).FontColor(Colors.Grey.Darken1);
                        column.Item().Text($"Auteur : {manifest.Author} · Généré le {manifest.GeneratedAt.ToString("g", CultureInfo.CurrentCulture)}").FontSize(10);
                        column.Item().LineHorizontal(1).LineColor(Colors.Grey.Lighten2);
                        column.Item().Text($"Étapes : {manifest.Steps.Count}").FontSize(11).SemiBold();
                    });

                    page.Footer().PaddingTop(8).AlignCenter().Text($"SnapSlate · {manifest.Title}").FontSize(9).FontColor(Colors.Grey.Darken1);
                });

                foreach (var step in manifest.Steps)
                {
                    document.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(32);
                        page.DefaultTextStyle(style => style.FontFamily("Segoe UI").FontSize(10).FontColor(Colors.Grey.Darken3));

                        page.Header().PaddingBottom(8).Row(row =>
                        {
                            row.RelativeItem().Text(manifest.Title).SemiBold().FontSize(11);
                            row.AutoItem().Text($"{step.Index:00}").FontSize(9).FontColor(Colors.Grey.Darken1);
                        });

                        page.Content().Column(column =>
                        {
                            column.Spacing(10);
                            column.Item().Text($"{step.Index:00}. {step.Title}").FontSize(20).SemiBold().FontColor(Colors.Blue.Darken3);

                            if (!string.IsNullOrWhiteSpace(step.Note))
                            {
                                column.Item().PaddingBottom(2).Text(step.Note).FontSize(11);
                            }

                            column.Item().PaddingTop(4).Border(1).BorderColor(Colors.Grey.Lighten2).Padding(8).Image(step.ImagePath).FitArea();

                            if (step.LegendItems.Count > 0)
                            {
                                column.Item().PaddingTop(6).Text("Légende").SemiBold().FontSize(12);
                                foreach (var item in step.LegendItems)
                                {
                                    column.Item().PaddingLeft(8).Text($"• {item.StickerLabel} — {item.Description}");
                                }
                            }

                            column.Item().PaddingTop(6).Text(step.SourceLabel).FontSize(9).FontColor(Colors.Grey.Darken1);
                        });

                        page.Footer().PaddingTop(8).AlignRight().Text($"SnapSlate · {step.Index:00}").FontSize(9).FontColor(Colors.Grey.Darken1);
                    });
                }
            })
            .WithMetadata(new DocumentMetadata
            {
                Title = manifest.Title,
                Author = manifest.Author,
                Subject = manifest.TemplateName,
                Keywords = $"SnapSlate,{manifest.TemplateName},{manifest.Audience}",
                Creator = "SnapSlate",
                Language = "fr-FR",
                CreationDate = manifest.GeneratedAt,
                ModifiedDate = manifest.GeneratedAt
            })
            .GeneratePdf(outputPath);
        });
    }

    public static Task ExportDocxAsync(GuideExportManifest manifest, string outputPath)
    {
        return Task.Run(async () =>
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            using var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
            document.PackageProperties.Title = manifest.Title;
            document.PackageProperties.Creator = manifest.Author;
            document.PackageProperties.LastModifiedBy = manifest.Author;
            document.PackageProperties.Subject = manifest.TemplateName;
            document.PackageProperties.Keywords = $"SnapSlate,{manifest.TemplateName},{manifest.Audience}";

            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());
            var body = mainPart.Document.Body!;

            body.Append(
                CreateHeadingParagraph(manifest.Title, 1, 26),
                CreateTextParagraph($"{manifest.TemplateName} · {manifest.Audience} · {manifest.Version}", 11, false),
                CreateTextParagraph($"Auteur : {manifest.Author} · Généré le {manifest.GeneratedAt.ToString("g", CultureInfo.CurrentCulture)}", 10, false),
                CreateTextParagraph($"Étapes : {manifest.Steps.Count}", 10, true));

            foreach (var step in manifest.Steps)
            {
                body.Append(
                    CreateSpacerParagraph(1),
                    CreateHeadingParagraph($"{step.Index:00}. {step.Title}", 2, 20));

                if (!string.IsNullOrWhiteSpace(step.Note))
                {
                    body.Append(CreateTextParagraph(step.Note, 11, false));
                }

                var (pixelWidth, pixelHeight) = await GetImageSizeAsync(step.ImagePath);
                var imageWidthEmu = 5_750_000L;
                var imageHeightEmu = Math.Max(1L, (long)Math.Round(imageWidthEmu * ((double)pixelHeight / Math.Max(1, pixelWidth))));

                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var stream = File.OpenRead(step.ImagePath))
                {
                    imagePart.FeedData(stream);
                }

                var drawing = CreateImageDrawing(mainPart.GetIdOfPart(imagePart), Path.GetFileName(step.ImagePath), imageWidthEmu, imageHeightEmu, (uint)(step.Index + 1));
                body.Append(new Paragraph(new Run(drawing)));

                if (step.LegendItems.Count > 0)
                {
                    body.Append(CreateHeadingParagraph("Légende", 3, 14));
                    foreach (var item in step.LegendItems)
                    {
                        body.Append(CreateTextParagraph($"• {item.StickerLabel} — {item.Description}", 10, false));
                    }
                }

                body.Append(CreateTextParagraph(step.SourceLabel, 9, false));
            }

            body.Append(new SectionProperties());
            mainPart.Document.Save();
        });
    }

    private static Paragraph CreateHeadingParagraph(string text, int headingLevel, int fontSize)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        if (headingLevel <= 2)
        {
            properties.SpacingBetweenLines = new SpacingBetweenLines
            {
                After = headingLevel == 1 ? "240" : "160"
            };
        }

        paragraph.Append(properties);
        paragraph.Append(CreateRun(text, fontSize, true));
        return paragraph;
    }

    private static Paragraph CreateTextParagraph(string text, int fontSize, bool bold)
    {
        var paragraph = new Paragraph();
        paragraph.Append(new ParagraphProperties
        {
            SpacingBetweenLines = new SpacingBetweenLines { After = "120" }
        });
        paragraph.Append(CreateRun(text, fontSize, bold));
        return paragraph;
    }

    private static Paragraph CreateSpacerParagraph(int points)
    {
        var paragraph = new Paragraph();
        paragraph.Append(new ParagraphProperties
        {
            SpacingBetweenLines = new SpacingBetweenLines { After = (points * 20).ToString(CultureInfo.InvariantCulture) }
        });
        paragraph.Append(new Run(new Text(string.Empty)));
        return paragraph;
    }

    private static Run CreateRun(string text, int fontSize, bool bold)
    {
        var run = new Run();
        var runProperties = new RunProperties
        {
            FontSize = new FontSize { Val = (fontSize * 2).ToString(CultureInfo.InvariantCulture) }
        };

        if (bold)
        {
            runProperties.Bold = new Bold();
        }

        run.Append(runProperties);
        var normalizedText = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalizedText.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                run.Append(new Break());
            }

            run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
        }

        return run;
    }

    private static Drawing CreateImageDrawing(string relationshipId, string fileName, long cx, long cy, uint docPropsId)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties { Id = docPropsId, Name = fileName },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = fileName },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                EditId = "50D07946"
            });
    }

    private static async Task<(uint Width, uint Height)> GetImageSizeAsync(string path)
    {
        using var stream = File.OpenRead(path);
        using var randomAccess = stream.AsRandomAccessStream();
        var decoder = await BitmapDecoder.CreateAsync(randomAccess);
        return (decoder.PixelWidth, decoder.PixelHeight);
    }
}
