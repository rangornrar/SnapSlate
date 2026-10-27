using System;
using System.Drawing;
using System.IO;

namespace SnapSlate;

internal static class ImageSafety
{
    public const int MaxDimension = 16384;
    public const long MaxPixels = 64_000_000;
    public const int MaxImageBytes = 32 * 1024 * 1024;

    public static void ValidateDimensions(int width, int height)
    {
        if (width < 1 || height < 1 || width > MaxDimension || height > MaxDimension || (long)width * height > MaxPixels)
        {
            throw new InvalidDataException("Image dimensions exceed the supported limits.");
        }
    }

    public static void Validate(byte[] bytes)
    {
        if (bytes.Length == 0 || bytes.Length > MaxImageBytes)
        {
            throw new InvalidDataException("Image size exceeds the supported limits.");
        }

        using var stream = new MemoryStream(bytes, writable: false);
        using var image = Image.FromStream(stream);
        ValidateDimensions(image.Width, image.Height);
    }
}
