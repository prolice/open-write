using System.Runtime.InteropServices;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace OpenWrite.App.Layout;

/// <summary>
/// Images EMF et WMF (aperçus d'objets OLE, équations de l'Éditeur d'équations 3.0, images clipart des documents
/// anciens) : WPF ne les décode pas ; elles sont jouées par GDI dans une image bitmap (fond blanc, 192 points par
/// pouce, plus grand côté borné), puis affichées comme une image ordinaire.
/// </summary>
static class Metafiles
{
    /// <summary>Le contenu est un EMF (signature « EMF » à l'octet 40) ou un WMF (en-tête « placeable » ou en-tête standard).</summary>
    public static bool IsMetafile(byte[] d) =>
        d.Length > 44 && BitConverter.ToUInt32(d, 0) == 1 && BitConverter.ToUInt32(d, 40) == 0x464D4520
        || d.Length > 22 && BitConverter.ToUInt32(d, 0) == 0x9AC6CDD7
        || d.Length > 18 && BitConverter.ToUInt16(d, 0) is 1 or 2 && BitConverter.ToUInt16(d, 2) == 9;

    public static BitmapSource? Render(byte[] data, double dpi = 192, int maxSide = 2400)
    {
        IntPtr emf = IntPtr.Zero;
        try
        {
            double wMm, hMm;
            if (BitConverter.ToUInt32(data, 0) == 1 && data.Length > 44 && BitConverter.ToUInt32(data, 40) == 0x464D4520)
            {
                emf = SetEnhMetaFileBits((uint)data.Length, data);
                if (emf == IntPtr.Zero) return null;
                var hdr = new ENHMETAHEADER();
                GetEnhMetaFileHeader(emf, (uint)Marshal.SizeOf<ENHMETAHEADER>(), ref hdr);
                // cadre en centièmes de millimètre
                wMm = (hdr.rclFrame.Right - hdr.rclFrame.Left) / 100.0;
                hMm = (hdr.rclFrame.Bottom - hdr.rclFrame.Top) / 100.0;
            }
            else
            {
                // WMF : en-tête « placeable » (taille, unités par pouce) retiré ; sinon taille par défaut
                byte[] bits = data;
                wMm = 50; hMm = 50;
                if (BitConverter.ToUInt32(data, 0) == 0x9AC6CDD7)
                {
                    short l = BitConverter.ToInt16(data, 6), t = BitConverter.ToInt16(data, 8), r = BitConverter.ToInt16(data, 10), b = BitConverter.ToInt16(data, 12);
                    ushort inch = BitConverter.ToUInt16(data, 14);
                    if (inch == 0) inch = 1440;
                    wMm = Math.Abs(r - l) * 25.4 / inch;
                    hMm = Math.Abs(b - t) * 25.4 / inch;
                    bits = data[22..];
                }
                var mfp = new METAFILEPICT { mm = 8, xExt = (int)(wMm * 100), yExt = (int)(hMm * 100), hMF = IntPtr.Zero };   // MM_ANISOTROPIC
                emf = SetWinMetaFileBits((uint)bits.Length, bits, IntPtr.Zero, ref mfp);
                if (emf == IntPtr.Zero) return null;
            }
            if (wMm <= 0.1 || hMm <= 0.1) return null;
            int w = (int)Math.Round(wMm / 25.4 * dpi), h = (int)Math.Round(hMm / 25.4 * dpi);
            double k = Math.Min(1, maxSide / (double)Math.Max(w, h));
            w = Math.Max(1, (int)(w * k));
            h = Math.Max(1, (int)(h * k));

            var screen = GetDC(IntPtr.Zero);
            var dc = CreateCompatibleDC(screen);
            ReleaseDC(IntPtr.Zero, screen);
            var bmi = new BITMAPINFO { biSize = 40, biWidth = w, biHeight = -h, biPlanes = 1, biBitCount = 32 };
            var dib = CreateDIBSection(dc, ref bmi, 0, out var pixels, IntPtr.Zero, 0);
            if (dib == IntPtr.Zero) { DeleteDC(dc); return null; }
            var old = SelectObject(dc, dib);
            try
            {
                // fond blanc, puis l'image
                int stride = w * 4;
                var white = new byte[stride * h];
                Array.Fill(white, (byte)0xFF);
                Marshal.Copy(white, 0, pixels, white.Length);
                var rect = new RECT { Left = 0, Top = 0, Right = w, Bottom = h };
                PlayEnhMetaFile(dc, emf, ref rect);
                GdiFlush();
                var buffer = new byte[stride * h];
                Marshal.Copy(pixels, buffer, 0, buffer.Length);
                for (int i = 3; i < buffer.Length; i += 4) buffer[i] = 0xFF;          // GDI ne remplit pas l'alpha
                var src = BitmapSource.Create(w, h, dpi * k, dpi * k, PixelFormats.Bgra32, null, buffer, stride);
                src.Freeze();
                return src;
            }
            finally
            {
                SelectObject(dc, old);
                DeleteObject(dib);
                DeleteDC(dc);
            }
        }
        catch (Exception ex) when (ex is ExternalException or ArgumentException or IndexOutOfRangeException or OutOfMemoryException) { return null; }
        finally { if (emf != IntPtr.Zero) DeleteEnhMetaFile(emf); }
    }

    [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] struct SIZE { public int Cx, Cy; }
    [StructLayout(LayoutKind.Sequential)]
    struct ENHMETAHEADER
    {
        public uint iType, nSize;
        public RECT rclBounds, rclFrame;
        public uint dSignature, nVersion, nBytes, nRecords;
        public ushort nHandles, sReserved;
        public uint nDescription, offDescription, nPalEntries;
        public SIZE szlDevice, szlMillimeters;
        public uint cbPixelFormat, offPixelFormat, bOpenGL;
        public SIZE szlMicrometers;
    }
    [StructLayout(LayoutKind.Sequential)] struct METAFILEPICT { public int mm, xExt, yExt; public IntPtr hMF; }
    [StructLayout(LayoutKind.Sequential)]
    struct BITMAPINFO
    {
        public int biSize, biWidth, biHeight;
        public short biPlanes, biBitCount;
        public int biCompression, biSizeImage, biXPelsPerMeter, biYPelsPerMeter, biClrUsed, biClrImportant;
        public int bmiColors;
    }

    [DllImport("gdi32.dll")] static extern IntPtr SetEnhMetaFileBits(uint size, byte[] data);
    [DllImport("gdi32.dll")] static extern IntPtr SetWinMetaFileBits(uint size, byte[] data, IntPtr hdcRef, ref METAFILEPICT mfp);
    [DllImport("gdi32.dll")] static extern uint GetEnhMetaFileHeader(IntPtr hemf, uint size, ref ENHMETAHEADER header);
    [DllImport("gdi32.dll")] static extern bool PlayEnhMetaFile(IntPtr hdc, IntPtr hemf, ref RECT rect);
    [DllImport("gdi32.dll")] static extern bool DeleteEnhMetaFile(IntPtr hemf);
    [DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr hdc);
    [DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr hdc);
    [DllImport("gdi32.dll")] static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPINFO bmi, uint usage, out IntPtr bits, IntPtr section, uint offset);
    [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);
    [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr obj);
    [DllImport("gdi32.dll")] static extern bool GdiFlush();
    [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr hwnd);
    [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);
}
