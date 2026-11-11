// Win32 helpers used by the watchdog: find dialogs owned by the Access process,
// read their text, and press a button so automation never hangs on a modal dialog.
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessVsc
{
    public class DialogButton
    {
        public IntPtr Hwnd;
        public int Id;
        public string Text;
    }

    public class DialogInfo
    {
        public IntPtr Hwnd;
        public string Title;
        public List<string> Texts = new List<string>();
        public List<DialogButton> Buttons = new List<DialogButton>();
    }

    // Windows PowerShell 5.1 の ConvertTo-Json は字下げが不揃いで ' & < > も \u00XX になるため整形し直す
    public static class Json
    {
        public static string Format(string json)
        {
            var sb = new StringBuilder(json.Length * 2);
            int indent = 0;
            bool inStr = false;
            for (int i = 0; i < json.Length; i++)
            {
                char ch = json[i];
                if (inStr)
                {
                    if (ch == '\\' && i + 5 < json.Length && json[i + 1] == 'u' && json[i + 2] == '0' && json[i + 3] == '0')
                    {
                        string hex = json.Substring(i + 4, 2).ToLowerInvariant();
                        if (hex == "26") { sb.Append('&'); i += 5; continue; }
                        if (hex == "27") { sb.Append('\''); i += 5; continue; }
                        if (hex == "3c") { sb.Append('<'); i += 5; continue; }
                        if (hex == "3e") { sb.Append('>'); i += 5; continue; }
                    }
                    sb.Append(ch);
                    if (ch == '\\' && i + 1 < json.Length) { sb.Append(json[++i]); }
                    else if (ch == '"') inStr = false;
                    continue;
                }
                switch (ch)
                {
                    case '"': inStr = true; sb.Append(ch); break;
                    case '{':
                    case '[':
                        {
                            int j = i + 1;
                            while (j < json.Length && char.IsWhiteSpace(json[j])) j++;
                            char close = ch == '{' ? '}' : ']';
                            if (j < json.Length && json[j] == close) { sb.Append(ch).Append(close); i = j; break; }
                            sb.Append(ch); indent++; NewLine(sb, indent); break;
                        }
                    case '}':
                    case ']': indent--; NewLine(sb, indent); sb.Append(ch); break;
                    case ',': sb.Append(ch); NewLine(sb, indent); break;
                    case ':': sb.Append(": "); break;
                    default:
                        if (!char.IsWhiteSpace(ch)) sb.Append(ch);
                        break;
                }
            }
            return sb.ToString();
        }

        static void NewLine(StringBuilder sb, int indent)
        {
            sb.Append("\r\n");
            sb.Append(' ', Math.Max(0, indent) * 2);
        }
    }

    public static class Win32
    {
        delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder sb, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr hwnd);
        [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

        const uint WM_COMMAND = 0x0111;
        const uint WM_CLOSE = 0x0010;

        public static uint ProcessIdOf(long hwnd)
        {
            uint pid;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return pid;
        }

        static string ClassOf(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, sb.Capacity); return sb.ToString(); }
        static string TextOf(IntPtr h) { var sb = new StringBuilder(8192); GetWindowText(h, sb, sb.Capacity); return sb.ToString(); }

        // Visible standard dialogs (#32770) that belong to the given process.
        public static List<DialogInfo> FindDialogs(uint pid)
        {
            var result = new List<DialogInfo>();
            EnumWindows((h, l) =>
            {
                uint p;
                GetWindowThreadProcessId(h, out p);
                if (p != pid || !IsWindowVisible(h) || ClassOf(h) != "#32770") return true;
                var d = new DialogInfo { Hwnd = h, Title = TextOf(h) };
                EnumChildWindows(h, (c, l2) =>
                {
                    if (!IsWindowVisible(c)) return true;
                    string cls = ClassOf(c);
                    string text = TextOf(c);
                    if (cls == "Button")
                    {
                        if (IsWindowEnabled(c)) d.Buttons.Add(new DialogButton { Hwnd = c, Id = GetDlgCtrlID(c), Text = text });
                    }
                    else if ((cls == "Static" || cls == "Edit" || cls == "RichEdit20W") && text.Length > 0)
                    {
                        d.Texts.Add(text);
                    }
                    return true;
                }, IntPtr.Zero);
                result.Add(d);
                return true;
            }, IntPtr.Zero);
            return result;
        }

        public static void Click(IntPtr dialog, DialogButton button)
        {
            PostMessage(dialog, WM_COMMAND, new IntPtr(button.Id & 0xFFFF), button.Hwnd);
        }

        public static void Close(IntPtr dialog)
        {
            PostMessage(dialog, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);

        // フォーカスを奪わずに最小化（.accde 作成時に Access を表示する必要があるため）
        public static void MinimizeNoActivate(long hwnd)
        {
            ShowWindow(new IntPtr(hwnd), 7); // SW_SHOWMINNOACTIVE
        }
    }

    // 帳票スナップショットの画像比較
    public class ImageDiffResult
    {
        public double Ratio;
        public long DiffPixels;
        public int Width;
        public int Height;
        public bool SizeChanged;
    }

    public static class ImageDiff
    {
        static int[] Load(string path, out int w, out int h)
        {
            using (var src = new System.Drawing.Bitmap(path))
            using (var bmp = new System.Drawing.Bitmap(src.Width, src.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                using (var g = System.Drawing.Graphics.FromImage(bmp)) { g.DrawImage(src, 0, 0, src.Width, src.Height); }
                w = bmp.Width; h = bmp.Height;
                var data = bmp.LockBits(new System.Drawing.Rectangle(0, 0, w, h), System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                var px = new int[w * h];
                for (int y = 0; y < h; y++) Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), px, y * w, w);
                bmp.UnlockBits(data);
                return px;
            }
        }

        // tolerance: 色の差（0-255）をどこまで同じとみなすか。diffPath に差分画像（違う所が赤）を書く
        public static ImageDiffResult Compare(string baseline, string current, string diffPath, int tolerance)
        {
            int wa, ha, wb, hb;
            var a = Load(baseline, out wa, out ha);
            var b = Load(current, out wb, out hb);
            int w = Math.Max(wa, wb), h = Math.Max(ha, hb);
            var outPx = new int[w * h];
            long diff = 0;
            for (int y = 0; y < h; y++)
            {
                for (int x = 0; x < w; x++)
                {
                    bool inA = x < wa && y < ha, inB = x < wb && y < hb;
                    int ca = inA ? a[y * wa + x] : -1, cb = inB ? b[y * wb + x] : -1;
                    bool same = inA && inB
                        && Math.Abs(((ca >> 16) & 255) - ((cb >> 16) & 255)) <= tolerance
                        && Math.Abs(((ca >> 8) & 255) - ((cb >> 8) & 255)) <= tolerance
                        && Math.Abs((ca & 255) - (cb & 255)) <= tolerance;
                    if (same)
                    {
                        int lum = (((cb >> 16) & 255) + ((cb >> 8) & 255) + (cb & 255)) / 3;
                        int g = 190 + lum * 65 / 255; // 薄いグレーで背景として表示
                        outPx[y * w + x] = unchecked((int)0xFF000000) | (g << 16) | (g << 8) | g;
                    }
                    else
                    {
                        diff++;
                        outPx[y * w + x] = unchecked((int)0xFFE00000); // 赤
                    }
                }
            }
            if (!string.IsNullOrEmpty(diffPath))
            {
                using (var bmp = new System.Drawing.Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    var data = bmp.LockBits(new System.Drawing.Rectangle(0, 0, w, h), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    for (int y = 0; y < h; y++) Marshal.Copy(outPx, y * w, IntPtr.Add(data.Scan0, y * data.Stride), w);
                    bmp.UnlockBits(data);
                    bmp.Save(diffPath, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            return new ImageDiffResult { Ratio = (double)diff / ((double)w * h), DiffPixels = diff, Width = w, Height = h, SizeChanged = (wa != wb || ha != hb) };
        }
    }
}
