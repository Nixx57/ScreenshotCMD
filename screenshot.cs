using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.VisualBasic;

/// Provides functions to capture the entire screen, or a particular window, and save it to a file. 

namespace Screenshot
{
    public partial class ScreenCapture
    {
        static bool fullscreen = true;
        static string file = "screenshot.bmp";
        static ImageFormat format = ImageFormat.Bmp;
        static string windowTitle = "";
        static int? originX = null;
        static int? originY = null;
        static int? endX = null;
        static int? endY = null;

        public static Image CaptureActiveWindow()
        {
            return CaptureWindow(User32.GetForegroundWindow());
        }

        public static Image CaptureScreen()
        {
            return CaptureWindow(User32.GetDesktopWindow());
        }

        private static Image CaptureWindow(IntPtr handle)
        {
            // get te hDC of the target window 
            IntPtr hdcSrc = User32.GetWindowDC(handle);
            // get the size 
            User32.RECT windowRect = new();
            if (originX != null && endX != null && originY != null && endY != null)
            {
                windowRect.left = (Int32)originX;
                windowRect.right = (Int32)endX;
                windowRect.top = (Int32)originY;
                windowRect.bottom = (Int32)endY;
            }
            else
            {
                User32.GetWindowRect(handle, ref windowRect);
            }
            int width = windowRect.right - windowRect.left;
            int height = windowRect.bottom - windowRect.top;
            // create a device context we can copy to 
            IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc);
            // create a bitmap we can copy it to, 
            // using GetDeviceCaps to get the width/height 
            IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc, width, height);
            // select the bitmap object 
            IntPtr hOld = GDI32.SelectObject(hdcDest, hBitmap);
            // bitblt over 
            GDI32.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, windowRect.left, windowRect.top, GDI32.SRCCOPY);
            // restore selection 
            GDI32.SelectObject(hdcDest, hOld);
            // clean up 
            GDI32.DeleteDC(hdcDest);
            User32.ReleaseDC(handle, hdcSrc);
            // get a .NET image object for it 
            Image img = Image.FromHbitmap(hBitmap);
            // free up the Bitmap object 
            GDI32.DeleteObject(hBitmap);
            return img;
        }

        public static void CaptureActiveWindowToFile(string filename, ImageFormat format)
        {
            Image img = CaptureActiveWindow();
            img.Save(filename, format);
        }

        public static void CaptureScreenToFile(string filename, ImageFormat format)
        {
            Image img = CaptureScreen();
            img.Save(filename, format);
        }

        static void ParseArguments()
        {
            String[] arguments = Environment.GetCommandLineArgs();

            switch (arguments.Length)
            {
                case 1:
                    PrintHelp();
                    Environment.Exit(0);
                    break;
                case 2:
                    break;
                case 3:
                    windowTitle = arguments[2];
                    fullscreen = false;
                    break;
                case 6:
                    originX = Int32.Parse(arguments[2]);
                    originY = Int32.Parse(arguments[3]);
                    endX = Int32.Parse(arguments[4]);
                    endY = Int32.Parse(arguments[5]);
                    break;
                case 7:
                    windowTitle = arguments[2]; fullscreen = false;
                    originX = Int32.Parse(arguments[3]);
                    originY = Int32.Parse(arguments[4]);
                    endX = Int32.Parse(arguments[5]);
                    endY = Int32.Parse(arguments[6]);
                    break;
                default:
                    PrintHelp();
                    Environment.Exit(0);
                    break;
            }

            if (arguments[1].ToLower().Equals("/h") || arguments[1].ToLower().Equals("/help"))
            {
                PrintHelp();
                Environment.Exit(0);
            }
            file = arguments[1];

            Dictionary<String, ImageFormat> formats =
            new()
            {
            { "bmp", ImageFormat.Bmp },
            { "emf", ImageFormat.Emf },
            { "exif", ImageFormat.Exif },
            { "jpg", ImageFormat.Jpeg },
            { "jpeg", ImageFormat.Jpeg },
            { "gif", ImageFormat.Gif },
            { "png", ImageFormat.Png },
            { "tiff", ImageFormat.Tiff },
            { "wmf", ImageFormat.Wmf }
            };

            String ext = "";
            if (file.LastIndexOf('.') > -1)
            {
                ext = file.ToLower().Substring(file.LastIndexOf('.') + 1, file.Length - file.LastIndexOf('.') - 1);
            }
            else
            {
                Console.WriteLine("Invalid file name - no extension");
                Environment.Exit(7);
            }

            try
            {
                format = formats[ext];
            }
            catch (Exception e)
            {
                Console.WriteLine("Probably wrong file format:" + ext);
                Console.WriteLine(e.ToString());
                Environment.Exit(8);
            }
        }

        static void PrintHelp()
        {
            //clears the extension from the script name
            String scriptName = Environment.GetCommandLineArgs()[0];
            scriptName = scriptName[..];
            Console.WriteLine(scriptName + " captures the screen or the active window and saves it to a file.");
            Console.WriteLine("");
            Console.WriteLine("Usage:");
            Console.WriteLine(" " + scriptName + " filename  [WindowTitle] [origin X] [origin Y] [end X] [end Y]");
            Console.WriteLine("");
            Console.WriteLine("filename - the file where the screen capture will be saved");
            Console.WriteLine("     allowed file extensions are - Bmp,Emf,Exif,Gif,Icon,Jpeg,Png,Tiff,Wmf.");
            Console.WriteLine("WindowTitle - instead of capture whole screen you can point to a window ");
            Console.WriteLine("     with a title which will put on focus and captuted.");
            Console.WriteLine("     For WindowTitle you can pass only the first few characters.");
            Console.WriteLine("     If don't want to change the current active window pass only \"\"");
            Console.WriteLine("Coordinates - capture a rectangle of screen");
        }

        public static void Main()
        {
            AppDomain.CurrentDomain.FirstChanceException += (sender, eventArgs) => Console.WriteLine(eventArgs.Exception.ToString());
            var hresult = User32.SetProcessDPIAware();

            ParseArguments();
            ScreenCapture sc = new();

            if (!fullscreen && !windowTitle.Equals(""))
            {
                try
                {
                    Interaction.AppActivate(windowTitle);
                    Console.WriteLine("setting " + windowTitle + " on focus");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Probably there's no window like " + windowTitle);
                    Console.WriteLine(e.ToString());
                    Environment.Exit(9);
                }
            }
            try
            {
                if (fullscreen)
                {
                    Console.WriteLine("Taking a capture of the whole screen to " + file);
                    CaptureScreenToFile(file, format);
                }
                else
                {
                    Console.WriteLine("Taking a capture of the active window to " + file);
                    CaptureActiveWindowToFile(file, format);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Check if file path is valid " + file);
                Console.WriteLine(e.ToString());
            }
        }

        /// Helper class containing Gdi32 API functions 

        private partial class GDI32
        {

            public const int SRCCOPY = 0x00CC0020; // BitBlt dwRop parameter 
            [LibraryImport("gdi32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static partial bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hObjectSource, int nXSrc, int nYSrc, int dwRop);
            [LibraryImport("gdi32.dll")]
            public static partial IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth,
                int nHeight);
            [LibraryImport("gdi32.dll")]
            public static partial IntPtr CreateCompatibleDC(IntPtr hDC);
            [LibraryImport("gdi32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static partial bool DeleteDC(IntPtr hDC);
            [LibraryImport("gdi32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static partial bool DeleteObject(IntPtr hObject);
            [LibraryImport("gdi32.dll")]
            public static partial IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
        }


        /// Helper class containing User32 API functions 

        private static partial class User32
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct RECT
            {
                public int left;
                public int top;
                public int right;
                public int bottom;
            }
            [LibraryImport("user32.dll")]
            public static partial IntPtr GetDesktopWindow();
            [LibraryImport("user32.dll")]
            public static partial IntPtr GetWindowDC(IntPtr hWnd);
            [LibraryImport("user32.dll")]
            public static partial IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);
            [LibraryImport("user32.dll")]
            public static partial IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);
            [LibraryImport("user32.dll")]
            public static partial IntPtr GetForegroundWindow();
            [LibraryImport("user32.dll")]
            public static partial int SetProcessDPIAware();
        }
    }
}