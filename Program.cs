using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace tcache
{
    class Program
    {
        [DllImport("gdi32.dll")] public static extern int DeleteObject(IntPtr hObject);
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void SHCreateItemFromParsingName(
            [In][MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            [In] IntPtr pbc,
            [In][MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [Out][MarshalAs(UnmanagedType.Interface, IidParameterIndex = 2)] out IShellItem ppv);
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("f676c15d-596a-4ce2-8234-33996f445db1")]
        interface IThumbnailCache
        {
            uint GetThumbnail(
                [In] IShellItem pShellItem,
                [In] uint cxyRequestedThumbSize,
                [In] WTS_FLAGS flags /*default:  WTS_FLAGS.WTS_EXTRACT*/,
                [Out][MarshalAs(UnmanagedType.Interface)] out ISharedBitmap ppvThumb,
                [Out] out WTS_CACHEFLAGS pOutFlags,
                [Out] out WTS_THUMBNAILID pThumbnailID
            );
            void GetThumbnailByID(
                [In, MarshalAs(UnmanagedType.Struct)] WTS_THUMBNAILID thumbnailID,
                [In] uint cxyRequestedThumbSize,
                [Out][MarshalAs(UnmanagedType.Interface)] out ISharedBitmap ppvThumb,
                [Out] out WTS_CACHEFLAGS pOutFlags
            );
        }
        [Flags]
        enum WTS_FLAGS : uint
        {
            WTS_EXTRACT = 0x00000000,
            WTS_INCACHEONLY = 0x00000001,
            WTS_FASTEXTRACT = 0x00000002,
            WTS_SLOWRECLAIM = 0x00000004,
            WTS_FORCEEXTRACTION = 0x00000008,
            WTS_EXTRACTDONOTCACHE = 0x00000020,
            WTS_SCALETOREQUESTEDSIZE = 0x00000040,
            WTS_SKIPFASTEXTRACT = 0x00000080,
            WTS_EXTRACTINPROC = 0x00000100
        }
        [Flags]
        enum WTS_CACHEFLAGS : uint
        {
            WTS_DEFAULT = 0x00000000,
            WTS_LOWQUALITY = 0x00000001,
            WTS_CACHED = 0x00000002
        }

        [StructLayout(LayoutKind.Sequential, Size = 16), Serializable]
        struct WTS_THUMBNAILID
        {
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 16)]
            byte[] rgbKey;
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        public interface IShellItem
        {
            void BindToHandler(IntPtr pbc,
                [MarshalAs(UnmanagedType.LPStruct)] Guid bhid,
                [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
                out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        };
        public enum SIGDN : uint
        {
            NORMALDISPLAY = 0,
            PARENTRELATIVEPARSING = 0x80018001,
            PARENTRELATIVEFORADDRESSBAR = 0x8001c001,
            DESKTOPABSOLUTEPARSING = 0x80028000,
            PARENTRELATIVEEDITING = 0x80031001,
            DESKTOPABSOLUTEEDITING = 0x8004c000,
            FILESYSPATH = 0x80058000,
            URL = 0x80068000
        }

        [ComImport()]
        [Guid("091162a4-bc96-411f-aae8-c5122cd03363")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ISharedBitmap
        {
            uint Detach([Out] out IntPtr phbm);
            uint GetFormat([Out] out WTS_ALPHATYPE pat);
            uint GetSharedBitmap([Out] out IntPtr phbm);
            uint GetSize([Out, MarshalAs(UnmanagedType.Struct)] out SIZE pSize);
            uint InitializeBitmap([In] IntPtr hbm, [In] WTS_ALPHATYPE wtsAT);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE
        {
            public int cx;
            public int cy;
            public SIZE(int cx, int cy)
            {
                this.cx = cx;
                this.cy = cy;
            }
        }
        public enum WTS_ALPHATYPE : uint
        {
            WTSAT_UNKNOWN = 0,
            WTSAT_RGB = 1,
            WTSAT_ARGB = 2
        }
        static void Main(string[] args)
        {
            WTS_CACHEFLAGS cFlags;
            WTS_THUMBNAILID bmpId;
            Guid iIdIShellItem = new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe");
            Guid CLSIDLocalThumbnailCache = new Guid("50ef4544-ac9f-4a8e-b21b-8a26180db13f");
            var TBCacheType = Type.GetTypeFromCLSID(CLSIDLocalThumbnailCache);
            string path = args.Length > 0 ? Path.GetFullPath(args[0]) : "";
            if (args.Length > 0 && Directory.Exists(path))
            {
                uint[] dimensions = new uint[] { 2560, 1920, 1280, 1024, 768, 256, 96, 48, 32 };
                var toptions = new ParallelOptions()
                {
                    MaxDegreeOfParallelism = Convert.ToInt32(Math.Ceiling((Environment.ProcessorCount * 0.75) * 2.0))
                };
                Parallel.ForEach(Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories), toptions, file =>
                {
                    try
                    {
                        Console.WriteLine(file);
                        IThumbnailCache TBCache = (IThumbnailCache)Activator.CreateInstance(TBCacheType);
                        IShellItem shellItem = null;
                        ISharedBitmap bmp = null;
                        IntPtr phbmp;
                        SHCreateItemFromParsingName(file, IntPtr.Zero, iIdIShellItem, out shellItem);
                        foreach (uint d in dimensions)
                        {
                            TBCache.GetThumbnail(shellItem, d, WTS_FLAGS.WTS_EXTRACTINPROC, out bmp, out cFlags, out bmpId);
                            if (bmp != null)
                            {
                                bmp.GetSharedBitmap(out phbmp);
                                DeleteObject(phbmp);
                                Marshal.ReleaseComObject(bmp);
                            }
                            bmp = null;
                        }
                        if (shellItem != null) Marshal.ReleaseComObject(shellItem);
                        shellItem = null;
                    }
                    catch { }
                });
            }
            else
            {
                Console.WriteLine("Thumbnail cache generator.\nUsage: {0} <path>", System.Diagnostics.Process.GetCurrentProcess().ProcessName);
            }
        }
    }
}
