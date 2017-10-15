using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ShellLib {
    [ComVisible(false)]
    public enum LVCFMT {
        LEFT = 0x0000,
        RIGHT = 0x0001,
        CENTER = 0x0002,
        JUSTIFYMASK = 0x0003,
        IMAGE = 0x0800,
        BITMAP_ON_RIGHT = 0x1000,
        COL_HAS_IMAGES = 0x8000
    }
    [Flags, ComVisible(false)]
    public enum SHCOLSTATE {
        TYPE_STR = 0x1,
        TYPE_INT = 0x2,
        TYPE_DATE = 0x3,
        TYPEMASK = 0xf,
        ONBYDEFAULT = 0x10,
        SLOW = 0x20,
        EXTENDED = 0x40,
        SECONDARYUI = 0x80,
        HIDDEN = 0x100,
        PREFER_VARCMP = 0x200
    }

    [ComVisible(false), StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public class LPCSHCOLUMNINIT {
        public uint dwFlags; //ulong
        public uint dwReserved; //ulong
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string wszFolder; //[MAX_PATH]; wchar
    }

    [ComVisible(false), StructLayout(LayoutKind.Sequential)]
    public struct SHCOLUMNID {
        public Guid fmtid; //GUID
        public uint pid; //DWORD
    }

    [ComVisible(false), StructLayout(LayoutKind.Sequential)]
    public class LPCSHCOLUMNID {
        public Guid fmtid; //GUID
        public uint pid; //DWORD
    }

    /// <summary>
    /// https://msdn.microsoft.com/ja-jp/library/windows/desktop/bb759751(v=vs.85).aspx
    /// </summary>
    [ComVisible(false), StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 1)]
    public struct SHCOLUMNINFO {
        public SHCOLUMNID scid; //SHCOLUMNID
        public ushort vt; //VARTYPE
        public LVCFMT fmt; //DWORD
        public uint cChars; //UINT
        public SHCOLSTATE csFlags;  //DWORD
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] //MAX_COLUMN_NAME_LEN
        public string wszTitle; //WCHAR
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] //MAX_COLUMN_DESC_LEN
        public string wszDescription; //WCHAR
    }

    [ComVisible(false), StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public class LPCSHCOLUMNDATA {
        public uint dwFlags; //ulong
        public uint dwFileAttributes; //dword
        public uint dwReserved; //ulong
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pwszExt; //wchar
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string wszFile; //[MAX_PATH]; wchar
    }

    /// E8025004Å`ÇÕå≈íË
    [ComVisible(false), ComImport, Guid("E8025004-1C42-11d2-BE2C-00A0C9A83DA1"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IColumnProvider {
        [PreserveSig()]
        int Initialize(LPCSHCOLUMNINIT psci);
        [PreserveSig()]
        int GetColumnInfo(int dwIndex, out SHCOLUMNINFO psci);

        /// <summary>
        /// Note: these objects must be threadsafe!  GetItemData _will_ be called
        /// simultaneously from multiple threads.
        /// </summary>
        [PreserveSig()]
        int GetItemData(LPCSHCOLUMNID pscid, LPCSHCOLUMNDATA pscd, out object /*VARIANT */ pvarData);
    }


    [ComVisible(false)]
    public abstract class ColumnProvider : IColumnProvider {
        [DllImport("Shell32.dll")]
        static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
        const int SHCNE_ASSOCCHANGED = 0x08000000;

        public abstract int Initialize(LPCSHCOLUMNINIT psci);
        public abstract int GetColumnInfo(int dwIndex, out SHCOLUMNINFO psci);
        public abstract int GetItemData(LPCSHCOLUMNID pscid, LPCSHCOLUMNDATA pscd, out object pvarData);
        static readonly String RegFolderName = "ColumnHandlers";


        [ComRegisterFunction]
        public static void Register(System.Type t) {
            {
                RegistryKey key = Registry.ClassesRoot.CreateSubKey($@"Folder\shellex\{RegFolderName}\" + t.GUID.ToString("B"));
                key.SetValue(string.Empty, "Ç»ÇÒÇ≈Ç‡Ç¢Ç¢íl");
                key.Close();
            }
            {
                RegistryKey key = Registry.ClassesRoot.CreateSubKey($@"CLSID\" + t.GUID.ToString("B"));
                key.SetValue(string.Empty, "Desktop.inièÓïÒ");
                key.Close();
            }

            // Tell Explorer to refresh
            SHChangeNotify(SHCNE_ASSOCCHANGED, 0, IntPtr.Zero, IntPtr.Zero);
        }

        [ComUnregisterFunction]
        public static void UnRegister(System.Type t) {
            try {
                Registry.ClassesRoot.DeleteSubKeyTree($@"Folder\shellex\{RegFolderName}\" + t.GUID.ToString("B"));
            } catch {
                // Ignore all
            }

            // Tell Explorer to refresh
            SHChangeNotify(SHCNE_ASSOCCHANGED, 0, IntPtr.Zero, IntPtr.Zero);
        }
    }
}

