using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ShellContextMenu
{
    [ComImport,Guid("000214e4-0000-0000-c000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IContextMenu
    {
        [PreserveSig]
        int QueryContextMenu(IntPtr hMenu, uint indexMenu, int idCmdFirst, int idCmdLast, CMF uFlags);

        [PreserveSig]
        int InvokeCommand(IntPtr pici);

        [PreserveSig]
        int GetCommandString(
            int idcmd, 
            GCS uflags, 
            int reserved, 
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder commandString, 
            int cch );
    }

    [ComImport, Guid("000214f4-0000-0000-c000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IContextMenu2
    {
        [PreserveSig]
        int QueryContextMenu(IntPtr hMenu, uint indexMenu, int idCmdFirst, int idCmdLast, CMF uFlags);

        [PreserveSig]
        int InvokeCommand(IntPtr pici);

        [PreserveSig]
        int GetCommandString(
            int idcmd, 
            GCS uflags,
            int reserved, 
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder commandString, 
            int cch);

        [PreserveSig]
        int HandleMenuMsg(uint uMsg, IntPtr wParam, IntPtr lParam);
    }


    [Flags]
    public enum GCS : uint
    {
        GCS_VERBA = 0x00000000,
        GCS_HELPTEXTA = 0x00000001,
        GCS_VALIDATEA = 0x00000002,
        GCS_VERBW = 0x00000004,
        GCS_HELPTEXTW = 0x00000005,
        GCS_VALIDATEW = 0x00000006,
        GCS_VERBICONW = 0x00000014,
        GCS_UNICODE = 0x00000004,
    }

    public enum CMF : uint
    {
        CMF_NORMAL = 0x00000000,
        CMF_DEFAULTONLY = 0x00000001,
        CMF_VERBSONLY = 0x00000002,
        CMF_EXPLORE = 0x00000004,
        CMF_NOVERBS = 0x00000008,
        CMF_CANRENAME = 0x00000010,
        CMF_NODEFAULT = 0x00000020,
        CMF_INCLUDESTATIC = 0x00000040,
        CMF_ITEMMENU = 0x00000080,
        CMF_EXTENDEDVERBS = 0x00000100,
        CMF_DISABLEDVERBS = 0x00000200,
        CMF_ASYNCVERBSTATE = 0x00000400,
        CMF_OPTIMIZEFORINVOKE = 0x00000800,
        CMF_SYNCCASCADEMENU = 0x00001000,
        CMF_DONOTPICKDEFAULT = 0x00002000,
        CMF_RESERVED = 0xFFFF0000
    }

    [ComImport,Guid("000214E6-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellFolder
    {
        [PreserveSig]
        uint ParseDisplayName(
            IntPtr hwnd,
            IntPtr pbc,
            string displayName,
            out int pchEaten,
            out IntPtr pidl,
            ref SFGAOF pdwAttributes
            );

        [PreserveSig]
        uint EnumObjects(
            IntPtr hwnd,
            SHCONTF flags,
            out IEnumIDList ppenumIDList);

        [PreserveSig]
        uint BindToObject(
            IntPtr pidl,
            IntPtr pbc,
            ref Guid iid,
            out IShellFolder ppv);

        [PreserveSig]
        uint BindToStorage(
            IntPtr pidl,
            IntPtr pbc,
            ref Guid iid,
            out IntPtr ppv);

        [PreserveSig]
        uint CompareIDs(int lParam, IntPtr pidl1, IntPtr pidl2);

        [PreserveSig]
        uint CreateViewObject(
            IntPtr hwndOwner,
            ref Guid iid,
            out IntPtr ppv);

        [PreserveSig]
        uint GetAttributesOf(
            uint cidl,
            [MarshalAs(UnmanagedType.LPArray,SizeParamIndex=0)] IntPtr[] apidl,
            ref SFGAOF rggInOut);

        [PreserveSig]
        uint GetUIObjectOf(
            IntPtr hwnd,
            uint cidl,
            [MarshalAs(UnmanagedType.LPArray,SizeParamIndex =1)] IntPtr[] apidl,
            ref Guid iid,
            uint rgfReserved,
            out IntPtr ppv);

        [PreserveSig]
        uint GetDisplayNameOf(IntPtr pidl, SHGDN uflags, out STRRET pName);

        [PreserveSig]
        uint SetNameOf(IntPtr hwnd, IntPtr pidl, string pszName, SHGDN uFlags, out IntPtr ppidl);
    }

    [ComImport,Guid("000214F2-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumIDList
    {
        [PreserveSig]
        uint Next(
            uint celt,
            [In, Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] rgelt,
            out int pceltFetched
            );

        [PreserveSig]
        uint Skip(uint celt);

        [PreserveSig]
        uint Reset();

        [PreserveSig]
        uint Clone(
            out IEnumIDList ppenum);
    }

    [Flags]
    public enum CSIDL : uint
    {
        CSIDL_DESKTOP = 0x0000,
    }

    [Flags]
    public enum SHGFI
    {
        SHGFI_ICON = 0x00000100,
        SHGFI_DISPLAYNAME = 0x000000200,
        SHGFI_TYPENAME = 0x000000400,
        SHGFI_ATTRIBUTES = 0x000000800,
        SHGFI_ICONLOCATION = 0x000001000,
        SHGFI_EXETYPE = 0x000002000,
        SHGFI_SYSICONINDEX = 0x000004000,
        SHGFI_LINKOVERLAY = 0x000008000,
        SHGFI_SELECTED = 0x000010000,
        SHGFI_ATTR_SPECIFIED = 0x000020000,
        SHGFI_LARGEICON = 0x000000000,
        SHGFI_SMALLICON = 0x000000001,
        SHGFI_OPENICON = 0x000000002,
        SHGFI_SHELLICONSIZE = 0x000000004,
        SHGFI_PIDL = 0x000000008,
        SHGFI_USEFILEATTRIBUTES = 0x000000010,
        SHGFI_ADDOVERLAYS = 0x000000020,
        SHGFI_OVERLAYINDEX = 0x000000040
    }

    [Flags]
    public enum STRRET : uint
    {
        STRRET_WSTR = 0,
        STRRET_OFFSET = 1,
        STRRET_CSTR = 2,
    }

    [Flags]
    public enum SFGAOF : uint
    {
        SFGAO_CANCOPY = 0x1,                   // Objects can be copied  (DROPEFFECT_COPY)
        SFGAO_CANMOVE = 0x2,                   // Objects can be moved   (DROPEFFECT_MOVE)
        SFGAO_CANLINK = 0x4,                   // Objects can be linked  (DROPEFFECT_LINK)
        SFGAO_STORAGE = 0x00000008,            // Supports BindToObject(IID_IStorage)
        SFGAO_CANRENAME = 0x00000010,          // Objects can be renamed
        SFGAO_CANDELETE = 0x00000020,          // Objects can be deleted
        SFGAO_HASPROPSHEET = 0x00000040,       // Objects have property sheets
        SFGAO_DROPTARGET = 0x00000100,         // Objects are drop target
        SFGAO_CAPABILITYMASK = 0x00000177,
        SFGAO_SYSTEM = 0x00001000,
        SFGAO_ENCRYPTED = 0x00002000,          // Object is encrypted (use alt color)
        SFGAO_ISSLOW = 0x00004000,             // 'Slow' object
        SFGAO_GHOSTED = 0x00008000,            // Ghosted icon
        SFGAO_LINK = 0x00010000,               // Shortcut (link)
        SFGAO_SHARE = 0x00020000,              // Shared
        SFGAO_READONLY = 0x00040000,           // Read-only
        SFGAO_HIDDEN = 0x00080000,             // Hidden object
        SFGAO_DISPLAYATTRMASK = 0x000FC000,
        SFGAO_FILESYSANCESTOR = 0x10000000,    // May contain children with SFGAO_FILESYSTEM
        SFGAO_FOLDER = 0x20000000,             // Support BindToObject(IID_IShellFolder)
        SFGAO_FILESYSTEM = 0x40000000,         // Is a win32 file system object (file/folder/root)
        SFGAO_HASSUBFOLDER = 0x80000000,       // May contain children with SFGAO_FOLDER
        SFGAO_CONTENTSMASK = 0x80000000,
        SFGAO_VALIDATE = 0x01000000,           // Invalidate cached information
        SFGAO_REMOVABLE = 0x02000000,          // Is this removeable media?
        SFGAO_COMPRESSED = 0x04000000,         // Object is compressed (use alt color)
        SFGAO_BROWSABLE = 0x08000000,          // Supports IShellFolder, but only implements CreateViewObject() (non-folder view)
        SFGAO_NONENUMERATED = 0x00100000,      // Is a non-enumerated object
        SFGAO_NEWCONTENT = 0x00200000,         // Should show bold in explorer tree
        SFGAO_CANMONIKER = 0x00400000,         // Defunct
        SFGAO_HASSTORAGE = 0x00400000,         // Defunct
        SFGAO_STREAM = 0x00400000,             // Supports BindToObject(IID_IStream)
        SFGAO_STORAGEANCESTOR = 0x00800000,    // May contain children with SFGAO_STORAGE or SFGAO_STREAM
        SFGAO_STORAGECAPMASK = 0x70C50008,     // For determining storage capabilities, ie for open/save semantics
    }
    [Flags]
    public enum SHGDN : uint
    {
        SHGDN_NORMAL = 0x0000,                 // Default (display purpose)
        SHGDN_INFOLDER = 0x0001,               // Displayed under a folder (relative)
        SHGDN_FOREDITING = 0x1000,             // For in-place editing
        SHGDN_FORADDRESSBAR = 0x4000,          // UI friendly parsing name (remove ugly stuff)
        SHGDN_FORPARSING = 0x8000,             // Parsing name for ParseDisplayName()
    }

    [Flags]
    public enum SHCONTF : uint
    {
        SHCONTF_CHECKING_FOR_CHILDREN = 0x0010,
        SHCONTF_FOLDERS = 0x0020,              // Only want folders enumerated (SFGAO_FOLDER)
        SHCONTF_NONFOLDERS = 0x0040,           // Include non folders
        SHCONTF_INCLUDEHIDDEN = 0x0080,        // Show items normally hidden
        SHCONTF_INIT_ON_FIRST_NEXT = 0x0100,   // Allow EnumObject() to return before validating enum
        SHCONTF_NETPRINTERSRCH = 0x0200,       // Hint that client is looking for printers
        SHCONTF_SHAREABLE = 0x0400,            // Hint that client is looking sharable resources (remote shares)
        SHCONTF_STORAGE = 0x0800,              // Include all items with accessible storage and their ancestors
        SHCONTF_NAVIGATION_ENUM = 0x01000,
        SHCONTF_FASTITEMS = 0x02000,
        SHCONTF_FLATLIST = 0x04000,
        SHCONTF_ENABLE_ASYNC = 0x08000,
        SHCONTF_INCLUDESUPERHIDDEN = 0x10000,
    }

    public static class ShellAPI
    {
        [DllImport("shell32.dll")]
        public extern static int SHGetDesktopFolder(out IntPtr ppshf );

        [DllImport("shell32.dll")]
        public extern static int SHBindToParent(
            IntPtr pidl,
            [MarshalAs(UnmanagedType.LPStruct)] Guid iid,
            out IntPtr ppv,
            ref IntPtr ppidlLast);

        [DllImport("user32.dll")]
        public extern static IntPtr CreatePopupMenu();

        [DllImport("user32.dll")]
        public extern static bool DestroyMenu(IntPtr hMenu);

        [DllImport("user32.dll")]
        public extern static int GetMenuItemCount(IntPtr hMenu);

        [DllImport("user32.dll")]
        public extern static int GetMenuString(
            IntPtr hMenu,
            uint uIDItem,
            StringBuilder str,
            int nMaxCount,
            uint uFlags);

    }

    enum MF
    {
        MF_BYCOMMAND = 0x00000000,
        MF_BYPOSITION = 0x00000400,
    }

}
