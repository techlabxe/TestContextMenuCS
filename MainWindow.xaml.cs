using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ShellContextMenu
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            var hwndSource = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSource != null)
            {
                hwndSource.AddHook(WndProc);
            }
        }


        private void PopupContextMenu( string fromPath, int x, int y )
        {
            IntPtr ppv;
            ShellAPI.SHGetDesktopFolder(out ppv);
            IShellFolder isf = (IShellFolder)Marshal.GetObjectForIUnknown(ppv);
            if (isf == null)
                return;

            IntPtr pidl = IntPtr.Zero;
            int eaten;
            SFGAOF attribs = default(SFGAOF);
            isf.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, fromPath, out eaten, out pidl, ref attribs);

            IntPtr pidlChild = IntPtr.Zero;

            ShellAPI.SHBindToParent(pidl, typeof(IShellFolder).GUID, out ppv, ref pidlChild);

            IShellFolder parentFolder = (IShellFolder)Marshal.GetObjectForIUnknown(ppv);
            IntPtr[] apidl = new IntPtr[1] { pidlChild };
            parentFolder.GetUIObjectOf(IntPtr.Zero, 1, apidl, typeof(IContextMenu).GUID, 0, out ppv);

            IContextMenu ctxMenuFirst = (IContextMenu)Marshal.GetObjectForIUnknown(ppv);
            var ctx2Guid = typeof(IContextMenu2).GUID;
            Marshal.QueryInterface(ppv, ref ctx2Guid, out ppv);
            IContextMenu2 ctxMenu = (IContextMenu2)Marshal.GetObjectForIUnknown(ppv);

            _ctxMenu = ctxMenu;

            IntPtr hMenu = ShellAPI.CreatePopupMenu();
            ctxMenu.QueryContextMenu(hMenu, 0, 1, 0x7fff, CMF.CMF_NORMAL);
            int menuCount = ShellAPI.GetMenuItemCount(hMenu);
            for (int i = 0; i < menuCount; ++i)
            {
                StringBuilder strbuf = new StringBuilder(256);
                int res = ShellAPI.GetMenuString(hMenu, (uint)i, strbuf, strbuf.Capacity, (int)MF.MF_BYPOSITION);
                if (res > 0)
                {
                    Debug.WriteLine(string.Format("{0:d2} {1}", i, strbuf.ToString()));
                }
            }

            const uint TPM_LEFTALIGN = 0;
            const uint TPM_RIGHTBUTTON = 0x0002;
            const uint TPM_RETURNCMD = 0x0100;

            uint flags = TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD;

            var hwndSource = PresentationSource.FromVisual(this) as HwndSource;
            ShellAPI.TrackPopupMenu(hMenu, flags, x, y, 0, hwndSource.Handle, IntPtr.Zero);

            ShellAPI.DestroyMenu(hMenu);
            _ctxMenu = null;
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if( _ctxMenu != null)
            {
                const int WM_INITMENUPOPUP = 0x117;

                if ( msg == WM_INITMENUPOPUP )
                {
                    Debug.WriteLine("WM_INITMENUPOPUP");
                    _ctxMenu.HandleMenuMsg((uint)msg, wParam, lParam);
                    handled = true;
                    return IntPtr.Zero;
                }
            }
            return IntPtr.Zero;
        }

        private IContextMenu2 _ctxMenu;

        private void Canvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var path = Environment.CurrentDirectory;
            var pt = this.PointToScreen( e.GetPosition(this) );

            int x = (int)pt.X;
            int y = (int)pt.Y;
            PopupContextMenu(path, x, y);
        }
    }
}
