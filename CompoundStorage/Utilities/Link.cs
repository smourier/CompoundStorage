using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace CompoundStorage.Utilities
{
    public class Link
    {
        private readonly IShellLinkW _link;
        private readonly IPersistFile _file;

        public Link(object nativeObject = null)
        {
            if (nativeObject == null)
            {
                nativeObject = new ShellLink();
            }

            if (!(nativeObject is IShellLinkW link))
                throw new ArgumentException(null, nameof(nativeObject));

            if (!(nativeObject is IPersistFile file))
                throw new ArgumentException(null, nameof(nativeObject));

            _link = link;
            _file = file;
        }

        public object NativeObject => _link;

        public string CurrentFileName
        {
            get
            {
                _file.GetCurFile(out var file);
                return file;
            }
        }

        public virtual IntPtr TargetIdList
        {
            get
            {
                _link.GetIDList(out var pidl);
                return pidl;
            }
            set => Storage.Throw(_link.SetIDList(value), true);
        }

        public virtual string TargetPath
        {
            get
            {
                var data = new WIN32_FIND_DATA_W();
                var linkPath = new string('\0', 1024);
                _link.GetPath(linkPath, 1024, ref data, 0);

                var zero = linkPath.IndexOf('\0');
                if (zero < 0)
                    return linkPath;

                return linkPath.Substring(0, zero);
            }
            set => Storage.Throw(_link.SetPath(value), true);
        }

        public virtual string Description
        {
            get
            {
                var sb = new StringBuilder(1024);
                _link.GetDescription(sb, sb.Capacity);
                return sb.ToString();
            }
            set => Storage.Throw(_link.SetDescription(value), true);
        }

        public virtual string WorkingDirectory
        {
            get
            {
                var sb = new StringBuilder(1024);
                _link.GetWorkingDirectory(sb, sb.Capacity);
                return sb.ToString();
            }
            set => Storage.Throw(_link.SetWorkingDirectory(value), true);
        }

        public virtual string Arguments
        {
            get
            {
                var sb = new StringBuilder(1024);
                _link.GetArguments(sb, sb.Capacity);
                return sb.ToString();
            }
            set => Storage.Throw(_link.SetArguments(value), true);
        }

        public virtual ushort Hotkey
        {
            get
            {
                _link.GetHotkey(out var value);
                return value;
            }
            set => Storage.Throw(_link.SetHotkey(value), true);
        }

        public virtual SW ShowCommand
        {
            get
            {
                _link.GetShowCmd(out var value);
                return value;
            }
            set => Storage.Throw(_link.SetShowCmd(value), true);
        }

        public Guid Clsid
        {
            get
            {
                _file.GetClassID(out var clsid);
                return clsid;
            }
        }

        public bool IsDirty => _file.IsDirty() != 0;

        public override string ToString() => TargetPath ?? TargetIdList.ToString();
        public virtual int SetIconLocation(string path, int index, bool throwOnError = true) => Storage.Throw(_link.SetIconLocation(path, index), throwOnError);
        public virtual int GetIconLocation(out string path, out int index, bool throwOnError = false)
        {
            var sb = new StringBuilder(1024);
            var hr = Storage.Throw(_link.GetIconLocation(sb, sb.Capacity, out index), throwOnError);
            if (hr != 0)
            {
                path = null;
                return hr;
            }

            path = sb.ToString();
            return 0;
        }

        public virtual int SetRelativePath(string path, bool throwOnError = true) => Storage.Throw(_link.SetRelativePath(path, 0), throwOnError);
        public virtual int Save(string path, bool throwOnError = true)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));

            return Storage.Throw(_file.Save(path, false), throwOnError);
        }

        public static Link FromNativeObject(object obj)
        {
            if (obj is Link l)
                return l;

            if (obj is IShellLinkW folder)
                return new Link(folder);

            return new Link(null);
        }

        public static Link Load(string path, bool throwOnError = true)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));

            var fs = (IPersistFile)new ShellLink();
            var hr = Storage.Throw(fs.Load(path, STGM.STGM_READ), throwOnError);
            if (hr != 0)
                return null;

            var link = (IShellLinkW)fs;
            return new Link(link);
        }

        public static Link Load(Stream stream, bool throwOnError = true)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            return Load(new ManagedIStream(stream), throwOnError);
        }

        public static Link Load(IStream stream, bool throwOnError = true)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            var fs = (IPersistStream)new ShellLink();
            var hr = Storage.Throw(fs.Load(stream), throwOnError);
            if (hr != 0)
                return null;

            var link = (IShellLinkW)fs;
            return new Link(link);
        }

        [ComImport, Guid("00021401-0000-0000-c000-000000000046")]
        private class ShellLink { }

        [ComImport, Guid("000214f9-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private partial interface IShellLinkW
        {
            [PreserveSig]
            int GetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile, int cch, ref WIN32_FIND_DATA_W pfd, uint fFlags);

            [PreserveSig]
            int GetIDList(out IntPtr ppidl);

            [PreserveSig]
            int SetIDList(IntPtr pidl);

            [PreserveSig]
            int GetDescription([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);

            [PreserveSig]
            int SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);

            [PreserveSig]
            int GetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);

            [PreserveSig]
            int SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);

            [PreserveSig]
            int GetArguments([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);

            [PreserveSig]
            int SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);

            [PreserveSig]
            int GetHotkey(out ushort pwHotkey);

            [PreserveSig]
            int SetHotkey(ushort wHotkey);

            [PreserveSig]
            int GetShowCmd(out SW piShowCmd);

            [PreserveSig]
            int SetShowCmd(SW iShowCmd);

            [PreserveSig]
            int GetIconLocation([MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);

            [PreserveSig]
            int SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);

            [PreserveSig]
            int SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);

            [PreserveSig]
            int Resolve(IntPtr hwnd, SLR fFlags);

            [PreserveSig]
            int SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [ComImport, Guid("0000010c-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal partial interface IPersist
        {
            [PreserveSig]
            int GetClassID(out Guid pClassID);
        }

        [ComImport, Guid("0000010b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistFile : IPersist
        {
            [PreserveSig]
            new int GetClassID(out Guid pClassID);

            [PreserveSig]
            int IsDirty();

            [PreserveSig]
            int Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, STGM dwMode);

            [PreserveSig]
            int Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);

            [PreserveSig]
            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);

            [PreserveSig]
            int GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
        }

        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IPersistStream : IPersist
        {
            [PreserveSig]
            new int GetClassID(out Guid pClassID);

            [PreserveSig]
            int IsDirty();

            [PreserveSig]
            int Load(IStream pstm);

            [PreserveSig]
            int Save(IStream pstm, bool fClearDirty);

            [PreserveSig]
            int GetSizeMax(out long pcbSize);
        }
    }
}
