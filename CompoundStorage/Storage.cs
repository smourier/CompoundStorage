using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using CompoundStorage.Utilities;

namespace CompoundStorage
{
    public sealed class Storage : IDisposable
    {
        public static Guid FMTID_SummaryInformation = new Guid("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}");
        public static Guid FMTID_DocSummaryInformation = new Guid("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}");
        public static Guid FMTID_UserDefinedProperties = new Guid("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");

        private IPropertySetStorage _psetStorage;
        private IStorage _storage;

        public Storage(string filePath, STGM mode = STGM.STGM_SHARE_DENY_NONE | STGM.STGM_DIRECT_SWMR, STGFMT format = STGFMT.STGFMT_ANY)
        {
            if (filePath == null)
                throw new ArgumentNullException(nameof(filePath));

            FilePath = filePath;

            Throw(StgOpenStorageEx(filePath, mode, format, 0, IntPtr.Zero, IntPtr.Zero, typeof(IStorage).GUID, out var storage), true, filePath);

            _storage = (IStorage)storage;
            _psetStorage = (IPropertySetStorage)storage;
        }

        internal Storage(IStorage storage)
        {
            _storage = storage;
            _psetStorage = (IPropertySetStorage)storage;
        }

        public string FilePath { get; }
        public object NativeObject => _storage;

        public IEnumerable<PropertyStorage> PropertyStorages
        {
            get
            {
                var pss = PSetStorage;
                var hr = pss.Enum(out var enumerator);
                if (hr < 0)
                    yield break;

                do
                {
                    var stat = new STATPROPSETSTG();
                    enumerator.Next(1, ref stat, out var fetched);
                    if (fetched != 1)
                        break;

                    pss.Open(stat.fmtid, STGM.STGM_SHARE_EXCLUSIVE, out var ps);
                    if (ps != null)
                        yield return new PropertyStorage(ps);
                }
                while (true);
            }
        }

        public IEnumerable<Element> Elements
        {
            get
            {
                var ps = PStorage;
                var hr = ps.EnumElements(0, IntPtr.Zero, 0, out var enumerator);
                if (hr < 0)
                    yield break;

                do
                {
                    var stat = new STATSTG();
                    enumerator.Next(1, ref stat, out var fetched);
                    if (fetched != 1)
                        break;

                    var type = (STGTY)stat.type;
                    switch (type)
                    {
                        case STGTY.STGTY_STREAM:
                            yield return new StreamElement(ps, stat);
                            break;

                        case STGTY.STGTY_STORAGE:
                            yield return new StorageElement(ps, stat);
                            break;

                        default:
                            yield return new Element(ps, stat);
                            break;
                    }
                }
                while (true);
            }
        }

        public int Commit(STGC flags, bool throwOnError = true) => Throw(PStorage.Commit(flags), throwOnError);
        public int Revert(bool throwOnError = true) => Throw(PStorage.Revert(), throwOnError);
        public int DestroyElement(string name, bool throwOnError = true) => Throw(PStorage.DestroyElement(name), throwOnError);
        public int RenameElement(string oldName, string newName, bool throwOnError = true) => Throw(PStorage.RenameElement(oldName, newName), throwOnError);
        public int MoveElementTo(string name, Storage storage, string newName, STGMOVE flags, bool throwOnError = true) => Throw(PStorage.MoveElementTo(name, storage.PStorage, newName, flags), throwOnError);
        public int SetClass(Guid clsid, bool throwOnError = true) => Throw(PStorage.SetClass(clsid), throwOnError);

        public int SetElementTimes(string name, DateTime? creationTime, DateTime? lastAccessTime, DateTime? lasWriteTime, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            long? ct;
            if (creationTime == null)
            {
                ct = null;
            }
            else
            {
                ct = creationTime.Value.ToFileTime();
            }

            long? lat;
            if (lastAccessTime == null)
            {
                lat = null;
            }
            else
            {
                lat = lastAccessTime.Value.ToFileTime();
            }

            long? lwt;
            if (lasWriteTime == null)
            {
                lwt = null;
            }
            else
            {
                lwt = lasWriteTime.Value.ToFileTime();
            }

            return SetElementTimes(name, ct, lat, lwt, throwOnError);
        }

        public int SetElementTimes(string name, long? creationTime, long? lastAccessTime, long? lastWriteTime, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            var creationTimePtr = IntPtr.Zero;
            if (creationTime.HasValue)
            {
                creationTimePtr = Marshal.AllocCoTaskMem(Marshal.SizeOf<long>());
                Marshal.StructureToPtr(creationTime.Value, creationTimePtr, false);
            }

            var lastAccessTimePtr = IntPtr.Zero;
            if (lastAccessTime.HasValue)
            {
                lastAccessTimePtr = Marshal.AllocCoTaskMem(Marshal.SizeOf<long>());
                Marshal.StructureToPtr(lastAccessTime.Value, lastAccessTimePtr, false);
            }

            var lastWriteTimePtr = IntPtr.Zero;
            if (lastWriteTime.HasValue)
            {
                lastWriteTimePtr = Marshal.AllocCoTaskMem(Marshal.SizeOf<long>());
                Marshal.StructureToPtr(lastWriteTime.Value, lastWriteTimePtr, false);
            }

            try
            {
                return Throw(PStorage.SetElementTimes(name, creationTimePtr, lastAccessTimePtr, lastWriteTimePtr), throwOnError);
            }
            finally
            {
                if (creationTimePtr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(creationTimePtr);
                }

                if (lastAccessTimePtr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(lastAccessTimePtr);
                }

                if (lastWriteTimePtr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(lastWriteTimePtr);
                }
            }
        }

        public IStream CreateStream(string name, STGM mode, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Throw(PStorage.CreateStream(name, mode, 0, 0, out var stream), throwOnError);
            return stream;
        }

        public IStream OpenNativeStream(string name, STGM mode, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Throw(PStorage.OpenStream(name, IntPtr.Zero, mode, 0, out var stream), throwOnError);
            return stream;
        }

        public Stream OpenStream(string name, STGM mode, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Throw(PStorage.OpenStream(name, IntPtr.Zero, mode, 0, out var stream), throwOnError);
            if (stream == null)
                return null;

            return new StreamOnIStream(stream, true);
        }

        public Storage CreateStorage(string name, STGM mode, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Throw(PStorage.CreateStorage(name, mode, 0, 0, out var storage), throwOnError);
            if (storage == null)
                return null;

            return new Storage(storage);
        }

        public Storage OpenStorage(string name, STGM mode, bool throwOnError = true)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Throw(PStorage.OpenStorage(name, null, mode, IntPtr.Zero, 0, out var storage), throwOnError);
            if (storage == null)
                return null;

            return new Storage(storage);
        }

        public PropertyStorage OpenPropertyStorage(Guid fmtid, STGM mode, bool throwOnError = true)
        {
            var ps = PSetStorage;
            var hr = Throw(ps.Open(fmtid, mode, out var storage), throwOnError);
            if (hr < 0 || storage == null)
                return null;

            return new PropertyStorage(storage);
        }

        public PropertyStorage CreatePropertyStorage(Guid fmtid, Guid clsid, PROPSETFLAG flags, STGM mode, bool throwOnError = true)
        {
            var ps = PSetStorage;
            var hr = Throw(ps.Create(fmtid, clsid, flags, mode, out var storage), throwOnError);
            if (hr < 0 || storage == null)
                return null;

            return new PropertyStorage(storage);
        }

        public int DeletePropertyStorage(Guid fmtid, bool throwOnError = true) => Throw(PSetStorage.Delete(fmtid), throwOnError);

        internal static int Throw(int hr, bool throwOnError, string name = null)
        {
            if (hr < 0 && throwOnError)
            {
                if (name == null)
                    throw new Win32Exception(hr);

                throw new Win32Exception(hr, new Win32Exception(hr).Message.Replace("%1", "'" + name + "'"));
            }

            return hr;
        }

        private IPropertySetStorage PSetStorage
        {
            get
            {
                var ptr = _psetStorage;
                if (ptr == null)
                    throw new ObjectDisposedException(nameof(NativeObject));

                return ptr;
            }
        }

        private IStorage PStorage
        {
            get
            {
                var ptr = _storage;
                if (ptr == null)
                    throw new ObjectDisposedException(nameof(NativeObject));

                return ptr;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct STATPROPSTG
        {
            public IntPtr lpwstrName;
            public int propid;
            public PropertyType vt;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct STATPROPSETSTG
        {
            public Guid fmtid;
            public Guid clsid;
            public PROPSETFLAG grfFlags;
            public FILETIME mtime;
            public FILETIME ctime;
            public FILETIME atime;
            public int dwOSVersion;
        }

        internal enum PRSPEC
        {
            PRSPEC_LPWSTR = 0,
            PRSPEC_PROPID = 1
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct PROPSPECunion
        {
            [FieldOffset(0)]
            public int propid;
            [FieldOffset(0)]
            public IntPtr lpwstr;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct PROPSPEC
        {
            public PRSPEC ulKind;
            public PROPSPECunion union;
        }

        [ComImport, Guid("0000013B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IEnumSTATPROPSETSTG
        {
            [PreserveSig]
            int Next(int celt, ref STATPROPSETSTG rgelt, out int pceltFetched);
            // rest ommited
        }

        [ComImport, Guid("00000139-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IEnumSTATPROPSTG
        {
            [PreserveSig]
            int Next(int celt, ref STATPROPSTG rgelt, out int pceltFetched);
            // rest ommited
        }

        [ComImport, Guid("0000000d-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IEnumSTATSTG
        {
            [PreserveSig]
            int Next(int celt, ref STATSTG rgelt, out int pceltFetched);
            // rest ommited
        }

        [ComImport, Guid("00000138-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IPropertyStorage
        {
            [PreserveSig]
            int ReadMultiple(int cpspec, ref PROPSPEC rgpspec, [In, Out] PropVariant rgpropvar); // only works for cpspec = 1

            [PreserveSig]
            int WriteMultiple(int cpspec, ref PROPSPEC rgpspec, PropVariant rgpropvar, int propidNameFirst); // only works for cpspec = 1

            [PreserveSig]
            int DeleteMultiple(int cpspec, ref PROPSPEC rgpspec); // only works for cpspec = 1

            [PreserveSig]
            int ReadPropertyNames(int cpropid, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid, [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 0)] string[] rglpwstrName);

            [PreserveSig]
            int WritePropertyNames(int cpropid, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] uint[] rgpropid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 0)] string[] rglpwstrName);

            [PreserveSig]
            int DeletePropertyNames();

            [PreserveSig]
            int Commit(STGC grfCommitFlags);

            [PreserveSig]
            int Revert();

            [PreserveSig]
            int Enum(out IEnumSTATPROPSTG ppenum);

            [PreserveSig]
            int SetTimes(ref FILETIME pctime, ref FILETIME patime, ref FILETIME pmtime);

            [PreserveSig]
            int SetClass([MarshalAs(UnmanagedType.LPStruct)] Guid clsid);

            [PreserveSig]
            int Stat(ref STATPROPSETSTG pstatpsstg);
        }

        [ComImport, Guid("0000013A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IPropertySetStorage
        {
            [PreserveSig]
            int Create([MarshalAs(UnmanagedType.LPStruct)] Guid rfmtid, [MarshalAs(UnmanagedType.LPStruct)] Guid pclsid, PROPSETFLAG grfFlags, STGM grfMode, out IPropertyStorage ppprstg);

            [PreserveSig]
            int Open([MarshalAs(UnmanagedType.LPStruct)] Guid rfmtid, STGM grfMode, out IPropertyStorage ppprstg);

            [PreserveSig]
            int Delete([MarshalAs(UnmanagedType.LPStruct)] Guid rfmtid);

            [PreserveSig]
            int Enum(out IEnumSTATPROPSETSTG ppenum);
        }

        [ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IStorage
        {
            [PreserveSig]
            int CreateStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, STGM grfMode, int reserved1, int reserved2, out IStream ppstm);

            [PreserveSig]
            int OpenStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr reserved1, STGM grfMode, int reserved2, out IStream ppstm);

            [PreserveSig]
            int CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, STGM grfMode, int reserved1, int reserved2, out IStorage ppstg);

            [PreserveSig]
            int OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, int reserved, out IStorage ppstg);

            [PreserveSig]
            int CopyTo(int ciidExclude, Guid[] rgiidExclude, IntPtr snbExclude, IStorage pstgDest);

            [PreserveSig]
            int MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgDest, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName, STGMOVE grfFlags);

            [PreserveSig]
            int Commit(STGC grfCommitFlags);

            [PreserveSig]
            int Revert();

            [PreserveSig]
            int EnumElements(int reserved1, IntPtr reserved2, int reserved3, out IEnumSTATSTG ppenum);

            [PreserveSig]
            int DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

            [PreserveSig]
            int RenameElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsOldName, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName);

            [PreserveSig]
            int SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr pctime, IntPtr patime, ref IntPtr pmtime);

            [PreserveSig]
            int SetClass([MarshalAs(UnmanagedType.LPStruct)] Guid clsid);

            [PreserveSig]
            int SetStateBits(int grfStateBits, int grfMask);

            [PreserveSig]
            int Stat(ref STATSTG pstatstg, STATFLAG grfStatFlag);
        }

        [DllImport("ole32", CharSet = CharSet.Unicode)]
        private static extern int StgOpenStorageEx(string pwcsName, STGM grfMode, STGFMT stgfmt, int grfAttrs, IntPtr pStgOptions, IntPtr reserved2, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppObjectOpen);

        public void Dispose()
        {
            Interlocked.Exchange(ref _storage, null);
            Interlocked.Exchange(ref _psetStorage, null);
        }
    }
}
