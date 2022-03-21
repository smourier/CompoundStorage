using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace CompoundStorage.Utilities
{
    public sealed class MemoryPropertyStore : IDisposable
    {
        private IPersistSerializedPropStorage _storage;

        public MemoryPropertyStore()
        {
            Storage.Throw(PSCreateMemoryPropertyStore(typeof(IPersistSerializedPropStorage).GUID, out var storage), true);
            _storage = storage;
        }

        public object NativeObject => _storage;

        public void Dispose()
        {
            Interlocked.Exchange(ref _storage, null);
        }

        private IPersistSerializedPropStorage PStorage
        {
            get
            {
                var ptr = _storage;
                if (ptr == null)
                    throw new ObjectDisposedException(nameof(NativeObject));

                return ptr;
            }
        }

        public byte[] Serialize(bool throwOnError = true)
        {
            Storage.Throw(PStorage.GetPropertyStorage(out var ptr, out var size), throwOnError);
            if (size == 0)
                return Array.Empty<byte>();

            var bytes = new byte[size];
            try
            {
                Marshal.Copy(ptr, bytes, 0, size);
                return bytes;
            }
            finally
            {
                Marshal.FreeCoTaskMem(ptr);
            }
        }

        public void Save(string filePath, bool clearDirty = true, bool throwOnError = true)
        {
            if (filePath == null)
                throw new ArgumentNullException(nameof(filePath));

            using (var file = File.OpenWrite(filePath))
                Save(file, clearDirty, throwOnError);
        }

        public void Save(Stream stream, bool clearDirty = true, bool throwOnError = true)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            var pstream = (Link.IPersistStream)PStorage;
            Storage.Throw(pstream.Save(new ManagedIStream(stream), clearDirty), throwOnError);
        }

        public static MemoryPropertyStore Load(string filePath, bool throwOnError = true)
        {
            if (filePath == null)
                throw new ArgumentNullException(nameof(filePath));

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                return Load(file, throwOnError);
        }

        public static MemoryPropertyStore Load(Stream stream, bool throwOnError = true)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            var ps = new MemoryPropertyStore();
            var pstream = (Link.IPersistStream)ps.PStorage;
            Storage.Throw(pstream.Load(new ManagedIStream(stream)), throwOnError);
            return ps;
        }

        private enum PERSIST_SPROPSTORE_FLAGS
        {
            FPSPS_DEFAULT = 0,
            FPSPS_READONLY = 0x1,
            FPSPS_TREAT_NEW_VALUES_AS_DIRTY = 0x2,
        };

        [ComImport, Guid("e318ad57-0aa0-450f-aca5-6fab7103d917"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistSerializedPropStorage
        {
            [PreserveSig]
            int SetFlags(PERSIST_SPROPSTORE_FLAGS flags);

            [PreserveSig]
            int SetPropertyStorage(byte[] psps, int cb);

            [PreserveSig]
            int GetPropertyStorage(out IntPtr ppsps, out int pcb);
        }

        [DllImport("propsys")]
        private extern static int PSCreateMemoryPropertyStore([MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IPersistSerializedPropStorage storage);
    }
}
