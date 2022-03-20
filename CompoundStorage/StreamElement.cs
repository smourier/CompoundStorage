using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using CompoundStorage.Utilities;

namespace CompoundStorage
{
    public sealed class StreamElement : Element
    {
        internal StreamElement(Storage.IStorage storage, STATSTG stat)
            : base(storage, stat)
        {
        }

        public IStream OpenNativeStream(STGM mode = STGM.STGM_SHARE_EXCLUSIVE, bool throwOnError = true)
        {
            var hr = Storage.Throw(PStorage.OpenStream(Name, IntPtr.Zero, mode, 0, out var stream), throwOnError, Name);
            if (hr < 0 || stream == null)
                return null;

            return stream;
        }

        public Stream OpenStream(STGM mode = STGM.STGM_SHARE_EXCLUSIVE, bool throwOnError = true)
        {
            var stream = OpenNativeStream(mode, throwOnError);
            if (stream == null)
                return null;

            return new StreamOnIStream(stream, true);
        }
    }
}
