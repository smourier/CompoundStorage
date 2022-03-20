using System;
using System.Runtime.InteropServices.ComTypes;

namespace CompoundStorage
{
    public sealed class StorageElement : Element
    {
        internal StorageElement(Storage.IStorage storage, STATSTG stat)
            : base(storage, stat)
        {
        }

        public object OpenStorage(STGM mode = STGM.STGM_SHARE_EXCLUSIVE, bool throwOnError = true)
        {
            var hr = Storage.Throw(PStorage.OpenStorage(Name, null, mode, IntPtr.Zero, 0, out var storage), throwOnError, Name);
            if (hr < 0 || storage == null)
                return null;

            return storage;
        }
    }
}
