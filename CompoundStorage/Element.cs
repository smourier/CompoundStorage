using System;
using System.Runtime.InteropServices.ComTypes;

namespace CompoundStorage
{
    public class Element
    {
        private Storage.IStorage _storage;

        internal Element(Storage.IStorage storage, STATSTG stat)
        {
            _storage = storage;
            Name = stat.pwcsName;
            Size = stat.cbSize;
            Type = (STGTY)stat.type;
        }

        public string Name { get; }
        public long Size { get; }
        public STGTY Type { get; }
        public object NativeObject => _storage;

        internal Storage.IStorage PStorage
        {
            get
            {
                var ptr = _storage;
                if (ptr == null)
                    throw new ObjectDisposedException(nameof(NativeObject));

                return ptr;
            }
        }

        public override string ToString() => Type + " '" + Name + "' (" + Size + ")";

        public int SetTimes(DateTime? creationTime, DateTime? lastAccessTime, DateTime? lasWriteTime, bool throwOnError = true) => new Storage(PStorage).SetElementTimes(Name, creationTime, lastAccessTime, lasWriteTime, throwOnError);
    }
}
