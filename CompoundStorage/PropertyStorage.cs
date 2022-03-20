using System;
using System.Collections.Generic;
using CompoundStorage.Utilities;

namespace CompoundStorage
{
    public sealed class PropertyStorage
    {
        private readonly Storage.IPropertyStorage _storage;

        internal PropertyStorage(Storage.IPropertyStorage storage)
        {
            _storage = storage;
            var stat = new Storage.STATPROPSETSTG();
            _storage.Stat(ref stat);
            FmtId = stat.fmtid;
        }

        public object NativeObject => _storage;
        public Guid FmtId { get; }

        public IEnumerable<Property> Properties
        {
            get
            {
                var ps = PStorage;
                var hr = ps.Enum(out var enumerator);
                if (hr < 0)
                    yield break;

                do
                {
                    var stat = new Storage.STATPROPSTG();
                    enumerator.Next(1, ref stat, out var fetched);
                    if (fetched != 1)
                        break;

                    var spec = new Storage.PROPSPEC();
                    if (stat.lpwstrName != IntPtr.Zero)
                    {
                        spec.ulKind = Storage.PRSPEC.PRSPEC_LPWSTR;
                        spec.union.lpwstr = stat.lpwstrName;
                    }
                    else
                    {
                        spec.ulKind = Storage.PRSPEC.PRSPEC_PROPID;
                        spec.union.propid = stat.propid;
                    }

                    using (var pv = new PropVariant())
                    {
                        _storage.ReadMultiple(1, ref spec, pv);
                        yield return new Property(FmtId, stat, pv.Value);
                    }
                }
                while (true);
            }
        }

        private Storage.IPropertyStorage PStorage
        {
            get
            {
                var ptr = _storage;
                if (ptr == null)
                    throw new ObjectDisposedException(nameof(NativeObject));

                return ptr;
            }
        }

        public override string ToString() => FmtId.ToString("B");
        public int Commit(STGC flags, bool throwOnError = true) => Storage.Throw(PStorage.Commit(flags), throwOnError);
    }
}
