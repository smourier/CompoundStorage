using System;
using System.Runtime.InteropServices;

namespace CompoundStorage
{
    public sealed class Property
    {
        internal Property(Guid fmtid, Storage.STATPROPSTG stat, object value)
        {
            FmtId = fmtid;
            Id = stat.propid;
            Type = stat.vt;
            Value = value;
            if (stat.lpwstrName != IntPtr.Zero)
            {
                Name = Marshal.PtrToStringUni(stat.lpwstrName);
                Marshal.FreeCoTaskMem(stat.lpwstrName);
            }
        }

        public string Name { get; }
        public Guid FmtId { get; }
        public int Id { get; }
        public PropertyType Type { get; }
        public object Value { get; }

        public override string ToString() => FmtId.ToString("B") + " " + Id + " => " + Value;
    }
}
