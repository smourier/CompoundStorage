using System;

namespace CompoundStorage
{
    public sealed class ClipData
    {
        public ClipData()
        {
        }

        public ClipData(Guid fmtid)
        {
            Fmtid = fmtid;
        }

        public ClipData(int format)
        {
            Format = format;
        }

        public ClipData(string format)
        {
            StringFormat = format;
        }

        public Guid Fmtid { get; }
        public int Format { get; }
        public string StringFormat { get; }
        public byte[] Data { get; set; }

        public override string ToString()
        {
            if (Fmtid != Guid.Empty)
                return Fmtid.ToString("B");

            if (StringFormat != null)
                return StringFormat;

            if (Format == 0)
                return string.Empty;

            return Format.ToString();
        }
    }
}
