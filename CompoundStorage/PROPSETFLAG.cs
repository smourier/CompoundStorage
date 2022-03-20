using System;

namespace CompoundStorage
{
    [Flags]
    public enum PROPSETFLAG
    {
        PROPSETFLAG_DEFAULT = 0,
        PROPSETFLAG_NONSIMPLE = 0x1,
        PROPSETFLAG_ANSI = 0x2,
        PROPSETFLAG_UNBUFFERED = 0x4,
        PROPSETFLAG_CASE_SENSITIVE = 0x8,
    }
}
