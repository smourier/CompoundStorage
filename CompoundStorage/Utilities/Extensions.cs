using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace CompoundStorage.Utilities
{
    internal static class Extensions
    {
        private const long _ticksPerMillisecond = 10000;
        private const long _ticksPerSecond = _ticksPerMillisecond * 1000;
        private const long _ticksPerMinute = _ticksPerSecond * 60;
        private const long _ticksPerHour = _ticksPerMinute * 60;
        private const long _ticksPerDay = _ticksPerHour * 24;
        private const int _daysPerYear = 365;
        private const int _daysPer4Years = _daysPerYear * 4 + 1;
        private const int _daysPer100Years = _daysPer4Years * 25 - 1;
        private const int _daysPer400Years = _daysPer100Years * 4 + 1;
        private const int _daysTo1601 = _daysPer400Years * 4;
        private const long _fileTimeOffset = _daysTo1601 * _ticksPerDay;

        public static long ToPositiveFileTime(this DateTime dt)
        {
            var ft = ToFileTime(dt);
            return ft < 0 ? 0 : ft;
        }

        public static long ToFileTime(this DateTime dt)
        {
            var ticks = dt.Kind != DateTimeKind.Utc ? dt.ToUniversalTime().Ticks : dt.Ticks;
            ticks -= _fileTimeOffset;
            return ticks;
        }

        public static DateTimeOffset ToDateTimeOffset(this FILETIME fileTime)
        {
            var ft = (((long)fileTime.dwHighDateTime) << 32) + fileTime.dwLowDateTime;
            return DateTimeOffset.FromFileTime(ft);
        }

        public static long CopyTo(this Stream input, Stream output, long count = long.MaxValue, int bufferSize = 0x14000)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            if (output == null)
                throw new ArgumentNullException(nameof(output));

            if (count <= 0)
                throw new ArgumentException(null, nameof(count));

            if (bufferSize <= 0)
                throw new ArgumentException(null, nameof(bufferSize));

            if (count < bufferSize)
            {
                bufferSize = (int)count;
            }

            var bytes = new byte[bufferSize];
            var total = 0;
            do
            {
                var max = (int)Math.Min(count - total, bytes.Length);
                var read = input.Read(bytes, 0, max);
                if (read == 0)
                    break;

                output.Write(bytes, 0, read);
                total += read;
                if (total == count)
                    break;
            }
            while (true);
            return total;
        }
    }
}
