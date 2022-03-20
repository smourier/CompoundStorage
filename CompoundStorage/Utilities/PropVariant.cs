using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CompoundStorage.Utilities
{
    [StructLayout(LayoutKind.Explicit)]
    public sealed class PropVariant : IDisposable
    {
#pragma warning disable IDE0044 // Add readonly modifier
        [FieldOffset(0)]
        private PropertyType _vt;

        [FieldOffset(8)]
        private IntPtr _ptr;

        [FieldOffset(8)]
        private int _int32;

        [FieldOffset(8)]
        private uint _uint32;

        [FieldOffset(8)]
        private byte _byte;

        [FieldOffset(8)]
        private sbyte _sbyte;

        [FieldOffset(8)]
        private short _int16;

        [FieldOffset(8)]
        private ushort _uint16;

        [FieldOffset(8)]
        private long _int64;

        [FieldOffset(8)]
        private ulong _uint64;

        [FieldOffset(8)]
        private double _double;

        [FieldOffset(8)]
        private float _single;

        [FieldOffset(8)]
        private short _boolean;

        [FieldOffset(8)]
        private FILETIME _filetime;

        [FieldOffset(8)]
        private PROPARRAY _ca;

        [FieldOffset(0)]
        private decimal _decimal;
#pragma warning restore IDE0044 // Add readonly modifier

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPARRAY
        {
            public int cElems;
            public IntPtr pElems;
        }

        private static readonly Lazy<int> _sizeOfVariant = new Lazy<int>(() => FindSizeOfVariant.SizeOf, false);

        public static int SizeOfVariant => _sizeOfVariant.Value;

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct FindSizeOfVariant
        {
            [MarshalAs(UnmanagedType.Struct)]
            public object var;
            public byte b;

            public static int SizeOf => (int)Marshal.OffsetOf(typeof(FindSizeOfVariant), nameof(b));
        }

        public PropVariant()
        {
        }

        public static object GetObjectForNative(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero)
                return null;

            var pv = Marshal.PtrToStructure<PropVariant>(ptr);
            var value = pv.Value;

            pv._vt = PropertyType.VT_EMPTY;
            pv._ptr = IntPtr.Zero;
            return value;
        }

        public static PropVariant ToLPSTR(string text)
        {
            if (text == null)
                return Null;

            var pv = new PropVariant
            {
                _ptr = Marshal.StringToCoTaskMemAnsi(text),
                _vt = PropertyType.VT_LPSTR
            };
            return pv;
        }

        public static PropVariant ToBSTR(string text)
        {
            if (text == null)
                return Null;

            var pv = new PropVariant
            {
                _ptr = Marshal.StringToBSTR(text),
                _vt = PropertyType.VT_BSTR
            };
            return pv;
        }

        private const uint DISP_E_PARAMNOTFOUND = 0x80020004;
        public static readonly PropVariant Empty = new PropVariant() { _vt = PropertyType.VT_EMPTY };
        public static readonly PropVariant Missing = new PropVariant() { _vt = PropertyType.VT_HRESULT, _uint32 = DISP_E_PARAMNOTFOUND };
        public static readonly PropVariant Null = new PropVariant() { _vt = PropertyType.VT_NULL };

        private void ConstructBlob(byte[] blob)
        {
            ConstructVector(blob, typeof(byte), PropertyType.VT_UI1);
            _vt = PropertyType.VT_BLOB;
        }

        private void ConstructVector(Array array)
        {
            if (array is bool[] bools)
            {
                var shorts = new short[bools.Length];
                for (var i = 0; i < bools.Length; i++)
                {
                    shorts[i] = bools[i] ? ((short)(-1)) : ((short)0);
                }
                ConstructVector(shorts, typeof(short), PropertyType.VT_BOOL);
                return;
            }

            var et = array.GetType().GetElementType();
            ConstructVector(array, et, FromType(et));
        }

        private void ConstructVector(Array array, Type type, PropertyType vt)
        {
            _vt = vt | PropertyType.VT_VECTOR;
            if (array.Length > 0)
            {
                int size;
                if (type == typeof(string))
                {
                    size = IntPtr.Size;
                }
                else
                {
                    size = Marshal.SizeOf(type);
                }

                size *= array.Length;
                var ptr = Marshal.AllocCoTaskMem(size);
                _ca.cElems = array.Length;
                _ca.pElems = ptr;

                if (type == typeof(string))
                {
                    for (var i = 0; i < array.Length; i++)
                    {
                        var str = Marshal.StringToCoTaskMemUni((string)array.GetValue(i));
                        Marshal.WriteIntPtr(ptr, IntPtr.Size * i, str);
                    }
                }
                else
                {
                    CopyMemory(ptr, Marshal.UnsafeAddrOfPinnedArrayElement(array, 0), size);
                }
            }
        }

        private static void Using(object resource, Action action)
        {
            try
            {
                action();
            }
            finally
            {
                (resource as IDisposable)?.Dispose();
            }
        }

        private static int GetCount(IEnumerable enumerable)
        {
            if (enumerable is ICollection col)
                return col.Count;

            var count = 0;
            var e = enumerable.GetEnumerator();
            Using(e, () =>
            {
                while (e.MoveNext())
                {
                    count++;
                }
            });
            return count;
        }

        private static Type GetElementType(Type collectionType)
        {
            foreach (Type iface in collectionType.GetInterfaces())
            {
                if (!iface.IsGenericType)
                    continue;

                if (iface.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                    return iface.GetGenericArguments()[0];

                if (iface.GetGenericTypeDefinition() == typeof(ICollection<>))
                    return iface.GetGenericArguments()[0];

                if (iface.GetGenericTypeDefinition() == typeof(IList<>))
                    return iface.GetGenericArguments()[0];
            }
            return null;
        }

        private static Type GetElementType(IEnumerable enumerable)
        {
            var et = GetElementType(enumerable.GetType());
            if (et != null)
                return et;

            foreach (var obj in enumerable)
            {
                return obj.GetType();
            }
            return null;
        }

        private void ConstructEnumerable(IEnumerable enumerable)
        {
            var et = GetElementType(enumerable);
            if (et == null)
                throw new Exception("Enumerable type '" + enumerable.GetType().FullName + "' is not supported.");

            var count = GetCount(enumerable);
            var array = Array.CreateInstance(et, count);
            var i = 0;
            foreach (var obj in enumerable)
            {
                array.SetValue(obj, i++);
            }
            ConstructVector(array);
        }

        private static Type FromType(PropertyType type)
        {
            switch (type)
            {
                case PropertyType.VT_I1:
                    return typeof(sbyte);

                case PropertyType.VT_UI1:
                    return typeof(byte);

                case PropertyType.VT_I2:
                    return typeof(short);

                case PropertyType.VT_UI2:
                    return typeof(ushort);

                case PropertyType.VT_UI4:
                case PropertyType.VT_UINT:
                    return typeof(uint);

                case PropertyType.VT_I8:
                    return typeof(long);

                case PropertyType.VT_UI8:
                    return typeof(ulong);

                case PropertyType.VT_R4:
                    return typeof(float);

                case PropertyType.VT_R8:
                    return typeof(double);

                case PropertyType.VT_BOOL:
                    return typeof(bool);

                case PropertyType.VT_I4:
                case PropertyType.VT_INT:
                case PropertyType.VT_ERROR:
                    return typeof(int);

                case PropertyType.VT_DATE:
                    return typeof(DateTime);

                case PropertyType.VT_FILETIME:
                    return typeof(FILETIME);

                case PropertyType.VT_BLOB:
                    return typeof(byte[]);

                case PropertyType.VT_CLSID:
                    return typeof(Guid);

                case PropertyType.VT_BSTR:
                case PropertyType.VT_LPSTR:
                case PropertyType.VT_LPWSTR:
                    return typeof(string);

                case PropertyType.VT_UNKNOWN:
                case PropertyType.VT_DISPATCH:
                    return typeof(object);

                case PropertyType.VT_CY:
                case PropertyType.VT_DECIMAL:
                    return typeof(decimal);

                case PropertyType.VT_CF:
                    return typeof(ClipData);

                default:
                    throw new Exception("Property type " + type + " is not supported.");
            }
        }

        private static PropertyType FromType(Type type)
        {
            if (type == null)
                return PropertyType.VT_NULL;

            var tc = Type.GetTypeCode(type);
            switch (tc)
            {
                case TypeCode.Boolean:
                    return PropertyType.VT_BOOL;

                case TypeCode.Byte:
                    return PropertyType.VT_UI1;

                case TypeCode.Char:
                    return PropertyType.VT_LPWSTR;

                case TypeCode.DateTime:
                    return PropertyType.VT_FILETIME;

                case TypeCode.DBNull:
                    return PropertyType.VT_NULL;

                case TypeCode.Decimal:
                    return PropertyType.VT_DECIMAL;

                case TypeCode.Double:
                    return PropertyType.VT_R8;

                case TypeCode.Empty:
                    return PropertyType.VT_EMPTY;

                case TypeCode.Int16:
                    return PropertyType.VT_I2;

                case TypeCode.Int32:
                    return PropertyType.VT_I4;

                case TypeCode.Int64:
                    return PropertyType.VT_I8;

                case TypeCode.SByte:
                    return PropertyType.VT_I1;

                case TypeCode.Single:
                    return PropertyType.VT_R4;

                case TypeCode.String:
                    return PropertyType.VT_LPWSTR;

                case TypeCode.UInt16:
                    return PropertyType.VT_UI2;

                case TypeCode.UInt32:
                    return PropertyType.VT_UI4;

                case TypeCode.UInt64:
                    return PropertyType.VT_UI8;

                default:
                    if (type == typeof(Guid))
                        return PropertyType.VT_CLSID;

                    if (type == typeof(FILETIME))
                        return PropertyType.VT_FILETIME;

                    if (type == typeof(ClipData))
                        return PropertyType.VT_CF;

                    throw new Exception("Value of type '" + type.FullName + "' is not supported.");
            }
        }

        public PropVariant(object value)
        {
            if (value == null)
            {
                _vt = PropertyType.VT_NULL;
                return;
            }

            if (value == Type.Missing)
            {
                _vt = PropertyType.VT_EMPTY;
                return;
            }

            if (Marshal.IsComObject(value))
            {
                _vt = PropertyType.VT_EMPTY;
                return;
            }

            if (value is PropVariant pv)
            {
                value = pv.Value;
            }

            if (value is char[] chars)
            {
                value = new string(chars);
            }

            if (value is char[][] charray)
            {
                var strings = new string[charray.GetLength(0)];
                for (var i = 0; i < charray.Length; i++)
                {
                    strings[i] = new string(charray[i]);
                }
                value = strings;
            }

            if (value is Array array)
            {
                ConstructVector(array);
                return;
            }

            if (!(value is string) && value is IEnumerable enumerable)
            {
                ConstructEnumerable(enumerable);
                return;
            }

            var tc = Type.GetTypeCode(value.GetType());
            switch (tc)
            {
                case TypeCode.Boolean:
                    _boolean = (bool)value ? (short)(-1) : (short)0;
                    break;

                case TypeCode.Byte:
                    _byte = (byte)value;
                    break;

                case TypeCode.Char:
                    chars = new[] { (char)value };
                    _ptr = Marshal.StringToCoTaskMemUni(new string(chars));
                    break;

                case TypeCode.DateTime:
                    var ft = ((DateTime)value).ToPositiveFileTime();
                    if (ft == 0)
                        break;

                    InitPropVariantFromFileTime(ref ft, this);
                    break;

                case TypeCode.Empty:
                case TypeCode.DBNull:
                    break;

                case TypeCode.Decimal:
                    _decimal = (decimal)value;
                    break;

                case TypeCode.Double:
                    _double = (double)value;
                    break;

                case TypeCode.Int16:
                    _int16 = (short)value;
                    break;

                case TypeCode.Int32:
                    _int32 = (int)value;
                    break;

                case TypeCode.Int64:
                    _int64 = (long)value;
                    break;

                case TypeCode.SByte:
                    _sbyte = (sbyte)value;
                    break;

                case TypeCode.Single:
                    _single = (float)value;
                    break;

                case TypeCode.String:
                    _ptr = Marshal.StringToCoTaskMemUni((string)value);
                    break;

                case TypeCode.UInt16:
                    _uint16 = (ushort)value;
                    break;

                case TypeCode.UInt32:
                    _uint32 = (uint)value;
                    break;

                case TypeCode.UInt64:
                    _uint64 = (ulong)value;
                    break;

                default:
                    if (value is Guid guid)
                    {
                        _ptr = Marshal.AllocCoTaskMem(16);
                        Marshal.Copy(guid.ToByteArray(), 0, _ptr, 16);
                        break;
                    }

                    if (value is FILETIME fileTime)
                    {
                        _filetime = fileTime;
                        break;
                    }
                    throw new Exception("Value of type '" + value.GetType().FullName + "' is not supported.");
            }

            _vt = FromType(value.GetType());
        }

        public PropertyType VarType { get => _vt; set => _vt = value; }
        public object Value
        {
            get
            {
                switch (_vt)
                {
                    case PropertyType.VT_EMPTY:
                    case PropertyType.VT_NULL:
                        return null;

                    case PropertyType.VT_I1:
                        return _sbyte;

                    case PropertyType.VT_UI1:
                        return _byte;

                    case PropertyType.VT_I2:
                        return _int16;

                    case PropertyType.VT_UI2:
                        return _uint16;

                    case PropertyType.VT_I4:
                    case PropertyType.VT_INT:
                        return _int32;

                    case PropertyType.VT_UI4:
                    case PropertyType.VT_UINT:
                        return _uint32;

                    case PropertyType.VT_I8:
                        return _int64;

                    case PropertyType.VT_UI8:
                        return _uint64;

                    case PropertyType.VT_R4:
                        return _single;

                    case PropertyType.VT_R8:
                        return _double;

                    case PropertyType.VT_BOOL:
                        return _int32 != 0;

                    case PropertyType.VT_ERROR:
                        return _int64;

                    case PropertyType.VT_CY:
                        return _decimal;

                    case PropertyType.VT_DATE:
                        return DateTime.FromOADate(_double);

                    case PropertyType.VT_FILETIME:
                        return DateTime.FromFileTime(_int64);

                    case PropertyType.VT_BSTR:
                        return Marshal.PtrToStringBSTR(_ptr);

                    case PropertyType.VT_BLOB:
                        var blob = new byte[_ca.cElems];
                        Marshal.Copy(_ca.pElems, blob, 0, _int32);
                        return blob;

                    case PropertyType.VT_CLSID:
                        var guid = new byte[16];
                        Marshal.Copy(_ptr, guid, 0, guid.Length);
                        return new Guid(guid);

                    case PropertyType.VT_LPSTR:
                        return Marshal.PtrToStringAnsi(_ptr);

                    case PropertyType.VT_LPWSTR:
                        return Marshal.PtrToStringUni(_ptr);

                    case PropertyType.VT_UNKNOWN:
                    case PropertyType.VT_DISPATCH:
                        return Marshal.GetObjectForIUnknown(_ptr);

                    case PropertyType.VT_DECIMAL:
                        return _decimal;

                    case PropertyType.VT_CF:
                        if (_ptr == IntPtr.Zero)
                            return null;

                        var cd = Marshal.PtrToStructure<CLIPDATA>(_ptr);
                        switch (cd.ulClipFmt)
                        {
                            case -1:
                            case -2:
                                var cf = Marshal.ReadInt32(cd.pClipData);
                                var size = cd.cbSize - 4;
                                var data = new ClipData(cf);
                                data.Data = new byte[size];
                                Marshal.Copy(cd.pClipData + 4, data.Data, 0, size);
                                return data;

                            case -3:
                                Guid guidData;
                                if (cd.pClipData == IntPtr.Zero)
                                {
                                    guidData = Guid.Empty;
                                }
                                else
                                {
                                    guidData = Marshal.PtrToStructure<Guid>(cd.pClipData);
                                }
                                return new ClipData(guidData);

                            default:
                                if (cd.ulClipFmt > 0 && cd.pClipData != IntPtr.Zero)
                                {
                                    var s = Marshal.PtrToStringUni(cd.pClipData);
                                    return new ClipData(s);
                                }
                                break;
                        }
                        return new ClipData();

                    default:
                        if ((_vt & PropertyType.VT_VECTOR) == PropertyType.VT_VECTOR)
                        {
                            var et = _vt & ~PropertyType.VT_VECTOR;
                            if (TryGetVectorValue(et, out object vector))
                                return vector;
                        }
                        throw new Exception("Value of property type " + _vt + " is not supported.");
                }
            }
        }

        public void Dispose() { PropVariantClear(this); GC.SuppressFinalize(this); }
        ~PropVariant() => Dispose();

        private bool TryGetVectorValue(PropertyType vt, out object value)
        {
            value = null;
            var ret = false;
            int size;
            switch (vt)
            {
                case PropertyType.VT_LPSTR:
                case PropertyType.VT_LPWSTR:
                    var strings = new string[_ca.cElems];
                    for (var i = 0; i < strings.Length; i++)
                    {
                        var str = Marshal.ReadIntPtr(_ca.pElems, IntPtr.Size * i);
                        strings[i] = vt == PropertyType.VT_LPSTR ? Marshal.PtrToStringAnsi(str) : Marshal.PtrToStringUni(str);
                    }

                    value = strings;
                    ret = true;
                    break;

                case PropertyType.VT_BSTR:
                    var bstrs = new string[_ca.cElems];
                    for (var i = 0; i < bstrs.Length; i++)
                    {
                        var str = Marshal.ReadIntPtr(_ca.pElems, IntPtr.Size * i);
                        bstrs[i] = Marshal.PtrToStringBSTR(str);
                    }

                    value = bstrs;
                    ret = true;
                    break;

                case PropertyType.VT_BOOL:
                    var shorts = new short[_ca.cElems];
                    size = _ca.cElems * Marshal.SizeOf(typeof(short));
                    CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(shorts, 0), _ca.pElems, size);
                    var bools = new bool[shorts.Length];
                    for (var i = 0; i < shorts.Length; i++)
                    {
                        bools[i] = shorts[i] != 0;
                    }

                    value = bools;
                    ret = true;
                    break;

                case PropertyType.VT_VARIANT:
                    var variants = new object[_ca.cElems];
                    for (var i = 0; i < variants.Length; i++)
                    {
                        variants[i] = GetObjectForNative(_ca.pElems + SizeOfVariant * i);
                    }

                    value = variants;
                    ret = true;
                    break;

                case PropertyType.VT_I1:
                case PropertyType.VT_UI1:
                case PropertyType.VT_I2:
                case PropertyType.VT_UI2:
                case PropertyType.VT_I4:
                case PropertyType.VT_INT:
                case PropertyType.VT_UI4:
                case PropertyType.VT_UINT:
                case PropertyType.VT_I8:
                case PropertyType.VT_UI8:
                case PropertyType.VT_R4:
                case PropertyType.VT_R8:
                case PropertyType.VT_ERROR:
                case PropertyType.VT_CY:
                case PropertyType.VT_DATE:
                case PropertyType.VT_FILETIME:
                case PropertyType.VT_CLSID:
                    var et = FromType(vt);
                    var values = Array.CreateInstance(et, _ca.cElems);
                    size = _ca.cElems * Marshal.SizeOf(et);
                    CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(values, 0), _ca.pElems, size);
                    value = values;
                    ret = true;
                    break;
            }
            return ret;
        }

        private const int _int32Size = 4;

        public IntPtr Serialize(out int size)
        {
            var hr = StgSerializePropVariant(this, out IntPtr ptr, out size);
            if (hr != 0)
                throw new Win32Exception(hr);

            return ptr;
        }

        public byte[] Serialize()
        {
            var hr = StgSerializePropVariant(this, out IntPtr ptr, out int size);
            if (hr != 0)
                throw new Win32Exception(hr);

            var bytes = new byte[size];
            Marshal.Copy(ptr, bytes, 0, bytes.Length);
            Marshal.FreeCoTaskMem(ptr);
            return bytes;
        }

        public static PropVariant Deserialize(byte[] bytes, bool throwOnError = true)
        {
            if (bytes == null)
                throw new ArgumentNullException(nameof(bytes));

            var pv = new PropVariant();
            var hr = StgDeserializePropVariant(bytes, bytes.Length, pv);
            if (hr != 0)
            {
                pv.Dispose();
                if (throwOnError)
                    throw new Win32Exception(hr);

                return null;
            }

            return pv;
        }

        public static PropVariant Deserialize(IntPtr ptr, int size, bool throwOnError = true)
        {
            if (ptr == IntPtr.Zero)
                throw new ArgumentNullException(nameof(ptr));

            var pv = new PropVariant();
            var hr = StgDeserializePropVariant(ptr, size, pv);
            if (hr != 0)
            {
                pv.Dispose();
                if (throwOnError)
                    throw new Win32Exception(hr);

                return null;
            }

            return pv;
        }

        public override string ToString()
        {
            var value = Value;
            if (value is string svalue)
                return "'" + svalue + "'";

            if (!(value is byte[]) && value is IEnumerable enumerable)
                return VarType + ": " + string.Join(", ", enumerable.OfType<object>());

            if (value is byte[] bytes)
                return VarType + ": bytes[" + bytes.Length + "]";

            return VarType + ": " + value;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CLIPDATA
        {
            public int cbSize;
            public int ulClipFmt;
            public IntPtr pClipData;
        }

        [DllImport("propsys")]
        private extern static int StgDeserializePropVariant(IntPtr ppProp, int cbMax, [Out] PropVariant ppropvar);

        [DllImport("propsys")]
        private extern static int StgDeserializePropVariant([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] ppProp, int cbMax, [Out] PropVariant ppropvar);

        [DllImport("propsys")]
        private extern static int StgSerializePropVariant(PropVariant ppropvar, out IntPtr ppProp, out int pcb);

        [DllImport("ole32")]
        private extern static int PropVariantClear([In, Out] PropVariant pvar);

        [DllImport("propsys")]
        private static extern int InitPropVariantFromFileTime(ref long pftIn, [Out] PropVariant ppropvar);

        [DllImport("kernel32", EntryPoint = "RtlMoveMemory")]
        private static extern void CopyMemory(IntPtr destination, IntPtr source, int length);
    }
}
