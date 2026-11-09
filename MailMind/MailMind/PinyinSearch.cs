using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text;

namespace MailMind;

internal static class PinyinSearch
{
    private static readonly ConcurrentDictionary<string, string> Cache = new();
    [DllImport("icu.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr utrans_openU(string id, int idLength, int direction, IntPtr rules, int rulesLength, IntPtr parseError, ref int error);
    [DllImport("icu.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern void utrans_transUChars(IntPtr transliterator, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder text, ref int length, int capacity, int start, ref int limit, ref int error);
    [DllImport("icu.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern void utrans_close(IntPtr transliterator);

    public static string Transliterate(string value) => Cache.GetOrAdd(value, Convert);
    private static string Convert(string value)
    {
        try
        {
            const string id = "Han-Latin; Latin-ASCII";
            int error = 0;
            var handle = utrans_openU(id, id.Length, 0, IntPtr.Zero, 0, IntPtr.Zero, ref error);
            if (handle == IntPtr.Zero || error > 0) return value;
            try
            {
                int capacity = Math.Max(64, value.Length * 16 + 1), length = value.Length, limit = length;
                var buffer = new StringBuilder(value, capacity);
                utrans_transUChars(handle, buffer, ref length, capacity, 0, ref limit, ref error);
                return error > 0 ? value : buffer.ToString().Replace(" ", "");
            }
            finally { utrans_close(handle); }
        }
        catch (DllNotFoundException) { return value; }
        catch (EntryPointNotFoundException) { return value; }
    }
}
