// The in-process executor of smartass batches (engine/smartass/xl/inproc): a template. build.py puts the build's hash
// in place of __HASH__, so every build has its own function names (a loaded add-in's names stay taken in that Excel).
// C# 5 (the .NET Framework compiler that ships with Windows); the batch format is described in batch.py, the result
// record in runner.py, and the COM executor there does the same over COM from Python.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using ExcelDna.Integration;

public static class SmartassRunner___HASH__
{
    const string Hash = "__HASH__";
    const int Version = 1;
    const int Links = 1;                                       // xlExcelLinks
    const int InlineMax = 30000;                               // below Excel's 32 767 characters per string
    const string Unbound = "no object {0}: the operation that makes it failed or did not run";
    const string NotAValue = "a COM object cannot be asked back";
    const string NotADate = "a date cannot be asked back (read Value2)";
    static readonly string[] QuietNames = { "ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation" };
    static readonly object[] QuietValues = { false, false, false, -4135 };
    const BindingFlags Get = BindingFlags.GetProperty | BindingFlags.InvokeMethod;
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    [ExcelFunction(Name = "SmartassPing___HASH__", IsHidden = true)]
    public static object Ping()
    {
        return Hash;
    }

    // Runs the job at jobPath; never throws. The job file is deleted once it is read: from then on the caller never
    // runs the batch again. The result record is the return value when it fits an Excel string (InlineMax);
    // otherwise it is written to jobPath + ".out" (UTF-8, no BOM) and the return value is "file". (Opening a file
    // Excel has just written costs the caller about 20 ms: the virus scanner reads it first.)
    [ExcelFunction(Name = "SmartassRun___HASH__", IsMacroType = true, IsHidden = true)]
    public static object Run(string jobPath)
    {
        var sw = Stopwatch.StartNew();
        var values = new List<KeyValuePair<string, object>>();
        var marks = new List<KeyValuePair<string, object>>();
        var skipped = new List<object>();
        var changed = new List<KeyValuePair<string, object>>();
        var unrestored = new List<object>();
        bool ok = true;
        int started = 0, index = -1;
        string message = null;
        object app = null;
        try
        {
            string text = File.ReadAllText(jobPath, Encoding.UTF8);
            try { File.Delete(jobPath); } catch (Exception) { }
            var json = new JavaScriptSerializer();
            json.MaxJsonLength = int.MaxValue;
            json.RecursionLimit = 64;
            var job = json.Deserialize<Dictionary<string, object>>(text);
            if (Convert.ToInt32(job["version"], Inv) != Version)
                throw new Exception("batch version " + job["version"] + ", this add-in runs version " + Version);
            var ops = (ArrayList)job["ops"];
            var fin = job.ContainsKey("finally") ? (ArrayList)job["finally"] : new ArrayList();
            app = ExcelDnaUtil.Application;
            if (job.ContainsKey("quiet") && (bool)job["quiet"]) Quiet(app, changed);
            var refs = new Dictionary<int, object>();
            refs[0] = app;
            try
            {
                for (int i = 0; i < ops.Count; i++)
                {
                    var op = (Dictionary<string, object>)ops[i];
                    started = i + 1;
                    try { Do(op, refs, values, marks, sw); }
                    catch (Exception e)
                    {
                        if (op.ContainsKey("try") && (bool)op["try"]) { skipped.Add(new object[] { i, Message(e) }); continue; }
                        ok = false; index = i; message = Message(e);
                        break;
                    }
                }
            }
            finally
            {
                for (int j = 0; j < fin.Count; j++)
                {
                    try { Do((Dictionary<string, object>)fin[j], refs, values, marks, sw); }
                    catch (Exception e) { skipped.Add(new object[] { ops.Count + j, Message(e) }); }
                }
                refs.Clear();
            }
        }
        catch (Exception e)
        {
            ok = false; index = -1; message = Message(e);
        }
        finally
        {
            Restore(app, changed, unrestored);
        }
        var sb = new StringBuilder();
        sb.Append("{\"ok\":").Append(ok ? "true" : "false");
        sb.Append(",\"started\":").Append(started.ToString(Inv));
        if (!ok)
        {
            sb.Append(",\"index\":").Append(index.ToString(Inv));
            sb.Append(",\"message\":"); Str(sb, message ?? "");
        }
        sb.Append(",\"values\":"); Pairs(sb, values);
        sb.Append(",\"marks\":"); Pairs(sb, marks);
        sb.Append(",\"ms\":"); Write(sb, sw.Elapsed.TotalMilliseconds);
        sb.Append(",\"skipped\":"); Write(sb, skipped);
        if (unrestored.Count > 0) { sb.Append(",\"unrestored\":"); Write(sb, unrestored); }
        sb.Append("}");
        string record = sb.ToString();
        if (record.Length <= InlineMax) return record;
        try
        {
            File.WriteAllText(jobPath + ".out", record, new UTF8Encoding(false));
        }
        catch (Exception)
        {
            return "no result file";
        }
        return "file";
    }

    // ------------------------------------------------------------------ the settings
    static void Quiet(object app, List<KeyValuePair<string, object>> changed)
    {
        for (int i = 0; i < QuietNames.Length; i++)
        {
            object before;
            try { before = Invoke(app, QuietNames[i], Get, new object[0]); }
            catch (Exception)
            {
                if (QuietNames[i] == "Calculation") continue;        // no workbook open: left alone
                throw;
            }
            if (!Equals(Convert.ChangeType(before, QuietValues[i].GetType(), Inv), QuietValues[i]))
            {
                try { Invoke(app, QuietNames[i], BindingFlags.SetProperty, new object[] { QuietValues[i] }); }
                catch (Exception)
                {
                    if (QuietNames[i] == "Calculation") continue;    // no workbook open: left alone
                    throw;
                }
                changed.Add(new KeyValuePair<string, object>(QuietNames[i], before));
            }
        }
    }

    // Each changed setting put back, the last first; "name=value: message" of each that could not be.
    static void Restore(object app, List<KeyValuePair<string, object>> changed, List<object> unrestored)
    {
        for (int i = changed.Count - 1; i >= 0; i--)
        {
            try { Invoke(app, changed[i].Key, BindingFlags.SetProperty, new object[] { changed[i].Value }); }
            catch (Exception e)
            {
                unrestored.Add(changed[i].Key + "=" + Convert.ToString(changed[i].Value, Inv) + ": " + Message(e));
            }
        }
    }

    // ------------------------------------------------------------------ operations
    static void Do(Dictionary<string, object> op, Dictionary<int, object> refs,
                   List<KeyValuePair<string, object>> values, List<KeyValuePair<string, object>> marks, Stopwatch sw)
    {
        string k = (string)op["k"];
        string name = (string)op["name"];
        if (k == "mark") { marks.Add(new KeyValuePair<string, object>(name, sw.Elapsed.TotalMilliseconds)); return; }
        if (k == "ret") { values.Add(new KeyValuePair<string, object>(name, Plain(Ref(refs, op["on"])))); return; }
        if (k == "set")
        {
            Invoke(Ref(refs, op["on"]), name, BindingFlags.SetProperty, new object[] { Arg(op["value"], refs) });
            return;
        }
        object r;
        if (k == "get") r = Invoke(Ref(refs, op["on"]), name, Get, new object[0]);
        else if (k == "call") r = Invoke(Ref(refs, op["on"]), name, Get, Args((ArrayList)op["args"], refs));
        else if (k == "verb") r = Verb(name, Args((ArrayList)op["args"], refs));
        else throw new Exception("unknown operation '" + k + "'");
        refs[Convert.ToInt32(op["as"], Inv)] = r;
    }

    // ------------------------------------------------------------------ IDispatch with LCID 0
    // Every call goes through IDispatch::Invoke with LCID 0, the call pywin32 makes: Excel then parses text as
    // en-US but reads a NumberFormat in the user's codes, and no .NET culture (InvokeMember's way) gives that
    // (en-US differs on number formats, the user's culture on parsing, the invariant one is refused).
    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatchRaw
    {
        [PreserveSig] int GetTypeInfoCount(out uint count);
        [PreserveSig] int GetTypeInfo(uint index, int lcid, out IntPtr info);
        [PreserveSig] int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] names,
                                        int count, int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] ids);
        [PreserveSig] int Invoke(int id, ref Guid riid, int lcid, short flags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS args,
                                 IntPtr result, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO info, IntPtr argErr);
    }

    [DllImport("oleaut32.dll")] static extern int VariantClear(IntPtr v);

    const int Lcid = 0;
    const short MethodOrGet = 3, Put = 4;                // DISPATCH_METHOD | DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT
    const int DispidPropertyPut = -3, DispException = unchecked((int)0x80020009);
    static readonly int VariantSize = IntPtr.Size == 8 ? 24 : 16;

    class ExcelError : Exception
    {
        public readonly int Code;
        public ExcelError(string message, int code) : base(message) { Code = code; }
    }

    static object Invoke(object target, string name, BindingFlags flags, object[] args)
    {
        var d = target as IDispatchRaw;
        if (d == null) throw new Exception(name + ": the target is not a COM object");
        var none = Guid.Empty;
        var ids = new int[1];
        int hr = d.GetIDsOfNames(ref none, new string[] { name }, 1, Lcid, ids);
        if (hr < 0) throw new ExcelError("Unknown name.", hr);              // as pywin32 says it
        bool put = flags == BindingFlags.SetProperty;
        int n = args.Length;
        IntPtr argv = Marshal.AllocCoTaskMem(VariantSize * Math.Max(n, 1));
        IntPtr result = Marshal.AllocCoTaskMem(VariantSize);
        IntPtr named = Marshal.AllocCoTaskMem(4);
        int made = 0;
        try
        {
            for (int b = 0; b < VariantSize; b++) Marshal.WriteByte(result, b, 0);
            for (; made < n; made++)                           // IDispatch takes the arguments last first
                Marshal.GetNativeVariantForObject(args[n - 1 - made], argv + made * VariantSize);
            var dp = new System.Runtime.InteropServices.ComTypes.DISPPARAMS();
            dp.rgvarg = argv;
            dp.cArgs = n;
            if (put)
            {
                Marshal.WriteInt32(named, DispidPropertyPut);
                dp.rgdispidNamedArgs = named;
                dp.cNamedArgs = 1;
            }
            var info = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
            hr = d.Invoke(ids[0], ref none, Lcid, put ? Put : MethodOrGet, ref dp, put ? IntPtr.Zero : result, ref info,
                          IntPtr.Zero);
            if (hr == DispException)
            {
                int code = info.scode != 0 ? info.scode : info.wCode;
                throw new ExcelError(info.bstrDescription ?? "Exception occurred.", code != 0 ? code : hr);
            }
            if (hr < 0) throw new ExcelError(Marshal.GetExceptionForHR(hr).Message, hr);
            return put ? null : Marshal.GetObjectForNativeVariant(result);
        }
        finally
        {
            for (int i = 0; i < made; i++) VariantClear(argv + i * VariantSize);
            VariantClear(result);
            Marshal.FreeCoTaskMem(argv);
            Marshal.FreeCoTaskMem(result);
            Marshal.FreeCoTaskMem(named);
        }
    }

    static object Ref(Dictionary<int, object> refs, object id)
    {
        int i = Convert.ToInt32(id, Inv);
        object o;
        if (!refs.TryGetValue(i, out o)) throw new Exception(string.Format(Inv, Unbound, i));
        return o;
    }

    static object[] Args(ArrayList a, Dictionary<int, object> refs)
    {
        var o = new object[a.Count];
        for (int i = 0; i < a.Count; i++) o[i] = Arg(a[i], refs);
        return o;
    }

    // A value of the batch as pywin32 sends it: None is VT_NULL (DBNull), MISSING an omitted argument, a float its
    // 64 bits in hex, a list a 1-D array, a grid a 2-D array.
    static object Arg(object v, Dictionary<int, object> refs)
    {
        if (v == null) return DBNull.Value;
        var d = v as Dictionary<string, object>;
        if (d == null) return v is decimal ? (object)Convert.ToDouble(v, Inv) : v;
        if (d.ContainsKey("r")) return Ref(refs, d["r"]);
        if (d.ContainsKey("d")) return BitConverter.Int64BitsToDouble(Convert.ToInt64((string)d["d"], 16));
        if (d.ContainsKey("missing")) return Type.Missing;
        if (d.ContainsKey("list"))
        {
            var l = (ArrayList)d["list"];
            var arr = new object[l.Count];
            for (int i = 0; i < l.Count; i++) arr[i] = Arg(l[i], refs);
            return arr;
        }
        if (d.ContainsKey("grid"))
        {
            var rows = (ArrayList)d["grid"];
            int h = rows.Count, w = ((ArrayList)rows[0]).Count;
            var g = new object[h, w];
            for (int i = 0; i < h; i++)
            {
                var row = (ArrayList)rows[i];
                for (int j = 0; j < w; j++) g[i, j] = Arg(row[j], refs);
            }
            return g;
        }
        throw new Exception("unknown value");
    }

    // ------------------------------------------------------------------ verbs
    static object Verb(string name, object[] args)
    {
        if (name == "break_links") return BreakLinks(args[0], (string)args[1]);
        throw new Exception("unknown verb '" + name + "'");
    }

    // Break the workbook's links whose file name is fileName (any case); its other links stay. LinkSources is read
    // first: BreakLink needs the exact link text Excel holds.
    static object BreakLinks(object wb, string fileName)
    {
        var links = Invoke(wb, "LinkSources", Get, new object[] { Links }) as Array;
        int broke = 0;
        if (links == null) return broke;
        foreach (object l in links)
        {
            string s = Convert.ToString(l, Inv);
            int cut = Math.Max(s.LastIndexOf('\\'), s.LastIndexOf('/'));
            if (string.Compare(s.Substring(cut + 1), fileName, StringComparison.OrdinalIgnoreCase) == 0)
            {
                Invoke(wb, "BreakLink", Get, new object[] { s, Links });
                broke++;
            }
        }
        return broke;
    }

    // ------------------------------------------------------------------ values back
    // As pywin32 gives them: arrays as lists of rows (1-D as one list), empty as null, an error cell as its integer,
    // a currency as a double; a COM object or a date is an error.
    static object Plain(object v)
    {
        if (v == null || v is DBNull) return null;
        if (v is string || v is bool || v is double || v is int || v is long || v is short || v is byte) return v;
        if (v is float || v is decimal) return Convert.ToDouble(v, Inv);
        if (v is DateTime) throw new Exception(NotADate);
        var a = v as Array;
        if (a != null)
        {
            var rows = new List<object>();
            if (a.Rank == 1)
            {
                foreach (object c in a) rows.Add(Plain(c));
                return rows;
            }
            if (a.Rank != 2) throw new Exception("an array of rank " + a.Rank + " cannot be asked back");
            int i0 = a.GetLowerBound(0), j0 = a.GetLowerBound(1);
            for (int i = 0; i < a.GetLength(0); i++)
            {
                var row = new List<object>();
                for (int j = 0; j < a.GetLength(1); j++) row.Add(Plain(a.GetValue(i0 + i, j0 + j)));
                rows.Add(row);
            }
            return rows;
        }
        if (Marshal.IsComObject(v)) throw new Exception(NotAValue);
        throw new Exception("a " + v.GetType().Name + " cannot be asked back");
    }

    static string Message(Exception e)
    {
        while ((e is TargetInvocationException) && e.InnerException != null) e = e.InnerException;
        var x = e as ExcelError;
        if (x != null) return x.Message + " (0x" + x.Code.ToString("X8", Inv) + ")";
        var c = e as COMException;
        if (c != null) return c.Message + " (0x" + c.ErrorCode.ToString("X8", Inv) + ")";
        return e.Message;
    }

    // ------------------------------------------------------------------ JSON out (ASCII; floats exact)
    static void Pairs(StringBuilder sb, List<KeyValuePair<string, object>> pairs)
    {
        sb.Append('{');
        for (int i = 0; i < pairs.Count; i++)
        {
            if (i > 0) sb.Append(',');
            Str(sb, pairs[i].Key);
            sb.Append(':');
            Write(sb, pairs[i].Value);
        }
        sb.Append('}');
    }

    static void Write(StringBuilder sb, object v)
    {
        if (v == null) { sb.Append("null"); return; }
        if (v is bool) { sb.Append((bool)v ? "true" : "false"); return; }
        if (v is string) { Str(sb, (string)v); return; }
        if (v is int || v is long || v is short || v is byte) { sb.Append(Convert.ToInt64(v, Inv).ToString(Inv)); return; }
        if (v is double)
        {
            double d = (double)v;
            if (double.IsNaN(d) || double.IsInfinity(d)) { sb.Append("null"); return; }
            string s = d.ToString("G17", Inv);
            if (d == 0 && 1 / d < 0) s = "-0";                  // G17 drops the sign of a negative zero
            sb.Append(s);
            if (s.IndexOf('.') < 0 && s.IndexOf('E') < 0) sb.Append(".0");
            return;
        }
        var list = v as IList;
        if (list != null)
        {
            sb.Append('[');
            for (int i = 0; i < list.Count; i++)
            {
                if (i > 0) sb.Append(',');
                Write(sb, list[i]);
            }
            sb.Append(']');
            return;
        }
        Str(sb, Convert.ToString(v, Inv));
    }

    static void Str(StringBuilder sb, string s)
    {
        sb.Append('"');
        foreach (char ch in s)
        {
            if (ch == '"') sb.Append("\\\"");
            else if (ch == '\\') sb.Append("\\\\");
            else if (ch < 0x20 || ch > 0x7e) sb.Append("\\u").Append(((int)ch).ToString("x4", Inv));
            else sb.Append(ch);
        }
        sb.Append('"');
    }
}
