// The in-process executor of smartass batches (engine/smartass/xl/inproc): a template. build.py puts the build's hash
// in place of __HASH__, so every build has its own function names (a loaded add-in's names stay taken in that Excel),
// and the jobs folder in place of __JOBS__. SmartassRun is a function, not a command, because its result is its
// return value (a command returns nothing); it refuses a cell caller and any path but a job file of that folder.
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
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using ExcelDna.Integration;

public static class SmartassRunner___HASH__
{
    const string Hash = "__HASH__";
    const int Version = 1;
    const int InlineMax = 30000;                               // below Excel's 32 767 characters per string
    const string JobsFolder = @"__JOBS__";                     // build.py: the one folder job files are run from
    const string NotRun = "not run: ";
    static readonly Regex JobName = new Regex(@"^[0-9a-f]{32}\.json\z");
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

    // Runs the job at jobPath; never throws. Nothing is read, deleted or run unless the call comes from outside a cell
    // (Application.Run) and jobPath is a job file of the one jobs folder compiled in: JobsFolder\<32 hex>.json, given
    // as its full, plain path. Otherwise the return value starts with "not run:" (or, in a cell, is #N/A) and nothing
    // is touched. The job file is then claimed: read and deleted (a job whose file cannot be deleted is not run: the
    // caller could take it for one that never started and run it again). The result record is the return value when
    // it fits an Excel string (InlineMax); otherwise it is written to jobPath + ".out" (UTF-8, no BOM) and the return
    // value is "file". (Opening a file Excel has just written costs the caller about 20 ms: the virus scanner reads
    // it first.)
    [ExcelFunction(Name = "SmartassRun___HASH__", IsMacroType = true, IsHidden = true)]
    public static object Run(string jobPath)
    {
        if (FromCell()) return ExcelError.ExcelErrorNA;
        string refused = Refusal(jobPath);
        if (refused != null) return NotRun + "refused: " + refused;
        string text;
        try
        {
            text = File.ReadAllText(jobPath, Encoding.UTF8);
        }
        catch (Exception e)
        {
            return NotRun + "the job could not be read: " + e.Message;
        }
        try
        {
            File.Delete(jobPath);
        }
        catch (Exception e)
        {
            return NotRun + "the job file could not be deleted: " + e.Message;
        }
        if (File.Exists(jobPath)) return NotRun + "the job file could not be deleted";
        return Execute(jobPath, text);
    }

    static bool FromCell()
    {
        try
        {
            return XlCall.Excel(XlCall.xlfCaller) is ExcelReference;
        }
        catch (Exception)
        {
            return true;                                          // the caller cannot be told: refused
        }
    }

    static string Refusal(string jobPath)
    {
        if (string.IsNullOrEmpty(jobPath)) return "no job path";
        if (jobPath.StartsWith(@"\\") || jobPath.IndexOf('/') >= 0) return "not a local plain path";
        if (jobPath.Length < 3 || jobPath[1] != ':' || jobPath.IndexOf(':', 2) >= 0) return "not a plain file path";
        string full;
        try { full = Path.GetFullPath(jobPath); }
        catch (Exception) { return "not a valid path"; }
        if (!string.Equals(full, jobPath, StringComparison.OrdinalIgnoreCase)) return "not a full plain path";
        if (!string.Equals(Path.GetDirectoryName(full), JobsFolder, StringComparison.OrdinalIgnoreCase))
            return "not in the jobs folder";
        if (!JobName.IsMatch(Path.GetFileName(full))) return "not a job file name";
        return null;
    }

    static object Execute(string jobPath, string text)
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
                    started = i + 1;                              // before anything about the operation can fail
                    var op = ops[i] as Dictionary<string, object>;
                    try { Do(op, refs, values, marks, sw); }
                    catch (Exception e)
                    {
                        if (op != null && op.ContainsKey("try") && Equals(op["try"], true))
                        {
                            skipped.Add(new object[] { i, Failure(i, op, e) });
                            continue;
                        }
                        ok = false; index = i; message = Failure(i, op, e);
                        break;
                    }
                }
            }
            finally
            {
                for (int j = 0; j < fin.Count; j++)
                {
                    var op = fin[j] as Dictionary<string, object>;
                    try { Do(op, refs, values, marks, sw); }
                    catch (Exception e) { skipped.Add(new object[] { ops.Count + j, Failure(ops.Count + j, op, e) }); }
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
        catch (Exception e)
        {
            return "no result: the result file could not be written: " + e.Message;
        }
        return "file";
    }

    // "operation <index> (<kind> <name>): <message>": how a failing operation's text starts in both executors.
    static string Failure(int index, Dictionary<string, object> op, Exception e)
    {
        string kind = op != null && op.ContainsKey("k") ? Convert.ToString(op["k"], Inv) : "?";
        string name = op != null && op.ContainsKey("name") ? Convert.ToString(op["name"], Inv) : "?";
        return "operation " + index.ToString(Inv) + " (" + kind + " " + name + "): " + Message(e);
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
        if (op == null) throw new Exception("not an operation");
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

    class CallError : Exception
    {
        public readonly int Code;
        public CallError(string message, int code) : base(message) { Code = code; }
    }

    static object Invoke(object target, string name, BindingFlags flags, object[] args)
    {
        var d = target as IDispatchRaw;
        if (d == null) throw new Exception(name + ": the target is not a COM object");
        var none = Guid.Empty;
        var ids = new int[1];
        int hr = d.GetIDsOfNames(ref none, new string[] { name }, 1, Lcid, ids);
        if (hr < 0) throw new CallError("Unknown name.", hr);              // as pywin32 says it
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
                throw new CallError(info.bstrDescription ?? "Exception occurred.", code != 0 ? code : hr);
            }
            if (hr < 0) throw new CallError(Marshal.GetExceptionForHR(hr).Message, hr);
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
    // Named steps an operation list cannot say (batch.VERBS, runner.VERBS; batch.py's module docstring says what
    // each does): each, the one loop a batch has.
    static object Verb(string name, object[] args)
    {
        if (name == "each" && args.Length == 3) return Each(args[0], args[1] as object[], args[2]);
        throw new Exception("unknown verb '" + name + "'");
    }

    // each(collection, paths, prefixes): for every item (Count, then Item(1..Count)), the value at each dotted path
    // ("*": the rest of the path over every item of the collection reached); null where it cannot be read. prefixes
    // (null or a list): only the items whose first value starts with one of them, case ignored.
    static object Each(object collection, object[] paths, object prefixes)
    {
        if (paths == null) throw new Exception("each: the paths are a list of names");
        var keep = prefixes as object[];
        if (prefixes != null && !(prefixes is DBNull) && keep == null) throw new Exception("each: the prefixes are a list");
        var split = new string[paths.Length][];
        for (int j = 0; j < paths.Length; j++) split[j] = Convert.ToString(paths[j], Inv).Split('.');
        int n = Convert.ToInt32(Invoke(collection, "Count", Get, new object[0]), Inv);
        var rows = new List<object>();
        for (int i = 1; i <= n; i++)
        {
            object item = Invoke(collection, "Item", Get, new object[] { i });
            var row = new object[paths.Length];
            for (int j = 0; j < paths.Length; j++) row[j] = PathValue(item, split[j], 0);
            if (keep == null || Starts(row.Length > 0 ? row[0] as string : null, keep)) rows.Add(row);
        }
        return rows.ToArray();
    }

    static bool Starts(string s, object[] prefixes)
    {
        if (s == null) return false;
        foreach (object p in prefixes)
            if (s.StartsWith(Convert.ToString(p, Inv), StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    static object PathValue(object obj, string[] parts, int k)
    {
        try
        {
            for (; k < parts.Length; k++)
            {
                if (parts[k] == "*")
                {
                    int n = Convert.ToInt32(Invoke(obj, "Count", Get, new object[0]), Inv);
                    var items = new object[n];
                    for (int i = 1; i <= n; i++)
                        items[i - 1] = PathValue(Invoke(obj, "Item", Get, new object[] { i }), parts, k + 1);
                    return items;
                }
                obj = Invoke(obj, parts[k], Get, new object[0]);
            }
            Plain(obj);                                           // a COM object or a date is not a value
            return obj;
        }
        catch (Exception)
        {
            return null;
        }
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
        var x = e as CallError;
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
