// The in-process executor of smartass batches (engine/smartass/xl/inproc): a template. build.py puts the build's hash
// in place of __HASH__, so every build has its own function names (a loaded add-in's names stay taken in that Excel),
// and the jobs folder in place of __JOBS__. SmartassRun is a function, not a command, because its result is its
// return value (a command returns nothing); it refuses a cell caller and any path but a job file of that folder.
// C# 5 (the .NET Framework compiler that ships with Windows); the batch format is described in batch.py, the result
// record in runner.py, and the COM executor there does the same over COM from Python. SmartassSaveLater saves our own
// analysis file after the reply (xl/inproc/later.py says how and why).
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
using CT = System.Runtime.InteropServices.ComTypes;
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
    const string Taken = "taken: ";                            // the verb exist's refusal, as runner.TAKEN
    static readonly string[] QuietNames = { "ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation" };
    static readonly object[] QuietValues = { false, false, false, -4135 };
    const BindingFlags Get = BindingFlags.GetProperty | BindingFlags.InvokeMethod;
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly Regex Token = new Regex(@"^[0-9a-f]{32}\z");
    const int DelayMaxMs = 10000;
    static Pending pending;                                    // the save asked for and not run yet (Excel's thread)
    static bool saving;                                        // that save is running: Excel answers calls meanwhile
    static string unwritten;                                   // the outcome of a save that could not be written
    const string Busy = "busy: the save after the reply is running";
    const string Unwritten = "unwritten: ";
    static System.Threading.Timer timer;                       // rooted here: a timer nothing holds can be collected

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
        if (saving) return Busy;                                 // called in the middle of the save: asked again
        SaveNow("before a batch");                               // a save still waiting runs first: never mid-job
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

    // ------------------------------------------------------------------ the save after the reply (xl/inproc/later.py)
    class Pending
    {
        public string Workbook, FullName, Token;
        public bool AutoSave;
        public double Queued;
    }

    // A job's first contact with this Excel (later.clear, before any plain COM call writes in it): a save after the
    // reply still waiting runs now; while one runs the answer is Busy (the caller asks again); then "clear", or the
    // outcome of a save that could not be written, once, as Unwritten + its record. Never throws.
    [ExcelFunction(Name = "SmartassClear___HASH__", IsMacroType = true, IsHidden = true)]
    public static object Clear()
    {
        if (FromCell()) return ExcelError.ExcelErrorNA;
        if (saving) return Busy;
        try { SaveNow("before a job"); }
        catch (Exception) { }
        string u = unwritten;
        unwritten = null;
        return u == null ? "clear" : Unwritten + u;
    }

    // Asks for our own workbook to be saved once Excel has painted the answer (the workbook open by that name; its
    // FullName is kept, and a workbook of that name with another FullName at save time is not saved: "closed"). The
    // save runs delayMs (at most DelayMaxMs) from now, as a macro when Excel is ready (ExcelAsyncUtil.QueueAsMacro),
    // or at the start of the next batch's Run if that comes first, or of the next SaveLater; then AutoSave goes back
    // on when autosave is True (also after a failed save) and the outcome is written to saved\<this Excel's pid>.json
    // beside the jobs folder. A save still waiting runs first. Excel answers calls in the middle of a save (a OneDrive
    // file, live 20.11.2026): while it runs, Run and SaveLater answer Busy having done nothing, and the caller asks
    // again (runner.asked). Returns "queued", Busy, or "not run: ..." for a refused request; never throws.
    [ExcelFunction(Name = "SmartassSaveLater___HASH__", IsMacroType = true, IsHidden = true)]
    public static object SaveLater(string workbook, bool autosave, string token, double delayMs)
    {
        if (FromCell()) return ExcelError.ExcelErrorNA;
        if (saving) return Busy;
        if (string.IsNullOrEmpty(workbook)) return NotRun + "refused: no workbook";
        if (token == null || !Token.IsMatch(token)) return NotRun + "refused: not a token";
        try
        {
            SaveNow("before the next");
            object books = Invoke(ExcelDnaUtil.Application, "Workbooks", Get, new object[0]);
            object wb;
            try { wb = Invoke(books, "Item", Get, new object[] { workbook }); }
            catch (Exception) { return NotRun + "refused: no workbook of that name is open"; }
            var p = new Pending();
            p.Workbook = workbook; p.Token = token; p.AutoSave = autosave; p.Queued = Now();
            p.FullName = Convert.ToString(Invoke(wb, "FullName", BindingFlags.GetProperty, new object[0]), Inv);
            pending = p;
            int delay = (int)Math.Max(0, Math.Min(double.IsNaN(delayMs) ? 0 : delayMs, DelayMaxMs));
            var old = timer;
            timer = new System.Threading.Timer(Fire, p, delay, System.Threading.Timeout.Infinite);
            if (old != null) old.Dispose();
            return "queued";
        }
        catch (Exception e)
        {
            pending = null;
            return NotRun + e.Message;
        }
    }

    // The timer (a pool thread): the save queued as a macro; it runs only if it is still the one waiting.
    static void Fire(object state)
    {
        var p = (Pending)state;
        try
        {
            ExcelAsyncUtil.QueueAsMacro(delegate
            {
                try { if (pending == p) SaveNow("after the reply"); }
                catch (Exception) { }
            });
        }
        catch (Exception) { }                                     // Excel is closing: it asks about the workbook
    }

    // The save waiting, if any, run now on Excel's thread; never throws. how: where it ran (for the outcome).
    static void SaveNow(string how)
    {
        var p = pending;
        if (p == null || saving) return;
        pending = null;
        saving = true;
        var t = timer;                                           // its work is done (or done now): not left behind
        timer = null;
        if (t != null) t.Dispose();
        double started = Now(), seconds = 0, switched = 0;
        var sw = Stopwatch.StartNew();
        string error = null, autosaveError = null;
        bool closed = false;
        try
        {
            object books = Invoke(ExcelDnaUtil.Application, "Workbooks", Get, new object[0]);
            object wb = null;
            try { wb = Invoke(books, "Item", Get, new object[] { p.Workbook }); }
            catch (Exception e) { closed = true; error = Message(e); }
            if (wb != null && !string.Equals(Convert.ToString(Invoke(wb, "FullName", BindingFlags.GetProperty,
                                                                     new object[0]), Inv), p.FullName,
                                              StringComparison.OrdinalIgnoreCase))
            {
                wb = null; closed = true; error = "another workbook of that name is open now";
            }
            if (wb != null)
            {
                try { Invoke(wb, "Save", Get, new object[0]); }
                catch (Exception e) { error = Message(e); }
                seconds = sw.Elapsed.TotalSeconds;
                if (p.AutoSave)
                {
                    try { Invoke(wb, "AutoSaveOn", BindingFlags.SetProperty, new object[] { true }); }
                    catch (Exception e) { autosaveError = Message(e); }
                    switched = sw.Elapsed.TotalSeconds - seconds;
                }
            }
        }
        catch (Exception e)
        {
            if (error == null) error = Message(e);
        }
        finally
        {
            saving = false;
        }
        if (seconds == 0) seconds = sw.Elapsed.TotalSeconds;
        string record = Record(p, how, error, autosaveError, closed, started, seconds, switched);
        try { WriteOutcome(p, record); }
        catch (Exception e)                                       // told by the next Clear (the next job's first call)
        {
            var sb = new StringBuilder(record, 0, record.Length - 1, record.Length + 64);
            sb.Append(",\"unwritten\":"); Str(sb, Message(e)); sb.Append("}");
            unwritten = sb.ToString();
        }
    }

    static double Now()
    {
        return (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
    }

    // The outcome, no figure in it, as JSON.
    static string Record(Pending p, string how, string error, string autosaveError, bool closed, double started,
                         double seconds, double switched)
    {
        var sb = new StringBuilder();
        sb.Append("{\"token\":"); Str(sb, p.Token);
        sb.Append(",\"workbook\":"); Str(sb, p.Workbook);
        sb.Append(",\"full_name\":"); Str(sb, p.FullName ?? "");
        sb.Append(",\"ok\":").Append(error == null ? "true" : "false");
        sb.Append(",\"error\":"); Write(sb, error);
        sb.Append(",\"closed\":").Append(closed ? "true" : "false");
        sb.Append(",\"autosave\":").Append(p.AutoSave ? "true" : "false");
        sb.Append(",\"autosave_error\":"); Write(sb, autosaveError);
        sb.Append(",\"how\":"); Str(sb, how);
        sb.Append(",\"queued\":"); Write(sb, p.Queued);
        sb.Append(",\"started\":"); Write(sb, started);
        sb.Append(",\"seconds\":"); Write(sb, seconds);
        sb.Append(",\"switched\":"); Write(sb, switched);
        sb.Append("}");
        return sb.ToString();
    }

    // The outcome written to saved\<pid>.json beside the jobs folder, whole (a temporary file renamed).
    static void WriteOutcome(Pending p, string record)
    {
        string folder = Path.Combine(Path.GetDirectoryName(JobsFolder), "saved");
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, Process.GetCurrentProcess().Id.ToString(Inv) + ".json");
        string tmp = path + "." + p.Token + ".tmp";
        File.WriteAllText(tmp, record, new UTF8Encoding(false));
        try
        {
            if (File.Exists(path)) File.Replace(tmp, path, null);
            else File.Move(tmp, path);
        }
        catch (Exception)
        {
            try { if (File.Exists(path)) File.Delete(path); File.Move(tmp, path); }
            catch (Exception) { try { File.Delete(tmp); } catch (Exception) { } throw; }
        }
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
                bool calc = !job.ContainsKey("calculation") || !Equals(job["calculation"], false);
            if (job.ContainsKey("quiet") && (bool)job["quiet"]) Quiet(app, changed, calc);
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
    static void Quiet(object app, List<KeyValuePair<string, object>> changed, bool calc)
    {
        for (int i = 0; i < QuietNames.Length; i++)
        {
            if (!calc && QuietNames[i] == "Calculation") continue;  // a batch without formulas leaves it alone
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
    const short MethodOrGet = 3, PropGet = 2, Put = 4;   // DISPATCH_METHOD | DISPATCH_PROPERTYGET, PROPERTYGET, PUT
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
            short kind = put ? Put : flags == BindingFlags.GetProperty ? PropGet : MethodOrGet;
            hr = d.Invoke(ids[0], ref none, Lcid, kind, ref dp, put ? IntPtr.Zero : result, ref info, IntPtr.Zero);
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
    // each does): each, the one loop a batch has; nudge; exist; mentions; swap.
    static object Verb(string name, object[] args)
    {
        if (name == "each" && args.Length == 2) return Each(args[0], args[1] as object[]);
        if (name == "nudge" && args.Length == 4) return Nudge(args[0], Convert.ToString(args[1], Inv), args[2], args[3]);
        if (name == "exist" && args.Length == 3) return Exist(args[0], args[1] as object[], args[2]);
        if (name == "mentions" && args.Length == 3) return Mentions(args[0], args[1] as object[], args[2] as object[]);
        if (name == "swap" && args.Length == 4) return Swap(args[0], Convert.ToString(args[1], Inv), args[2], args[3]);
        throw new Exception("unknown verb '" + name + "'");
    }

    // mentions(workbook, texts, skip): every formula of the workbook outside the skipped sheets that mentions one of
    // the texts (case ignored), as rows [kind, where, ref, formula] (batch.py's module docstring; runner._mentions):
    // per worksheet its cells (read in chunks, only those that hold a formula), its charts' series; then
    // the chart sheets' series and the defined names. In-process nothing is busy: an item that cannot be read is
    // left out, a sheet whose cells cannot be read is an "unread" row.

    static object Mentions(object workbook, object[] texts, object[] skip)
    {
        if (texts == null || skip == null) throw new Exception("mentions: the texts and the sheets to skip are lists");
        var want = new string[texts.Length];
        for (int i = 0; i < texts.Length; i++) want[i] = Convert.ToString(texts[i], Inv);
        var skipped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (object s in skip) skipped.Add(Convert.ToString(s, Inv));
        var rows = new List<object>();
        object sheets = Invoke(workbook, "Worksheets", Get, new object[0]);
        int n = Count(sheets);
        for (int i = 1; i <= n; i++)
        {
            object ws = Invoke(sheets, "Item", Get, new object[] { i });
            string name = Convert.ToString(Invoke(ws, "Name", Get, new object[0]), Inv);
            if (skipped.Contains(name)) continue;
            try { MentionsInCells(ws, name, want, rows); }
            catch (Exception e) { rows.Add(new object[] { "unread", name, "", Message(e) }); }
            object charts;
            int m;
            try
            {
                charts = Invoke(ws, "ChartObjects", Get, new object[0]);
                m = Count(charts);
            }
            catch (Exception) { continue; }
            for (int c = 1; c <= m; c++)
            {
                object chart;
                try { chart = Invoke(Invoke(charts, "Item", Get, new object[] { c }), "Chart", Get, new object[0]); }
                catch (Exception) { continue; }
                MentionsInSeries(chart, "series", name, c.ToString(Inv) + "/", want, rows);
            }
        }
        object chartSheets = Invoke(workbook, "Charts", Get, new object[0]);
        int k = Count(chartSheets);
        for (int i = 1; i <= k; i++)
        {
            object chart;
            string name;
            try
            {
                chart = Invoke(chartSheets, "Item", Get, new object[] { i });
                name = Convert.ToString(Invoke(chart, "Name", Get, new object[0]), Inv);
            }
            catch (Exception) { continue; }
            MentionsInSeries(chart, "chart", name, "", want, rows);
        }
        object names = Invoke(workbook, "Names", Get, new object[0]);
        int count = Count(names);
        for (int j = 1; j <= count; j++)
        {
            try
            {
                object item = Invoke(names, "Item", Get, new object[] { j });
                object formula = Invoke(item, "RefersToLocal", BindingFlags.GetProperty, new object[0]);
                if (!Says(formula as string, want)) continue;     // the name's own Name only for one that mentions
                object nm = Invoke(item, "Name", BindingFlags.GetProperty, new object[0]);
                rows.Add(new object[] { "name", Convert.ToString(nm, Inv), j.ToString(Inv), formula });
            }
            catch (Exception) { }
        }
        return rows.ToArray();
    }

    static int Count(object collection)
    {
        return Convert.ToInt32(Invoke(collection, "Count", Get, new object[0]), Inv);
    }

    static bool Says(string text, string[] want)
    {
        if (text == null) return false;
        foreach (string t in want)
            if (text.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        return false;
    }

    static string Column(int c)
    {
        string s = "";
        for (; c > 0; c = (c - 1) / 26) s = (char)('A' + (c - 1) % 26) + s;
        return s;
    }

    static string Area(int row, int col, int rows, int cols)
    {
        string first = Column(col) + row.ToString(Inv);
        if (rows == 1 && cols == 1) return first;
        return first + ":" + Column(col + cols - 1) + (row + rows - 1).ToString(Inv);
    }

    static int Int(object obj, string name)
    {
        return Convert.ToInt32(Invoke(obj, name, Get, new object[0]), Inv);
    }

    const int ChunkCells = 100000;                              // as runner.CHUNK_CELLS
    const int XlCellTypeFormulas = -4123;
    const int XlCellTypeLastCell = 11;

    // mentions on one worksheet's cells (runner._cells): the used range's Formula2 read in chunks of rows of at most
    // ChunkCells cells, only the chunks Intersect finds a formula in (SpecialCells once; none, while SpecialCells
    // answers on the sheet: done; a refusal: "unread"); a protected sheet refuses SpecialCells, so each of its chunks
    // is read (a hidden formula makes the read throw: "unread").
    static void MentionsInCells(object ws, string name, string[] want, List<object> rows)
    {
        object used = Invoke(ws, "UsedRange", Get, new object[0]);
        int r0 = Int(used, "Row"), c0 = Int(used, "Column");
        int nr = Int(Invoke(used, "Rows", Get, new object[0]), "Count");
        int nc = Int(Invoke(used, "Columns", Get, new object[0]), "Count");
        object formulas = null;
        if (!Equals(Invoke(ws, "ProtectContents", Get, new object[0]), true))
        {
            try { formulas = Invoke(used, "SpecialCells", Get, new object[] { XlCellTypeFormulas }); }
            catch (Exception)
            {
                // "No cells were found" has the code of any other refusal and a text in Excel's language: no formula
                // only when SpecialCells answers on the sheet at all (its last cell); a refusal is the caller's "unread"
                bool answers = true;
                try { Invoke(used, "SpecialCells", Get, new object[] { XlCellTypeLastCell }); }
                catch (Exception) { answers = false; }
                if (!answers) throw;
                return;
            }
        }
        object app = Invoke(ws, "Application", Get, new object[0]);
        int step = Math.Max(1, ChunkCells / nc);
        var arrays = new HashSet<string>();
        for (int top = r0; top < r0 + nr; top += step)
        {
            int bottom = Math.Min(top + step, r0 + nr) - 1;
            object chunk = Invoke(ws, "Range", Get, new object[] {
                Invoke(ws, "Cells", Get, new object[] { top, c0 }),
                Invoke(ws, "Cells", Get, new object[] { bottom, c0 + nc - 1 }) });
            if (formulas != null && Invoke(app, "Intersect", Get, new object[] { formulas, chunk }) == null) continue;
            object grid = Invoke(chunk, "Formula2", Get, new object[0]);
            var a = grid as Array;
            int h = a == null ? 1 : a.GetLength(0), w = a == null ? 1 : a.GetLength(1);
            for (int i = 0; i < h; i++)
            {
                for (int j = 0; j < w; j++)
                {
                    object v = a == null ? grid : a.GetValue(a.GetLowerBound(0) + i, a.GetLowerBound(1) + j);
                    var text = v as string;
                    if (!Says(text, want)) continue;
                    object cell = Invoke(ws, "Cells", Get, new object[] { top + i, c0 + j });
                    if (!Equals(Invoke(cell, "HasFormula", Get, new object[0]), true)) continue;   // text, no formula
                    if (Equals(Invoke(cell, "HasArray", Get, new object[0]), true))
                    {
                        object arr = Invoke(cell, "CurrentArray", Get, new object[0]);
                        string reference = Area(Int(arr, "Row"), Int(arr, "Column"),
                                                Int(Invoke(arr, "Rows", Get, new object[0]), "Count"),
                                                Int(Invoke(arr, "Columns", Get, new object[0]), "Count"));
                        if (arrays.Add(reference))
                            rows.Add(new object[] { "array", name, reference,
                                                    Invoke(arr, "FormulaArray", Get, new object[0]) });
                    }
                    else rows.Add(new object[] { "cell", name, Area(top + i, c0 + j, 1, 1), text });
                }
            }
        }
    }

    // swap(object, property, expected, value): the text property set to value only when it reads expected exactly;
    // true when it was set, false when it read anything else (left as it is).
    static object Swap(object obj, string name, object expected, object value)
    {
        if (!IsProperty(obj, name)) throw new Exception(name + " is not a property");
        object now = Invoke(obj, name, BindingFlags.GetProperty, new object[0]);
        if (expected == null || !string.Equals(now as string, expected as string, StringComparison.Ordinal))
            return false;
        Invoke(obj, name, BindingFlags.SetProperty, new object[] { value });
        return true;
    }

    static void MentionsInSeries(object chart, string kind, string where, string prefix, string[] want,
                                 List<object> rows)
    {
        object series;
        int n;
        try
        {
            series = Invoke(chart, "SeriesCollection", Get, new object[0]);
            n = Count(series);
        }
        catch (Exception) { return; }
        for (int j = 1; j <= n; j++)
        {
            object formula;
            try
            {
                object item = Invoke(series, "Item", Get, new object[] { j });
                formula = Invoke(item, "Formula", BindingFlags.GetProperty, new object[0]);
            }
            catch (Exception) { continue; }
            if (Says(formula as string, want))
                rows.Add(new object[] { kind, where, prefix + j.ToString(Inv), formula });
        }
    }

    // exist(collection, keys, refuse): the items' own names of the keys that name an item of the collection
    // (Item(key) does not fail and the item's Name is the key, case ignored; when Item gives an item of another name,
    // the one named so is looked for among the items: Named), in the keys' order; with refuse true a key that is
    // there fails the operation instead (Taken and those names).
    static object Exist(object collection, object[] keys, object refuse)
    {
        if (keys == null) throw new Exception("exist: the keys are a list of texts");
        var found = new List<object>();
        foreach (object key in keys)
        {
            object name;
            try
            {
                object item = Invoke(collection, "Item", Get, new object[] { key });
                name = Invoke(item, "Name", BindingFlags.GetProperty, new object[0]);
            }
            catch (Exception) { continue; }                       // not an item
            if (string.Equals(name as string, Convert.ToString(key, Inv), StringComparison.OrdinalIgnoreCase))
                found.Add(name);
            else
            {
                object own = Named(collection, Convert.ToString(key, Inv));
                if (own != null) found.Add(own);
            }
        }
        if (Equals(refuse, true) && found.Count > 0)
        {
            var parts = new string[found.Count];
            for (int i = 0; i < found.Count; i++) parts[i] = Convert.ToString(found[i], Inv);
            throw new Exception(Taken + string.Join(", ", parts));
        }
        return found.ToArray();
    }

    // The Name of the collection's item named key (case ignored), by a walk over its items, or null: only when
    // Item(key) gave an item of another name (with a name both the workbook and the active sheet have, Excel's
    // Workbook.Names.Item gives the sheet's own, live 20.11.2026).
    static object Named(object collection, string key)
    {
        int n = Convert.ToInt32(Invoke(collection, "Count", Get, new object[0]), Inv);
        for (int i = 1; i <= n; i++)
        {
            object name;
            try
            {
                object item = Invoke(collection, "Item", Get, new object[] { i });
                name = Invoke(item, "Name", BindingFlags.GetProperty, new object[0]);
            }
            catch (Exception) { continue; }                       // unreadable: not this one
            if (string.Equals(name as string, key, StringComparison.OrdinalIgnoreCase)) return name;
        }
        return null;
    }

    // each(collection, paths): for every item (Count, then Item(1..Count)), the value at each dotted path ("*": the
    // rest of the path over every item of the collection reached); null where it cannot be read.
    static object Each(object collection, object[] paths)
    {
        if (paths == null) throw new Exception("each: the paths are a list of names");
        var split = new string[paths.Length][];
        for (int j = 0; j < paths.Length; j++) split[j] = Convert.ToString(paths[j], Inv).Split('.');
        int n = Convert.ToInt32(Invoke(collection, "Count", Get, new object[0]), Inv);
        var rows = new List<object>();
        for (int i = 1; i <= n; i++)
        {
            object item = Invoke(collection, "Item", Get, new object[] { i });
            var row = new object[paths.Length];
            for (int j = 0; j < paths.Length; j++) row[j] = PathValue(item, split[j], 0);
            rows.Add(row);
        }
        return rows.ToArray();
    }

    // Whether name is a property of obj's type that reads without arguments, from its type information (cached
    // per interface GUID and member, as the COM executor keys it; a type without a real GUID is never cached); no
    // type information or no such member: false, cached only for a name the type does not know or a walk that went
    // through. Excel runs a method even when it is asked for as a property (measured 20.11.2026: Delete in a path
    // deleted the sheet), so each asks this first. Every descriptor is released, also when a call throws midway.
    static readonly Dictionary<string, bool> Properties = new Dictionary<string, bool>();
    static readonly Guid DispatchIid = new Guid("00020400-0000-0000-C000-000000000046");
    const int UnknownName = unchecked((int)0x80020006);         // DISP_E_UNKNOWNNAME

    static bool IsProperty(object obj, string name)
    {
        var d = obj as IDispatchRaw;
        if (d == null) return false;
        IntPtr pti;
        if (d.GetTypeInfo(0, Lcid, out pti) < 0 || pti == IntPtr.Zero) return false;
        var ti = (CT.ITypeInfo)Marshal.GetObjectForIUnknown(pti);
        Marshal.Release(pti);
        IntPtr pta = IntPtr.Zero;
        try
        {
            ti.GetTypeAttr(out pta);
            var ta = (CT.TYPEATTR)Marshal.PtrToStructure(pta, typeof(CT.TYPEATTR));
            string key = ta.guid == Guid.Empty || ta.guid == DispatchIid
                ? null : ta.guid.ToString("N") + "." + name.ToLowerInvariant();
            bool known;
            if (key != null && Properties.TryGetValue(key, out known)) return known;
            var none = Guid.Empty;
            var ids = new int[1];
            int hr = d.GetIDsOfNames(ref none, new string[] { name }, 1, Lcid, ids);
            if (hr < 0)
            {
                if (key != null && hr == UnknownName) Properties[key] = false;
                return false;
            }
            bool prop = false;
            for (int i = 0; i < ta.cFuncs && !prop; i++)
            {
                IntPtr pfd;
                ti.GetFuncDesc(i, out pfd);
                try
                {
                    var fd = (CT.FUNCDESC)Marshal.PtrToStructure(pfd, typeof(CT.FUNCDESC));
                    prop = fd.memid == ids[0] && (fd.invkind & CT.INVOKEKIND.INVOKE_PROPERTYGET) != 0
                           && fd.cParams == fd.cParamsOpt;
                }
                finally
                {
                    ti.ReleaseFuncDesc(pfd);
                }
            }
            for (int i = 0; i < ta.cVars && !prop; i++)
            {
                IntPtr pvd;
                ti.GetVarDesc(i, out pvd);
                try
                {
                    var vd = (CT.VARDESC)Marshal.PtrToStructure(pvd, typeof(CT.VARDESC));
                    prop = vd.memid == ids[0];
                }
                finally
                {
                    ti.ReleaseVarDesc(pvd);
                }
            }
            if (key != null) Properties[key] = prop;
            return prop;
        }
        finally
        {
            if (pta != IntPtr.Zero) ti.ReleaseTypeAttr(pta);
        }
    }

    // nudge(object, property, delta, above): the numeric property + delta when it is more than above (or above is
    // null); the value it ends with. A value that is no number fails.
    static object Nudge(object obj, string name, object delta, object above)
    {
        if (!IsProperty(obj, name)) throw new Exception(name + " is not a property");
        object v = Invoke(obj, name, BindingFlags.GetProperty, new object[0]);
        if (!(v is double || v is int || v is long || v is short || v is float || v is decimal || v is byte))
            throw new Exception(name + " is not a number");
        double x = Convert.ToDouble(v, Inv);
        if (above != null && !(above is DBNull) && !(x > Convert.ToDouble(above, Inv))) return v;
        double nx = x + Convert.ToDouble(delta, Inv);
        Invoke(obj, name, BindingFlags.SetProperty, new object[] { nx });
        return nx;
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
                if (!IsProperty(obj, parts[k])) return null;      // a method's name: never called (Excel would run
                                                                  // it even for a property get)
                obj = Invoke(obj, parts[k], BindingFlags.GetProperty, new object[0]);
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
