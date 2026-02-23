// JVLinkBridge – thin COM bridge that calls JVRead inside the .NET CLR,
// avoiding the 0xC0000409 crash that occurs when JVRead is invoked from
// Python/pywin32.
//
// Usage (environment variables or CLI args override defaults):
//   JV_DATASPEC                 RACE            (JVOpen dataspec)
//   JV_FROMDATE                 20240101000000  (JVOpen fromdate)
//   JV_OPTION                   1               (JVOpen option)
//   JV_SAVE_PATH                C:\ProgramData\JRA-VAN\Data
//   JV_READ_MAX_WAIT_SEC        60
//   JV_READ_INTERVAL_SEC        0.5
//   JV_SLEEP_AFTER_OPEN_SEC     1.0   (delay after JVOpen before JVRead; set 0 to skip)
//   JV_ENABLE_UI_PROPERTIES     1     (call JVSetUIProperties with safe defaults)
//   JV_ENABLE_STATUS_POLL       1     (poll JVStatus after open until ready or timeout)
//   JV_STATUS_POLL_MAX_WAIT_SEC 10
//   JV_STATUS_POLL_INTERVAL_SEC 0.5
//   JV_READ_REQUIRE_STATUS_ZERO 1     (gate JVRead: require JVStatus==0 before calling JVRead)
//
// Debug:
//   JVBRIDGE_DEBUG              1     (prints step logs to stderr)
//
// Outputs a single JSON line to stdout, then exits 0 on success / 1 on error.

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

// ── helpers ──────────────────────────────────────────────────────────────────

static string Env(string key, string fallback) =>
    Environment.GetEnvironmentVariable(key) is { Length: > 0 } v ? v : fallback;

static double EnvDouble(string key, double fallback) =>
    double.TryParse(Env(key, ""), out var d) ? d : fallback;

static bool EnvBool(string key) =>
    Env(key, "0").Trim() == "1";

static string[] GetReadArgsPreview(int size, string filename, string bufferPreview) => [
    size.ToString(),
    filename,
    bufferPreview.Length > 50 ? bufferPreview[..50] : bufferPreview,
];

// ── parameters ───────────────────────────────────────────────────────────────

string dataspec              = Env("JV_DATASPEC",                "RACE");
string fromdate              = Env("JV_FROMDATE",                "20240101000000");
int    option                = int.TryParse(Env("JV_OPTION", "1"), out var o) ? o : 1;
string savePath              = Env("JV_SAVE_PATH",               @"C:\ProgramData\JRA-VAN\Data");
double maxWaitSec            = EnvDouble("JV_READ_MAX_WAIT_SEC",        60.0);
double intervalSec           = EnvDouble("JV_READ_INTERVAL_SEC",         0.5);
double sleepAfterOpenSec     = EnvDouble("JV_SLEEP_AFTER_OPEN_SEC",      1.0);
bool   enableUiProperties    = EnvBool("JV_ENABLE_UI_PROPERTIES");
bool   enableStatusPoll      = EnvBool("JV_ENABLE_STATUS_POLL");
double statusPollMaxWaitSec  = EnvDouble("JV_STATUS_POLL_MAX_WAIT_SEC", 10.0);
double statusPollIntervalSec = EnvDouble("JV_STATUS_POLL_INTERVAL_SEC",  0.5);
bool   requireStatusZero     = EnvBool("JV_READ_REQUIRE_STATUS_ZERO");
bool   debugSteps            = EnvBool("JVBRIDGE_DEBUG");

// CLI args override env vars (positional: dataspec fromdate option)
if (args.Length >= 1) dataspec = args[0];
if (args.Length >= 2) fromdate = args[1];
if (args.Length >= 3 && int.TryParse(args[2], out var oa)) option = oa;

// ── debug logging (stderr only) ──────────────────────────────────────────────

void D(string msg)
{
    if (!debugSteps) return;
    var ts = DateTime.Now.ToString("HH:mm:ss.fff");
    var tid = Environment.CurrentManagedThreadId;
    Console.Error.WriteLine($"[{ts}][tid:{tid}] {msg}");
}

string PreviewArg(object? a)
{
    if (a is null) return "null";
    if (a is string s)
    {
        var head = s.Length > 80 ? s[..80] + "..." : s;
        return $"string(len={s.Length})=\"{head}\"";
    }
    return $"{a.GetType().FullName}={a}";
}

// ── JSON output options ───────────────────────────────────────────────────────

var jsonOptions = new JsonSerializerOptions
{
    WriteIndented          = false,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    PropertyNamingPolicy   = JsonNamingPolicy.SnakeCaseLower,
    Encoder                = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
};

Console.OutputEncoding = Encoding.UTF8;

// ── COM activation ───────────────────────────────────────────────────────────

var result = new BridgeResult();

try
{
    D("STEP 0: program start");
    D("apartment=" + Thread.CurrentThread.GetApartmentState());
    D($"params: dataspec={dataspec}, fromdate={fromdate}, option={option}, savePath={savePath}");
    D($"flags: enableUiProperties={enableUiProperties}, enableStatusPoll={enableStatusPoll}, requireStatusZero={requireStatusZero}");

    D("STEP 1: Type.GetTypeFromProgID(JVDTLab.JVLink)");
    Type jvType = Type.GetTypeFromProgID("JVDTLab.JVLink")
        ?? throw new InvalidOperationException("JVDTLab.JVLink ProgID not found. Is JV-Link installed?");

    D($"STEP 2: Activator.CreateInstance({jvType.FullName})");
    object jv = Activator.CreateInstance(jvType)
        ?? throw new InvalidOperationException("Failed to create JVDTLab.JVLink instance.");

    dynamic djv = jv;
    D("STEP 3: COM instance created");

    if (debugSteps)
    {
        D("STEP 3.1: Dump IDispatch signature (if available)");
        DispatchIntrospection.DumpDispatchSignature(jv, D);
    }

    object? Invoke(string method, params object[] p)
    {
        D($"CALL {method} args: {string.Join(", ", p.Select(PreviewArg))}");
        var ret = jvType.InvokeMember(method, BindingFlags.InvokeMethod, null, jv, p);
        D($"RET  {method} => {PreviewArg(ret)}");
        return ret;
    }
    // ── setup ────────────────────────────────────────────────────────────────
    result.Stage = "init";

    D("STEP init: JVInit");
    int initRet      = (int)(Invoke("JVInit",        0)        ?? -9999);

    D("STEP init: JVSetSavePath");
    int savePathRet  = (int)(Invoke("JVSetSavePath", savePath) ?? -9999);

    D("STEP init: JVSetSaveFlag");
    int saveFlagRet  = (int)(Invoke("JVSetSaveFlag", 1)        ?? -9999);

    D("STEP init: JVSetPayFlag");
    int payFlagRet   = (int)(Invoke("JVSetPayFlag",  0)        ?? -9999);

    // ── optional UI properties ────────────────────────────────────────────────
    int? uiPropertiesRet = null;
    if (enableUiProperties)
    {
        D("STEP init: JVSetUIProperties (optional)");
        try
        {
            uiPropertiesRet = (int)(Invoke("JVSetUIProperties") ?? -9999);
        }
        catch (Exception uiEx)
        {
            uiPropertiesRet = -9999;
            result.UiPropertiesError = uiEx.Message;
            D("STEP init: JVSetUIProperties exception: " + uiEx);
        }
    }

    result.Setup = new SetupInfo
    {
        Init         = initRet,
        SavePath     = savePathRet,
        SaveFlag     = saveFlagRet,
        PayFlag      = payFlagRet,
        UiProperties = uiPropertiesRet,
    };

    // ── JVOpen ───────────────────────────────────────────────────────────────
    result.Stage = "open";
    object[] openArgs = [dataspec, fromdate, option, 0, 0, ""];
    D("STEP open: before JVOpen openArgs=" + string.Join(", ", openArgs.Select(PreviewArg)));

    var openPm = new ParameterModifier(6);
    openPm[0] = false;
    openPm[1] = false;
    openPm[2] = false;
    openPm[3] = true;
    openPm[4] = true;
    openPm[5] = true;

    int openRet = (int)(jvType.InvokeMember("JVOpen",
        BindingFlags.InvokeMethod,
        null, jv, openArgs, [openPm], null, null) ?? -9999);

    D("STEP open: after  JVOpen openArgs=" + string.Join(", ", openArgs.Select(PreviewArg)));

    int readcount     = Convert.ToInt32(openArgs[3]);
    int downloadcount = Convert.ToInt32(openArgs[4]);
    string lastts     = Convert.ToString(openArgs[5]) ?? "";

    result.Open = new OpenInfo
    {
        Dataspec          = dataspec,
        Fromdate          = fromdate,
        Option            = option,
        Ret               = openRet,
        Readcount         = readcount,
        Downloadcount     = downloadcount,
        Lastfiletimestamp = lastts,
    };

    if (openRet < 0)
    {
        result.Ok    = false;
        result.Error = $"JVOpen returned {openRet}";
        D($"STEP open: openRet={openRet} (error)");
    }
    else
    {
        if (sleepAfterOpenSec > 0)
        {
            D($"STEP post_open_sleep: sleeping {sleepAfterOpenSec}s");
            Thread.Sleep((int)(sleepAfterOpenSec * 1000));
        }

        var statusSnapshots = new List<StatusSnapshot>();
        if (enableStatusPoll || requireStatusZero)
        {
            result.Stage = "status_poll";
            D("STEP status_poll: entering");
            var pollDeadline = DateTime.UtcNow.AddSeconds(statusPollMaxWaitSec);
            while (DateTime.UtcNow < pollDeadline)
            {
                int statusRet;
                try
                {
                    D("CALL JVStatus");
                    statusRet = (int)(Invoke("JVStatus") ?? -9999);
                    D($"RET  JVStatus => {statusRet}");
                }
                catch (COMException comEx)
                {
                    D("STEP status_poll: COMException: " + comEx);
                    statusSnapshots.Add(new StatusSnapshot
                    {
                        Timestamp = DateTime.UtcNow.ToString("o"),
                        Status    = -9999,
                        Error     = comEx.Message,
                        Hresult   = $"0x{(uint)comEx.ErrorCode:X8}",
                    });
                    break;
                }
                statusSnapshots.Add(new StatusSnapshot
                {
                    Timestamp = DateTime.UtcNow.ToString("o"),
                    Status    = statusRet,
                });
                if (statusRet == 0) break;
                Thread.Sleep((int)(statusPollIntervalSec * 1000));
            }
            result.StatusPoll = statusSnapshots;
            D("STEP status_poll: done");
        }

        bool proceedToRead = true;
        if (requireStatusZero)
        {
            bool statusZeroSeen = statusSnapshots.Any(s => s.Status == 0);
            if (!statusZeroSeen)
            {
                result.Ok    = false;
                result.Stage = "status_poll";
                result.Error = "JVStatus did not reach 0 within the poll window; JVRead not attempted (JV_READ_REQUIRE_STATUS_ZERO=1)";
                proceedToRead = false;
                D("STEP gate: status 0 not seen; skipping JVRead");
            }
        }

        if (proceedToRead)
        {
            result.Stage = "read";
            D("STEP read: entering JVRead loop");

            bool found       = false;
            int  readRet     = -9999;
            int  size        = 0;
            string buff      = "";
            string filename  = "";
            var deadline     = DateTime.UtcNow.AddSeconds(maxWaitSec);
            var attempts     = new List<AttemptInfo>();

            while (DateTime.UtcNow < deadline)
            {

                // JVRead loop body (dynamic/ref object to avoid InvokeMember byref issues)
                object b = "";
                object s = 0;
                object f = "";

                D("STEP read: about to invoke JVRead (dynamic/ref)");
                try
                {
                    // Many COM servers behave better with dynamic dispatch than InvokeMember+ParameterModifier
                    readRet = Convert.ToInt32(djv.JVRead(ref b, ref s, ref f));
                }
                catch (Exception ex)
                {
                    D("STEP read: JVRead threw managed exception (dynamic): " + ex);
                    attempts.Add(new AttemptInfo
                    {
                        Ret                   = readRet,
                        Size                  = 0,
                        Filename              = "",
                        BuffHead              = "",
                        ReadArgsTypes         = [b?.GetType().FullName ?? "null", s?.GetType().FullName ?? "null", f?.GetType().FullName ?? "null"],
                        ReadArgsValuesPreview = GetReadArgsPreview(0, "", ""),
                    });
                    result.Read = new ReadInfo
                    {
                        Found        = false,
                        Ret          = readRet,
                        Size         = 0,
                        Filename     = "",
                        BuffHead     = "",
                        AttemptsTail = attempts.Count > 10 ? attempts[^10..] : attempts,
                    };
                    throw;
                }

                buff     = Convert.ToString(b) ?? "";
                size     = Convert.ToInt32(s);
                filename = Convert.ToString(f) ?? "";

                D($"STEP read: JVRead ret={readRet}, size={size}, filename={PreviewArg(filename)}, buff={PreviewArg(buff)}");

                attempts.Add(new AttemptInfo
                {
                    Ret                   = readRet,
                    Size                  = size,
                    Filename              = filename,
                    BuffHead              = buff.Length > 30 ? buff[..30] : buff,
                    BufferPreview         = buff.Length > 200 ? buff[..200] : buff,
                    ReadArgsTypes         = [b?.GetType().FullName ?? "null", s?.GetType().FullName ?? "null", f?.GetType().FullName ?? "null"],
                    ReadArgsValuesPreview = GetReadArgsPreview(size, filename, buff),
                });

                if (readRet == 0 && size > 0) { found = true; break; }
                if (readRet == -3) { Thread.Sleep((int)(intervalSec * 1000)); continue; }
                break;




            }

            result.Read = new ReadInfo
            {
                Found        = found,
                Ret          = readRet,
                Size         = size,
                Filename     = filename,
                BuffHead     = buff.Length > 200 ? buff[..200] : buff,
                AttemptsTail = attempts.Count > 10 ? attempts[^10..] : attempts,
            };

            if (!found && readRet != 0)
            {
                result.Error = $"JVRead returned {readRet}; size={size}; filename={filename}";
                D("STEP read: not found; " + result.Error);
            }
        }
    }

    string? savedErrorStage = result.Error is not null ? result.Stage : null;
    result.Stage = "close";

    D("STEP close: JVClose");
    result.Close = (int)(Invoke("JVClose") ?? -9999);
    D("STEP close: JVClose done");

    result.Ok = result.Error is null;

    if (savedErrorStage is not null) result.Stage = savedErrorStage;

    D("STEP final: ReleaseComObject");
    Marshal.ReleaseComObject(jv);
    D("STEP final: done");
}
catch (COMException comEx)
{
    result.Ok      = false;
    result.Hresult = $"0x{(uint)comEx.ErrorCode:X8}";
    result.Error   = comEx.Message;
    D("CATCH COMException: " + comEx);
}
catch (Exception ex)
{
    result.Ok    = false;
    result.Error = ex.ToString();
    D("CATCH Exception: " + ex);
}

Console.WriteLine(JsonSerializer.Serialize(result, jsonOptions));
return result.Ok ? 0 : 1;

// ── record types ─────────────────────────────────────────────────────────────

record BridgeResult
{
    public bool                  Ok                { get; set; }
    public string?               Stage             { get; set; }
    public string?               Error             { get; set; }
    public string?               Hresult           { get; set; }
    public string?               UiPropertiesError { get; set; }
    public SetupInfo?            Setup             { get; set; }
    public OpenInfo?             Open              { get; set; }
    public List<StatusSnapshot>? StatusPoll        { get; set; }
    public ReadInfo?             Read              { get; set; }
    public int?                  Close             { get; set; }
}

record SetupInfo
{
    public int  Init         { get; set; }
    public int  SavePath     { get; set; }
    public int  SaveFlag     { get; set; }
    public int  PayFlag      { get; set; }
    public int? UiProperties { get; set; }
}

record OpenInfo
{
    public string Dataspec          { get; set; } = "";
    public string Fromdate          { get; set; } = "";
    public int    Option            { get; set; }
    public int    Ret               { get; set; }
    public int    Readcount         { get; set; }
    public int    Downloadcount     { get; set; }
    public string Lastfiletimestamp { get; set; } = "";
}

record StatusSnapshot
{
    public string  Timestamp { get; set; } = "";
    public int     Status    { get; set; }
    public string? Error     { get; set; }
    public string? Hresult   { get; set; }
}

record ReadInfo
{
    public bool              Found        { get; set; }
    public int               Ret          { get; set; }
    public int               Size         { get; set; }
    public string            Filename     { get; set; } = "";
    public string            BuffHead     { get; set; } = "";
    public string?           DecodeError  { get; set; }
    public List<AttemptInfo> AttemptsTail { get; set; } = [];
}

record AttemptInfo
{
    public int      Ret                   { get; set; }
    public int      Size                  { get; set; }
    public string   Filename              { get; set; } = "";
    public string   BuffHead              { get; set; } = "";
    public string?  BufferPreview         { get; set; }
    public string?  DecodeError           { get; set; }
    public int?     Truncated             { get; set; }
    public string[] ReadArgsTypes         { get; set; } = [];
    public string[] ReadArgsValuesPreview { get; set; } = [];
}

// ── IDispatch introspection (to confirm runtime signature of JVRead) ─────────

[ComImport, Guid("00020400-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDispatch
{
    [PreserveSig] int GetTypeInfoCount(out uint pctinfo);
    [PreserveSig] int GetTypeInfo(uint iTInfo, uint lcid, out System.Runtime.InteropServices.ComTypes.ITypeInfo ppTInfo);
    [PreserveSig] int GetIDsOfNames(ref Guid riid, IntPtr rgszNames, uint cNames, uint lcid, IntPtr rgDispId);
    [PreserveSig] int Invoke();
}

static class DispatchIntrospection
{
    static string VarEnumName(short vt) => ((VarEnum)vt).ToString();

    public static void DumpDispatchSignature(object comObj, Action<string> log)
    {
        try
        {
            var disp = (IDispatch)comObj;
            disp.GetTypeInfoCount(out var cnt);
            log($"IDispatch.GetTypeInfoCount={cnt}");
            if (cnt == 0) return;

            disp.GetTypeInfo(0, 0, out var ti);

            ti.GetTypeAttr(out var pTypeAttr);
            var typeAttr = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPEATTR>(pTypeAttr);

            for (int i = 0; i < typeAttr.cFuncs; i++)
            {
                ti.GetFuncDesc(i, out var pFuncDesc);
                var fd = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.FUNCDESC>(pFuncDesc);

                var names = new string[fd.cParams + 1];
                ti.GetNames(fd.memid, names, names.Length, out var got);
                var name = got > 0 ? names[0] : $"memid:{fd.memid}";

                log($"IDispatch func[{i}] name={name} memid={fd.memid} cParams={fd.cParams} returnvt={VarEnumName(fd.elemdescFunc.tdesc.vt)}");

                if (string.Equals(name, "JVRead", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "JVSetUIProperties", StringComparison.OrdinalIgnoreCase))
                {
                    log("=== Runtime IDispatch signature for JVRead ===");
                    log($"return vt={VarEnumName(fd.elemdescFunc.tdesc.vt)}");

                    for (int p = 0; p < fd.cParams; p++)
                    {
                        var elem = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.ELEMDESC>(
                            fd.lprgelemdescParam + p * Marshal.SizeOf<System.Runtime.InteropServices.ComTypes.ELEMDESC>());

                        var pname = (p + 1 < got) ? names[p + 1] : $"param{p}";
                        log($"param[{p}] name={pname} vt={VarEnumName(elem.tdesc.vt)} wParamFlags=0x{elem.desc.paramdesc.wParamFlags:X}");
                    }
                }

                ti.ReleaseFuncDesc(pFuncDesc);
            }

            ti.ReleaseTypeAttr(pTypeAttr);
        }
        catch (Exception ex)
        {
            log("DumpDispatchSignature failed: " + ex);
        }
    }
}