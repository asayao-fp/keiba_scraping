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
//   JVBRIDGE_DEBUG              1     (emit step-by-step diagnostic logs to stderr)
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

// CLI args override env vars (positional: dataspec fromdate option [--debug-steps])
bool debugSteps = EnvBool("JVBRIDGE_DEBUG");
var positional = new List<string>();
foreach (var arg in args)
{
    if (arg == "--debug-steps") debugSteps = true;
    else positional.Add(arg);
}
if (positional.Count >= 1) dataspec = positional[0];
if (positional.Count >= 2) fromdate = positional[1];
if (positional.Count >= 3 && int.TryParse(positional[2], out var oa)) option = oa;

// ── debug logging ────────────────────────────────────────────────────────────

static string SafeTrunc(string? s, int max = 200) =>
    s is null ? "<null>" : s.Length <= max ? s : s[..max] + $"…(+{s.Length - max})";

static string ArgDesc(object? v) =>
    v is string s ? $"string({s.Length})" :
    v is null     ? "null" :
                    v.GetType().Name;

void DbgStep(string msg)
{
    if (!debugSteps) return;
    var ts = DateTime.UtcNow.ToString("HH:mm:ss.fff");
    var tid = Environment.CurrentManagedThreadId;
    Console.Error.WriteLine($"[DBG {ts} T{tid}] {msg}");
    Console.Error.Flush();
}

void DbgComBefore(string method, object[] parms)
{
    if (!debugSteps) return;
    var argDescs = string.Join(", ", ((object?[])parms).Select(ArgDesc));
    DbgStep($"COM>> {method}({argDescs})");
}

void DbgComAfter(string method, object? ret, object?[]? outParms = null)
{
    if (!debugSteps) return;
    var retStr = ret is null ? "null" : ret.ToString();
    var msg = $"COM<< {method} ret={retStr}";
    if (outParms != null)
    {
        var outs = string.Join(", ", outParms.Select(p =>
            p is string sv ? $"\"{SafeTrunc(sv)}\"" : (p?.ToString() ?? "null")));
        msg += $" out=[{outs}]";
    }
    DbgStep(msg);
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

DbgStep($"start pid={Environment.ProcessId} debug={debugSteps}");
DbgStep($"params dataspec={dataspec} fromdate={fromdate} option={option} savePath={SafeTrunc(savePath,80)}");

var result = new BridgeResult();

try
{
    DbgStep("COM: GetTypeFromProgID JVDTLab.JVLink");
    Type jvType = Type.GetTypeFromProgID("JVDTLab.JVLink")
        ?? throw new InvalidOperationException("JVDTLab.JVLink ProgID not found. Is JV-Link installed?");
    DbgStep($"COM: type={jvType.FullName}");

    DbgStep("COM: Activator.CreateInstance");
    object jv = Activator.CreateInstance(jvType)
        ?? throw new InvalidOperationException("Failed to create JVDTLab.JVLink instance.");
    DbgStep("COM: instance created");

    object? Invoke(string method, params object[] p)
    {
        DbgComBefore(method, p);
        var ret = jvType.InvokeMember(method,
            System.Reflection.BindingFlags.InvokeMethod, null, jv, p);
        DbgComAfter(method, ret);
        return ret;
    }

    // ── setup ────────────────────────────────────────────────────────────────
    result.Stage = "init";
    DbgStep("stage=init");
    int initRet      = (int)(Invoke("JVInit",        0)        ?? -9999);
    int savePathRet  = (int)(Invoke("JVSetSavePath", savePath) ?? -9999);
    int saveFlagRet  = (int)(Invoke("JVSetSaveFlag", 1)        ?? -9999);
    int payFlagRet   = (int)(Invoke("JVSetPayFlag",  0)        ?? -9999);
    DbgStep($"init done: init={initRet} savePath={savePathRet} saveFlag={saveFlagRet} payFlag={payFlagRet}");

    // ── optional UI properties ────────────────────────────────────────────────
    int? uiPropertiesRet = null;
    if (enableUiProperties)
    {
        DbgStep("COM: JVSetUIProperties (optional)");
        try
        {
            uiPropertiesRet = (int)(Invoke("JVSetUIProperties", 0, 0, 0, 0) ?? -9999);
        }
        catch (Exception uiEx)
        {
            DbgStep($"JVSetUIProperties exception: {uiEx.GetType().Name}: {uiEx.Message}");
            uiPropertiesRet = -9999;
            result.UiPropertiesError = uiEx.Message;
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
    DbgStep($"stage=open dataspec={dataspec} fromdate={fromdate} option={option}");
    // Signature: JVOpen(dataspec, fromdate, option,
    //                   ref readcount, ref downloadcount, ref lastfiletimestamp)
    object[] openArgs = [dataspec, fromdate, option, 0, 0, ""];
    DbgComBefore("JVOpen", openArgs);
    int openRet       = (int)(jvType.InvokeMember("JVOpen",
        BindingFlags.InvokeMethod, null, jv, openArgs) ?? -9999);
    int readcount     = Convert.ToInt32(openArgs[3]);
    int downloadcount = Convert.ToInt32(openArgs[4]);
    string lastts     = Convert.ToString(openArgs[5]) ?? "";
    DbgComAfter("JVOpen", openRet, [readcount, downloadcount, lastts]);

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
        DbgStep($"JVOpen failed ret={openRet}");
    }
    else
    {
        DbgStep($"JVOpen ok ret={openRet} readcount={readcount} downloadcount={downloadcount}");

        // ── post-open sleep ───────────────────────────────────────────────────
        if (sleepAfterOpenSec > 0)
        {
            DbgStep($"sleep {sleepAfterOpenSec}s after open");
            Thread.Sleep((int)(sleepAfterOpenSec * 1000));
            DbgStep("sleep done");
        }

        // ── JVStatus polling ──────────────────────────────────────────────────
        var statusSnapshots = new List<StatusSnapshot>();
        if (enableStatusPoll || requireStatusZero)
        {
            result.Stage = "status_poll";
            DbgStep($"stage=status_poll max={statusPollMaxWaitSec}s");
            var pollDeadline = DateTime.UtcNow.AddSeconds(statusPollMaxWaitSec);
            while (DateTime.UtcNow < pollDeadline)
            {
                int statusRet;
                try
                {
                    DbgStep("COM>> JVStatus()");
                    statusRet = (int)(jvType.InvokeMember("JVStatus",
                        BindingFlags.InvokeMethod, null, jv, null) ?? -9999);
                    DbgStep($"COM<< JVStatus ret={statusRet}");
                }
                catch (COMException comEx)
                {
                    DbgStep($"JVStatus COMException 0x{(uint)comEx.ErrorCode:X8}: {comEx.Message}");
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
                if (statusRet == 0) break; // ready
                Thread.Sleep((int)(statusPollIntervalSec * 1000));
            }
            result.StatusPoll = statusSnapshots;
            DbgStep($"status_poll done snapshots={statusSnapshots.Count}");
        }

        // ── gate: require JVStatus == 0 before calling JVRead ────────────────
        bool proceedToRead = true;
        if (requireStatusZero)
        {
            bool statusZeroSeen = statusSnapshots.Any(s => s.Status == 0);
            if (!statusZeroSeen)
            {
                DbgStep("gate: JVStatus never reached 0, skipping JVRead");
                result.Ok    = false;
                result.Stage = "status_poll";
                result.Error = "JVStatus did not reach 0 within the poll window; JVRead not attempted (JV_READ_REQUIRE_STATUS_ZERO=1)";
                proceedToRead = false;
            }
        }

        if (proceedToRead)
        {
        // ── JVRead with retry ─────────────────────────────────────────────────
        result.Stage = "read";
        DbgStep($"stage=read max={maxWaitSec}s interval={intervalSec}s");
        // Signature (TypeLib): JVRead(out BSTR buff, out int size, out BSTR filename) -> int
        // All three parameters are PARAMFLAG_FOUT (byref/out); pass placeholder values
        // and read back from readArgs after InvokeMember returns.
        bool found     = false;
        int  readRet   = -9999;
        int  size      = 0;
        string buff    = "";
        string filename = "";
        var   deadline  = DateTime.UtcNow.AddSeconds(maxWaitSec);
        var   attempts  = new List<AttemptInfo>();
        int   attempt   = 0;

        while (DateTime.UtcNow < deadline)
        {
            attempt++;
            var readArgs = new object[] { "", 0, "" };
            var pm = new ParameterModifier(3);
            pm[0] = true; pm[1] = true; pm[2] = true;

            if (debugSteps)
                DbgStep($"JVRead attempt={attempt} args=[\"\", 0, \"\"] types=[{string.Join(", ", readArgs.Select(ArgDesc))}]");
            try
            {
                readRet = (int)(jvType.InvokeMember("JVRead",
                    BindingFlags.InvokeMethod,
                    null, jv, readArgs, [pm], null, null) ?? -9999);
            }
            catch (Exception readEx)
            {
                DbgStep($"JVRead attempt={attempt} exception {readEx.GetType().Name}: {readEx.Message}");
                Console.Error.WriteLine($"[EXCEPTION] JVRead attempt={attempt}: {readEx}");
                Console.Error.Flush();
                attempts.Add(new AttemptInfo
                {
                    Ret                   = readRet,
                    Size                  = 0,
                    Filename              = "",
                    BuffHead              = "",
                    ReadArgsTypes         = readArgs.Select(a => a?.GetType().FullName ?? "null").ToArray(),
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

            buff     = Convert.ToString(readArgs[0]) ?? "";
            size     = Convert.ToInt32(readArgs[1]);
            filename = Convert.ToString(readArgs[2]) ?? "";
            DbgStep($"JVRead attempt={attempt} ret={readRet} size={size} filename={SafeTrunc(filename,80)} buff={SafeTrunc(buff,80)}");

            attempts.Add(new AttemptInfo
            {
                Ret                   = readRet,
                Size                  = size,
                Filename              = filename,
                BuffHead              = buff.Length > 30 ? buff[..30] : buff,
                BufferPreview         = buff.Length > 200 ? buff[..200] : buff,
                ReadArgsTypes         = readArgs.Select(a => a?.GetType().FullName ?? "null").ToArray(),
                ReadArgsValuesPreview = GetReadArgsPreview(size, filename, buff),
            });

            if (readRet == 0 && size > 0) { found = true; break; }
            if (readRet == -3) { Thread.Sleep((int)(intervalSec * 1000)); continue; }
            break; // any other value → stop retrying
        }

        DbgStep($"read loop done found={found} readRet={readRet} attempts={attempt}");

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
        }
        } // end if (proceedToRead)
    }

    // ── JVClose ──────────────────────────────────────────────────────────────
    // Save error stage before "close" overwrites it, so the output Stage reflects
    // where the failure actually occurred (e.g. "status_poll", "open"), not "close".
    string? savedErrorStage = result.Error is not null ? result.Stage : null;
    result.Stage = "close";
    DbgStep("stage=close");
    DbgStep("COM>> JVClose()");
    result.Close = (int)(Invoke("JVClose") ?? -9999);
    DbgStep($"COM<< JVClose ret={result.Close}");
    result.Ok    = result.Error is null;
    // Restore the stage to where the failure originated (not "close")
    if (savedErrorStage is not null) result.Stage = savedErrorStage;

    DbgStep("COM: ReleaseComObject");
    Marshal.ReleaseComObject(jv);
    DbgStep("done ok=" + result.Ok);
}
catch (COMException comEx)
{
    result.Ok      = false;
    result.Hresult = $"0x{(uint)comEx.ErrorCode:X8}";
    result.Error   = comEx.Message;
    // result.Stage already set to the stage where the fault occurred
    Console.Error.WriteLine($"[EXCEPTION COMException stage={result.Stage}] 0x{(uint)comEx.ErrorCode:X8}: {comEx}");
    Console.Error.Flush();
}
catch (Exception ex)
{
    result.Ok    = false;
    result.Error = ex.ToString();
    // result.Stage already set
    Console.Error.WriteLine($"[EXCEPTION {ex.GetType().Name} stage={result.Stage}] {ex}");
    Console.Error.Flush();
}

// ── output ───────────────────────────────────────────────────────────────────

Console.WriteLine(JsonSerializer.Serialize(result, jsonOptions));
return result.Ok ? 0 : 1;

// ── record types ─────────────────────────────────────────────────────────────

record BridgeResult
{
    public bool                  Ok               { get; set; }
    public string?               Stage            { get; set; }
    public string?               Error            { get; set; }
    public string?               Hresult          { get; set; }
    public string?               UiPropertiesError { get; set; }
    public SetupInfo?            Setup            { get; set; }
    public OpenInfo?             Open             { get; set; }
    public List<StatusSnapshot>? StatusPoll       { get; set; }
    public ReadInfo?             Read             { get; set; }
    public int?                  Close            { get; set; }
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
