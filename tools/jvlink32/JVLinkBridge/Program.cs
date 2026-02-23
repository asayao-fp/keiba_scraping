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

// CLI args override env vars (positional: dataspec fromdate option)
if (args.Length >= 1) dataspec = args[0];
if (args.Length >= 2) fromdate = args[1];
if (args.Length >= 3 && int.TryParse(args[2], out var oa)) option = oa;

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
    Type jvType = Type.GetTypeFromProgID("JVDTLab.JVLink")
        ?? throw new InvalidOperationException("JVDTLab.JVLink ProgID not found. Is JV-Link installed?");

    object jv = Activator.CreateInstance(jvType)
        ?? throw new InvalidOperationException("Failed to create JVDTLab.JVLink instance.");

    object? Invoke(string method, params object[] p) =>
        jvType.InvokeMember(method,
            System.Reflection.BindingFlags.InvokeMethod, null, jv, p);

    // ── setup ────────────────────────────────────────────────────────────────
    result.Stage = "init";
    int initRet      = (int)(Invoke("JVInit",        0)        ?? -9999);
    int savePathRet  = (int)(Invoke("JVSetSavePath", savePath) ?? -9999);
    int saveFlagRet  = (int)(Invoke("JVSetSaveFlag", 1)        ?? -9999);
    int payFlagRet   = (int)(Invoke("JVSetPayFlag",  0)        ?? -9999);

    // ── optional UI properties ────────────────────────────────────────────────
    int? uiPropertiesRet = null;
    if (enableUiProperties)
    {
        try
        {
            uiPropertiesRet = (int)(Invoke("JVSetUIProperties", 0, 0, 0, 0) ?? -9999);
        }
        catch (Exception uiEx)
        {
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
    // Signature: JVOpen(dataspec, fromdate, option,
    //                   ref readcount, ref downloadcount, ref lastfiletimestamp)
    object[] openArgs = [dataspec, fromdate, option, 0, 0, ""];
    int openRet       = (int)(Invoke("JVOpen", openArgs) ?? -9999);
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
    }
    else
    {
        // ── post-open sleep ───────────────────────────────────────────────────
        if (sleepAfterOpenSec > 0)
            Thread.Sleep((int)(sleepAfterOpenSec * 1000));

        // ── JVStatus polling ──────────────────────────────────────────────────
        var statusSnapshots = new List<StatusSnapshot>();
        if (enableStatusPoll || requireStatusZero)
        {
            result.Stage = "status_poll";
            var pollDeadline = DateTime.UtcNow.AddSeconds(statusPollMaxWaitSec);
            while (DateTime.UtcNow < pollDeadline)
            {
                int statusRet;
                try
                {
                    statusRet = (int)(Invoke("JVStatus") ?? -9999);
                }
                catch (COMException comEx)
                {
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
        }

        // ── gate: require JVStatus == 0 before calling JVRead ────────────────
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
            }
        }

        if (proceedToRead)
        {
        // ── JVRead with retry ─────────────────────────────────────────────────
        result.Stage = "read";
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

        while (DateTime.UtcNow < deadline)
        {
            var readArgs = new object[] { "", 0, "" };
            var pm = new ParameterModifier(3);
            pm[0] = true; pm[1] = true; pm[2] = true;

            try
            {
                readRet = (int)(jvType.InvokeMember("JVRead",
                    BindingFlags.InvokeMethod,
                    null, jv, readArgs, [pm], null, null) ?? -9999);
            }
            catch
            {
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
    result.Close = (int)(Invoke("JVClose") ?? -9999);
    result.Ok    = result.Error is null;
    // Restore the stage to where the failure originated (not "close")
    if (savedErrorStage is not null) result.Stage = savedErrorStage;

    Marshal.ReleaseComObject(jv);
}
catch (COMException comEx)
{
    result.Ok      = false;
    result.Hresult = $"0x{(uint)comEx.ErrorCode:X8}";
    result.Error   = comEx.Message;
    // result.Stage already set to the stage where the fault occurred
}
catch (Exception ex)
{
    result.Ok    = false;
    result.Error = ex.ToString();
    // result.Stage already set
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
