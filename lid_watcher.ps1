# Lid watcher (for "S0 low power idle" laptops).
# Turning the screen off through Windows makes this laptop enter standby and the server stops answering.
# So when the lid is closed we keep Windows "awake with display on" and instead tell each external monitor
# to power itself off over DDC/CI (VCP code 0xD6). When the lid is opened the monitors are powered on again.
$created = $false
$mutex = New-Object System.Threading.Mutex($false, 'WarehouseLidWatcher', [ref]$created)
if (-not $created) { exit }   # already running
Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class LidWatcher : Form {
    public static string LogPath = null;
    static void Log(string t) { try { if (LogPath != null) System.IO.File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + t + "\r\n"); } catch { } }

    [DllImport("user32.dll")] static extern IntPtr RegisterPowerSettingNotification(IntPtr hRecipient, ref Guid PowerSettingGuid, int Flags);
    [DllImport("kernel32.dll")] static extern uint SetThreadExecutionState(uint esFlags);
    delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdc, IntPtr rect, IntPtr data);
    [DllImport("user32.dll")] static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr clip, MonitorEnumProc proc, IntPtr data);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct PHYSICAL_MONITOR { public IntPtr hPhysicalMonitor; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string desc; }
    [DllImport("dxva2.dll")] static extern bool GetNumberOfPhysicalMonitorsFromHMONITOR(IntPtr hMonitor, out uint count);
    [DllImport("dxva2.dll")] static extern bool GetPhysicalMonitorsFromHMONITOR(IntPtr hMonitor, uint size, [Out] PHYSICAL_MONITOR[] arr);
    [DllImport("dxva2.dll")] static extern bool SetVCPFeature(IntPtr hMonitor, byte code, uint value);
    [DllImport("dxva2.dll")] static extern bool DestroyPhysicalMonitors(uint size, PHYSICAL_MONITOR[] arr);

    static Guid LID = new Guid("BA3E0F4D-B817-4094-A2D1-D56379E6A0F3");
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    struct PBS { public Guid PowerSetting; public int DataLength; public byte Data; }
    bool first = true;
    Timer offTimer = new Timer();

    public LidWatcher() {
        ShowInTaskbar = false; FormBorderStyle = FormBorderStyle.None; Opacity = 0; Width = 1; Height = 1;
        offTimer.Interval = 2000;   // wait until Windows has rearranged displays after the lid closed
        offTimer.Tick += (o, e) => { offTimer.Stop(); SetMonitorPower(4); };
    }
    protected override void SetVisibleCore(bool value) { if (!IsHandleCreated) CreateHandle(); base.SetVisibleCore(false); }
    protected override void OnHandleCreated(EventArgs e) {
        base.OnHandleCreated(e);
        uint r = SetThreadExecutionState(0x80000001);   // ES_CONTINUOUS | ES_SYSTEM_REQUIRED
        Guid g = LID; IntPtr h = RegisterPowerSettingNotification(Handle, ref g, 0);
        Log("watcher started (DDC mode), execstate=" + r + ", register=" + h);
    }
    // value: 1 = on, 4 = off (DPM off)
    static void SetMonitorPower(uint value) {
        List<IntPtr> mons = new List<IntPtr>();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (hm, hdc, rc, d) => { mons.Add(hm); return true; }, IntPtr.Zero);
        foreach (IntPtr hm in mons) {
            uint n;
            if (!GetNumberOfPhysicalMonitorsFromHMONITOR(hm, out n) || n == 0) { Log("monitor: no physical handle"); continue; }
            PHYSICAL_MONITOR[] arr = new PHYSICAL_MONITOR[n];
            if (!GetPhysicalMonitorsFromHMONITOR(hm, n, arr)) { Log("monitor: get handle failed"); continue; }
            foreach (PHYSICAL_MONITOR pm in arr) {
                bool ok = SetVCPFeature(pm.hPhysicalMonitor, 0xD6, value);
                Log("monitor '" + pm.desc + "' power " + (value == 1 ? "ON" : "OFF") + ": " + (ok ? "ok" : "failed (DDC/CI not supported or disabled)"));
            }
            DestroyPhysicalMonitors(n, arr);
        }
    }
    protected override void WndProc(ref Message m) {
        if (m.Msg == 0x218 && m.WParam.ToInt32() == 0x8013) {
            PBS p = (PBS)Marshal.PtrToStructure(m.LParam, typeof(PBS));
            if (p.PowerSetting == LID) {
                Log("lid state=" + p.Data + (first ? " (initial)" : ""));
                first = false;                 // initial state counts too (e.g. started while the lid is already closed)
                if (p.Data == 0) {        // lid closed: keep display "on" for Windows (no standby), power monitors off via DDC
                    SetThreadExecutionState(0x80000003);   // + ES_DISPLAY_REQUIRED
                    offTimer.Start();
                } else {                       // lid opened: monitors back on, normal display timeout again
                    offTimer.Stop();
                    SetMonitorPower(1);
                    SetThreadExecutionState(0x80000001);
                }
            }
        }
        base.WndProc(ref m);
    }
}
"@
$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory $logDir | Out-Null }
[LidWatcher]::LogPath = Join-Path $logDir "lid.log"
[System.Windows.Forms.Application]::Run((New-Object LidWatcher))
