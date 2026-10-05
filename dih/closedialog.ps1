# Auto-closes Task Dialog save prompts of the bds.exe WE start for a build
# (the IDE asks whether to save a .dproj it upgraded).
# Uses TDM_CLICK_BUTTON (0x0466) with IDNO (7) to click "No"/"Don't Save".
#
# Usage: powershell -NoProfile -File closedialog.ps1 -ProcessName bds
#                   [-TimeoutSeconds 900] [-PidFile <file>]
#
# THREE RULES, each of them from a defect (user report 2026-10-05: "at the end
# it says press any key, but then I only land on a new line and have to close
# the window with the X"):
#  * IT MUST END. The old loop was `while ($true)` and only broke after it had
#    closed a dialog - so in the normal case (no dialog at all) it polled
#    forever. Started with `start /b` it shares the caller's CONSOLE, so that
#    console cannot close while it lives: the install looked finished, the
#    keypress ended cmd, and the window stayed until the X killed the orphan.
#  * ONLY OUR OWN IDE. It matched ANY process called bds, which on a developer
#    machine is the user's running IDE - and it would post "No" into THAT
#    IDE's dialogs, for as long as it lived. Only a process that started AFTER
#    this watcher can be the build IDE.
#  * SAY WHO YOU ARE. With -PidFile the caller can stop the watcher the moment
#    the build is over instead of relying on the timeout.
param(
    [string]$ProcessName = "bds",
    [int]$TimeoutSeconds = 900,
    [string]$PidFile = ""
)

$started = Get-Date
if ($PidFile) {
    try { $PID | Set-Content -Path $PidFile -Encoding ASCII } catch { }
}

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class DialogCloser {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    const uint TDM_CLICK_BUTTON = 0x0466;
    const int IDNO = 7;

    public static bool CloseDialogsForProcess(int processId) {
        bool found = false;
        EnumWindows((hWnd, lParam) => {
            uint wndPid;
            GetWindowThreadProcessId(hWnd, out wndPid);
            if ((int)wndPid != processId) return true;

            StringBuilder className = new StringBuilder(256);
            GetClassName(hWnd, className, 256);
            if (className.ToString() != "#32770" || !IsWindowVisible(hWnd)) return true;

            PostMessage(hWnd, TDM_CLICK_BUTTON, (IntPtr)IDNO, IntPtr.Zero);
            found = true;
            return false;
        }, IntPtr.Zero);
        return found;
    }
}
"@

# The build IDE: a process of that name which did not exist before us. The
# user's own IDE started earlier and is none of our business.
function Get-BuildIdeProcesses {
    @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object {
        try { $_.StartTime -ge $started } catch { $false }
    })
}

$deadline = $started.AddSeconds($TimeoutSeconds)
$seen = $false
while ((Get-Date) -lt $deadline) {
    $procs = Get-BuildIdeProcesses
    if ($procs.Count -eq 0) {
        # Gone again? Then the build is over and there is nothing left to
        # watch. Never seen one? Keep waiting - it is still starting up.
        if ($seen) { break }
        Start-Sleep -Milliseconds 500
        continue
    }
    $seen = $true
    foreach ($p in $procs) {
        [DialogCloser]::CloseDialogsForProcess($p.Id) | Out-Null
    }
    Start-Sleep -Milliseconds 300
}

if ($PidFile) {
    try { Remove-Item -Path $PidFile -Force -ErrorAction SilentlyContinue } catch { }
}
