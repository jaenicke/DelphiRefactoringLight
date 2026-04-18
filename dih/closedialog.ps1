# Auto-closes Task Dialog save prompts from bds.exe
# Uses TDM_CLICK_BUTTON (0x0466) with IDNO (7) to click "No"/"Don't Save"
# Usage: powershell -NoProfile -File closedialog.ps1 -ProcessName bds
param([string]$ProcessName = "bds")

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

# Wait for the process to start, then poll for dialogs until the process exits
$closed = 0
while ($true) {
    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $proc) {
        if ($closed -gt 0) { break }  # Process gone after we closed dialogs - done
        Start-Sleep -Milliseconds 500
        continue
    }
    if ([DialogCloser]::CloseDialogsForProcess($proc.Id)) {
        $closed++
        Start-Sleep -Milliseconds 200  # Brief pause before checking for next dialog
    } else {
        Start-Sleep -Milliseconds 500
    }
}
