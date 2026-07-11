# Script to fix ANSI/UTF-8 encoding corruption in VB6 .frm and .bas files
# Replaces EF BF BD byte sequences with correct ANSI Spanish characters
# Based on word-level context mapping
# Usage: powershell -ExecutionPolicy Bypass -File fix-encoding.ps1

# Corrupted word -> Correct ANSI byte sequence (as hex strings)
# Key: The corrupted bytes (as hex), Value: The correct ANSI bytes (as hex)
# Corrupted char is always EF BF BD (3 bytes) -> Replace with correct 1-byte ANSI char

function Replace-CorruptedWord {
    param([byte[]]$Bytes, [string]$CorruptedHex, [string]$CorrectHex)
    
    $corrupted = [byte[]]($CorruptedHex -split ' ' | ForEach-Object { [Convert]::ToByte($_, 16) })
    $correct = [byte[]]($CorrectHex -split ' ' | ForEach-Object { [Convert]::ToByte($_, 16) })
    
    $result = $Bytes
    $i = 0
    $count = 0
    while ($i -le $result.Length - $corrupted.Length) {
        $match = $true
        for ($j = 0; $j -lt $corrupted.Length; $j++) {
            if ($result[$i + $j] -ne $corrupted[$j]) { $match = $false; break }
        }
        if ($match) {
            # Replace corrupted bytes with correct bytes
            $newResult = New-Object byte[] ($result.Length - $corrupted.Length + $correct.Length)
            [Array]::Copy($result, 0, $newResult, 0, $i)
            [Array]::Copy($correct, 0, $newResult, $i, $correct.Length)
            [Array]::Copy($result, $i + $corrupted.Length, $newResult, $i + $correct.Length, $result.Length - $i - $corrupted.Length)
            $result = $newResult
            $i += $correct.Length
            $count++
        } else {
            $i++
        }
    }
    return @{Result=$result; Count=$count}
}

# Build replacement table: [CorruptedHex] = CorrectHex
# E.g. "EF BF BD" (corrupted) -> "F3" (ó in ANSI)
$replaceTable = @{}

# Helper: add a replacement where corrupted bytes [hex] map to correct single char [hex]
function Add-Replacement([string]$Word, [string]$Correct) {
    # Convert word to bytes, find all EF BF BD positions, build replacement
    $wordBytes = [byte[]]$Word.Split(' ') | ForEach-Object { [Convert]::ToByte($_, 16) }
    $correctBytes = [byte[]]$Correct.Split(' ') | ForEach-Object { [Convert]::ToByte($_, 16) }
    
    # Build full replacement: keep ASCII unchanged, replace EF BF BD with correct char
    $fullCorrupted = @()
    $fullCorrect = @()
    $i = 0
    while ($i -lt $wordBytes.Length) {
        if ($i -le $wordBytes.Length - 3 -and $wordBytes[$i] -eq 0xEF -and $wordBytes[$i+1] -eq 0xBF -and $wordBytes[$i+2] -eq 0xBD) {
            # This is a corrupted char position
            $fullCorrupted += 0xEF, 0xBF, 0xBD
            $fullCorrect += $correctBytes[$fullCorrect.Length / $correctBytes.Length]  # cycle through correct bytes
            $i += 3
        } else {
            $fullCorrupted += $wordBytes[$i]
            $fullCorrect += $wordBytes[$i]
            $i++
        }
    }
    
    $replaceTable[$fullCorrupted -join ' '] = $fullCorrect -join ' '
}

# Now define all replacements using hex byte sequences
# Each corrupted word in hex, and its correct ANSI equivalent in hex

# First, the corrupted word patterns (hex, with EF BF BD for each corrupted char)
# ó (0xF3) replacements - most common
$corrections = @(
    @("43 6F 6E 66 69 67 75 72 61 63 69 EF BF BD 6E", "43 6F 6E 66 69 67 75 72 61 63 69 F3 6E"), # Configuración
    @("4F 70 65 72 61 63 69 EF BF BD 6E", "4F 70 65 72 61 63 69 F3 6E"), # Operación
    @("54 72 61 6E 73 61 63 63 69 EF BF BD 6E", "54 72 61 6E 73 61 63 63 69 F3 6E"), # Transacción
    @("44 65 73 63 72 69 70 63 69 EF BF BD 6E", "44 65 73 63 72 69 70 63 69 F3 6E"), # Descripción
    @("44 69 72 65 63 63 69 EF BF BD 6E", "44 69 72 65 63 63 69 F3 6E"), # Dirección
    @("44 65 66 69 6E 69 63 69 EF BF BD 6E", "44 65 66 69 6E 69 63 69 F3 6E"), # Definición
    @("41 63 74 75 61 6C 69 7A 61 63 69 EF BF BD 6E", "41 63 74 75 61 6C 69 7A 61 63 69 F3 6E"), # Actualización
    @("45 6C 69 6D 69 6E 61 63 69 EF BF BD 6E", "45 6C 69 6D 69 6E 61 63 69 F3 6E"), # Eliminación
    @("43 6F 74 69 7A 61 63 69 EF BF BD 6E", "43 6F 74 69 7A 61 63 69 F3 6E"), # Cotización
    @("4C 69 71 75 69 64 61 63 69 EF BF BD 6E", "4C 69 71 75 69 64 61 63 69 F3 6E"), # Liquidación
    @("49 6E 66 6F 72 6D 61 63 69 EF BF BD 6E", "49 6E 66 6F 72 6D 61 63 69 F3 6E"), # Información
    @("49 6D 70 72 65 73 69 EF BF BD 6E", "49 6D 70 72 65 73 69 F3 6E"), # Impresión
    @("4D 6F 64 69 66 69 63 61 63 69 EF BF BD 6E", "4D 6F 64 69 66 69 63 61 63 69 F3 6E"), # Modificación
    @("46 61 63 74 75 72 61 63 69 EF BF BD 6E", "46 61 63 74 75 72 61 63 69 F3 6E"), # Facturación
    @("41 70 6C 69 63 61 63 69 EF BF BD 6E", "41 70 6C 69 63 61 63 69 F3 6E"), # Aplicación
    @("44 6F 63 75 6D 65 6E 74 61 63 69 EF BF BD 6E", "44 6F 63 75 6D 65 6E 74 61 63 69 F3 6E"), # Documentación
    @("43 61 6E 63 65 6C 61 63 69 EF BF BD 6E", "43 61 6E 63 65 6C 61 63 69 F3 6E"), # Cancelación
    @("43 6F 6E 73 69 67 6E 61 63 69 EF BF BD 6E", "43 6F 6E 73 69 67 6E 61 63 69 F3 6E"), # Consignación
    @("41 6E 75 6C 61 63 69 EF BF BD 6E", "41 6E 75 6C 61 63 69 F3 6E"), # Anulación
    @("45 6D 69 73 69 EF BF BD 6E", "45 6D 69 73 69 F3 6E"), # Emisión
    @("42 6F 6E 69 66 69 63 61 63 69 EF BF BD 6E", "42 6F 6E 69 66 69 63 61 63 69 F3 6E"), # Bonificación
    @("44 69 73 6D 69 6E 75 63 69 EF BF BD 6E", "44 69 73 6D 69 6E 75 63 69 F3 6E"), # Disminución
    @("43 6F 6E 64 69 63 69 EF BF BD 6E", "43 6F 6E 64 69 63 69 F3 6E"), # Condición
    @("44 69 76 69 73 69 EF BF BD 6E", "44 69 76 69 73 69 F3 6E"), # División
    @("4F 70 63 69 EF BF BD 6E", "4F 70 63 69 F3 6E"), # Opción
    
    # Words with ñ (0xF1)
    @("43 6F 6D 70 61 EF BF BD 69 61", "43 6F 6D 70 61 F1 69 61"), # Compañía
    @("43 6F 6D 70 61 EF BF BD 69 61 73", "43 6F 6D 70 61 F1 69 61 73"), # Compañías
    @("43 4F 4D 50 41 EF BF BD 49 41", "43 4F 4D 50 41 F1 49 41"), # COMPAÑÍA
    @("63 6F 6D 70 61 EF BF BD 69 61", "63 6F 6D 70 61 F1 69 61"), # compañía
    
    # Words with ó in other positions
    @("4D EF BF BD 64 75 6C 6F", "4D F3 64 75 6C 6F"), # Módulo
    @("4D EF BF BD 64 75 6C 6F 33", "4D F3 64 75 6C 6F 33"), # Módulo3
    @("64 EF BF BD 61", "64 ED 61"), # día
    @("44 EF BF BD 61", "44 ED 61"), # Día
    @("44 EF BF BD 61 73", "44 ED 61 73"), # Días
    
    # Words with ú (0xFA)
    @("54 EF BF BD 74 75 6C 6F", "54 FA 74 75 6C 6F"), # Título
    @("50 65 72 EF BF BD", "50 65 72 FA"), # Perú
    
    # Words with é (0xE9)
    @("43 65 72 72 EF BF BD", "43 65 72 72 E9"), # Cerró
    @("65 73 74 EF BF BD", "65 73 74 E1"), # está with á (0xE1)
    
    # Inverted question mark (0xBF) - but EF BF BD is also the corruption pattern!
    # These start with the corrupted char
    @("EF BF BD 44 65 73 65 61", "BF 44 65 73 65 61"), # ¿Desea
    @("EF BF BD 46 61 63", "BF 46 61 63"), # ¿Fac
    
    # Relación
    @("52 65 6C 61 63 69 EF BF BD 6E", "52 65 6C 61 63 69 F3 6E"), # Relación
    @("52 61 7A EF BF BD 6E", "52 61 7A F3 6E"), # Razón
    
    @("52 65 73 EF BF BD 6D 65 6E", "52 65 73 FA 6D 65 6E"),
    
    # "N° Lista a Usar :" -> "Nro Lista a Usar :"
    @("4E EF BF BD 20 4C 69 73 74 61 20 61 20 55 73 61 72 20 3A", "4E 72 6F 20 4C 69 73 74 61 20 61 20 55 73 61 72 20 3A"),
    
    # General: N° -> Nro (aplica a todos los N° corrompidos)
    @("4E EF BF BD", "4E 72 6F")
)

# Apply all corrections
$replaceTable = @{}
foreach ($correction in $corrections) {
    $corruptedHex = $correction[0]
    $correctHex = $correction[1]
    $replaceTable[$corruptedHex] = $correctHex
}

# Sort by descending length of corrupted hex (longest match first)
$sortedKeys = $replaceTable.Keys | Sort-Object { $_.Split(' ').Count } -Descending

function Fix-FileEncoding {
    param([string]$FilePath)
    
    Write-Host "Processing: $FilePath"
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    $originalBytes = $bytes
    $totalReplacements = 0
    
    foreach ($key in $sortedKeys) {
        $corrupted = [byte[]]($key.Split(' ') | ForEach-Object { [Convert]::ToByte($_, 16) })
        $correct = [byte[]]($replaceTable[$key].Split(' ') | ForEach-Object { [Convert]::ToByte($_, 16) })
        
        $i = 0
        while ($i -le $bytes.Length - $corrupted.Length) {
            $match = $true
            for ($j = 0; $j -lt $corrupted.Length; $j++) {
                if ($bytes[$i + $j] -ne $corrupted[$j]) { $match = $false; break }
            }
            if ($match) {
                $newBytes = New-Object byte[] ($bytes.Length - $corrupted.Length + $correct.Length)
                [Array]::Copy($bytes, 0, $newBytes, 0, $i)
                [Array]::Copy($correct, 0, $newBytes, $i, $correct.Length)
                [Array]::Copy($bytes, $i + $corrupted.Length, $newBytes, $i + $correct.Length, $bytes.Length - $i - $corrupted.Length)
                $bytes = $newBytes
                $i += $correct.Length
                $totalReplacements++
            } else {
                $i++
            }
        }
    }
    
    if ($totalReplacements -gt 0) {
        [System.IO.File]::WriteAllBytes($FilePath, $bytes)
        Write-Host "  -> $totalReplacements replacements made" -ForegroundColor Green
        return $true
    }
    else {
        # Check if file has ANY EF BF BD sequences
        $remaining = 0
        for ($i = 0; $i -lt $bytes.Length - 2; $i++) {
            if ($bytes[$i] -eq 0xEF -and $bytes[$i+1] -eq 0xBF -and $bytes[$i+2] -eq 0xBD) { $remaining++ }
        }
        if ($remaining -gt 0) {
            Write-Host "  -> No known patterns matched, $remaining unresolved" -ForegroundColor Yellow
        } else {
            Write-Host "  -> No corrupted chars found" -ForegroundColor Gray
        }
        return $false
    }
}

$files = Get-ChildItem -LiteralPath (Get-Location).Path -Recurse -Include "*.frm", "*.bas" | Where-Object { $_.Extension -match '\.(frm|bas)$' } | Sort-Object FullName

$fixed = 0
$total = 0
foreach ($file in $files) {
    $total++
    if (Fix-FileEncoding -FilePath $file.FullName) { $fixed++ }
}

Write-Host "`n========================================"
Write-Host "Complete."
Write-Host "Files processed: $total"
Write-Host "Files with fixes: $fixed"
Write-Host "========================================"
