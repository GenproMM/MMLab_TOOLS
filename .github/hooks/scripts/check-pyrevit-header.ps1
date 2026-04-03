# PostToolUse hook: verify PyRevit CPython3 header after file edits
# Fires after any tool use. Checks if a .py file was edited and
# ensures the mandatory #! python3 header is still intact.

$ErrorActionPreference = "SilentlyContinue"

# Read JSON from stdin
$rawInput = [Console]::In.ReadToEnd()
if (-not $rawInput) { exit 0 }

$hookInput = $rawInput | ConvertFrom-Json
if (-not $hookInput) { exit 0 }

# Extract file path from tool input (covers replace_string_in_file, create_file, etc.)
$filePath = $null
if ($hookInput.toolInput.filePath) {
    $filePath = $hookInput.toolInput.filePath
}
elseif ($hookInput.toolInput.path) {
    $filePath = $hookInput.toolInput.path
}

# Also check multi_replace — first replacement's filePath
if (-not $filePath -and $hookInput.toolInput.replacements) {
    $first = $hookInput.toolInput.replacements | Select-Object -First 1
    if ($first.filePath) { $filePath = $first.filePath }
}

# Skip if not a .py file or file doesn't exist
if (-not $filePath) { exit 0 }
if ($filePath -notlike "*.py") { exit 0 }
if (-not (Test-Path $filePath)) { exit 0 }

# Read first two lines
$lines = Get-Content $filePath -TotalCount 2
$line1 = if ($lines.Count -ge 1) { $lines[0].Trim() } else { "" }
$line2 = if ($lines.Count -ge 2) { $lines[1].Trim() } else { "" }

$headerOk = ($line1 -eq "#! python3") -and ($line2 -eq "# -*- coding: utf-8 -*-")

if (-not $headerOk) {
    $msg = @"
CRITICAL: File '$filePath' is missing the required PyRevit CPython3 header.
The first two lines MUST be exactly:
  #! python3
  # -*- coding: utf-8 -*-
Without this header PyRevit will execute the script under IronPython 2 instead of CPython3.
Fix this immediately before making any other changes.
"@
    $output = @{
        continue      = $true
        systemMessage = $msg
    }
    $output | ConvertTo-Json -Compress | Write-Output
    exit 0
}

# Header is fine — no output needed
exit 0
