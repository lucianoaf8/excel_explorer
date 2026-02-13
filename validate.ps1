<# 
Validation-only checklist for CSV/XLSX profiling via ydata-profiling + pandas.
- No installs. No venv creation. No shell redirections. No exits.
- Prints PASS/FAIL per check with details.
#>

$ErrorActionPreference = "Continue"
$results = New-Object System.Collections.Generic.List[object]
function Add-Result([string]$Name, [string]$Status, [string]$Detail) {
  $results.Add([pscustomobject]@{ Check = $Name; Status = $Status; Detail = $Detail })
}
function Try-Run([string]$Cmd, [int]$TimeoutSec = 20) {
  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = "powershell.exe"
    $psi.Arguments = "-NoProfile -NonInteractive -Command $Cmd"
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    if (-not $p.WaitForExit($TimeoutSec*1000)) { $p.Kill(); return @{ ok=$false; out=""; err="Timeout $TimeoutSec s" } }
    $out = $p.StandardOutput.ReadToEnd()
    $err = $p.StandardError.ReadToEnd()
    return @{ ok = ($p.ExitCode -eq 0); out = $out.Trim(); err = $err.Trim(); code = $p.ExitCode }
  } catch {
    return @{ ok=$false; out=""; err=$_.Exception.Message; code=$null }
  }
}

Write-Host "== Environment checks =="

# 0) PowerShell + OS
try {
  $psver = $PSVersionTable.PSVersion.ToString()
  $isWin = $IsWindows
  Add-Result "PowerShell version" "PASS" $psver
  Add-Result "OS platform" "PASS" ($(if ($isWin) {"Windows"} else {"Non-Windows"}))
} catch { Add-Result "PowerShell/OS" "FAIL" $_.Exception.Message }

# 1) TEMP write permission
try {
  $tmp = [System.IO.Path]::Combine($env:TEMP, "validate_$(Get-Random).tmp")
  "test" | Set-Content -Path $tmp -Encoding ASCII
  $exists = Test-Path $tmp
  if ($exists) { Remove-Item $tmp -ErrorAction SilentlyContinue }
  Add-Result "TEMP write permission" ($(if ($exists) {"PASS"} else {"FAIL"}), "Temp: $env:TEMP")
} catch { Add-Result "TEMP write permission" "FAIL" $_.Exception.Message }

# 2) PATH basics (WindowsApps is often needed for user shims)
if ($isWin) {
  $apps = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps"
  $inPath = ($env:Path.Split(';') -contains $apps)
  Add-Result "PATH contains WindowsApps" ($(if ($inPath) {"PASS"} else {"WARN"}), $apps)
}

# 3) Python discovery
$foundPythons = @()
# Prefer Python Launcher
$pyLauncher = Get-Command py -ErrorAction SilentlyContinue
if ($pyLauncher) {
  $r = Try-Run "py -0p"
  if ($r.ok) {
    $lines = $r.out -split "`r?`n" | Where-Object { $_ -match "^\s*\d+\.\d+" }
    foreach ($l in $lines) {
      # example line: -3.12-64        C:\Python312\python.exe
      if ($l -match "(-(?<ver>\d+\.\d+).*\s+)(?<path>\S+python\.exe)") {
        $foundPythons += [pscustomobject]@{ Ver = $Matches['ver']; Path = $Matches['path'] }
      }
    }
    Add-Result "Python Launcher (py)" "PASS" $r.out
  } else {
    Add-Result "Python Launcher (py)" "WARN" $r.err
  }
} else {
  Add-Result "Python Launcher (py)" "WARN" "Not found"
}

# Fallback: direct commands
foreach ($cmd in @("python3.12","python3.11","python3.10","python3","python")) {
  $gc = Get-Command $cmd -ErrorAction SilentlyContinue
  if ($gc) {
    $v = Try-Run "$($gc.Source) -c 'import sys;print(f""{sys.version_info.major}.{sys.version_info.minor}"")'"
    if ($v.ok) {
      $foundPythons += [pscustomobject]@{ Ver = $v.out; Path = $gc.Source }
    }
  }
}

if ($foundPythons.Count -gt 0) {
  $dedup = $foundPythons | Sort-Object Ver,Path -Unique
  $list  = ($dedup | ForEach-Object { "$($_.Ver) => $($_.Path)" }) -join "; "
  Add-Result "Python interpreters found" "PASS" $list
} else {
  Add-Result "Python interpreters found" "FAIL" "None on PATH / via launcher"
}

# 4) Determine ydata-profiling compatible interpreter (<3.13, >=3.7). Prefer 3.12, then 3.11, 3.10.
$candidates = $foundPythons | Where-Object {
  $mj,$mn = $_.Ver.Split('.')
  [int]$mj -eq 3 -and [int]$mn -ge 7 -and [int]$mn -lt 13
} | Sort-Object { [version]$_.Ver } -Descending

if ($candidates.Count -gt 0) {
  $preferred = $candidates | Where-Object { $_.Ver -in @("3.12","3.11","3.10") } | Select-Object -First 1
  if (-not $preferred) { $preferred = $candidates[0] }
  Add-Result "Compatible Python (<3.13)" "PASS" "$($preferred.Ver) @ $($preferred.Path)"
} else {
  Add-Result "Compatible Python (<3.13)" "FAIL" "No 3.12/3.11/3.10 detected"
}

# 5) pip availability per candidate (no install)
foreach ($py in $candidates) {
  $pipv = Try-Run "`"$($py.Path)`" -m pip --version"
  $status = $(if ($pipv.ok) {"PASS"} else {"FAIL"})
  Add-Result "pip for Python $($py.Ver)" $status $pipv.out + $(if ($pipv.err){"; " + $pipv.err}else{""})
}

# 6) ydata-profiling availability on index (no install)
if ($candidates.Count -gt 0) {
  $py = $preferred
  $idx = Try-Run "`"$($py.Path)`" -m pip index versions ydata-profiling"
  if ($idx.ok) {
    Add-Result "Index: ydata-profiling versions (Python $($py.Ver))" "PASS" ($idx.out -split "`r?`n" | Select-Object -First 3 | Out-String).Trim()
  } else {
    Add-Result "Index: ydata-profiling versions (Python $($py.Ver))" "WARN" ($idx.err ? $idx.err : "pip index unsupported or offline")
  }
}

# 7) Package import checks (only if installed). No failures if missing; report status.
if ($candidates.Count -gt 0) {
$py = $preferred
$checkPy = [System.IO.Path]::Combine($env:TEMP, "pkg_check_$([Guid]::NewGuid().ToString('N')).py")
@'
import json, importlib, sys
mods = ["ydata_profiling","pandas","openpyxl","pyxlsb"]
res = {}
for m in mods:
    try:
        mod = importlib.import_module(m)
        ver = getattr(mod, "__version__", None)
        if ver is None and hasattr(mod, "__dict__") and "__version__" in mod.__dict__:
            ver = mod.__dict__["__version__"]
        res[m] = {"present": True, "version": ver}
    except Exception as e:
        res[m] = {"present": False, "error": str(e)}
print(json.dumps(res))
'@ | Set-Content -Path $checkPy -Encoding UTF8

  $ir = Try-Run "`"$($py.Path)`" `"$checkPy`""
  if ($ir.ok -and $ir.out) {
    try {
      $obj = $ir.out | ConvertFrom-Json
      foreach ($k in $obj.PSObject.Properties.Name) {
        $o = $obj.$k
        if ($o.present -eq $true) { Add-Result "Import $k (Python $($py.Ver))" "PASS" ("version: " + ($o.version ? $o.version : "n/a")) }
        else { Add-Result "Import $k (Python $($py.Ver))" "WARN" $o.error }
      }
    } catch {
      Add-Result "Imports parse (Python $($py.Ver))" "WARN" "Unexpected output: $($ir.out)"
    }
  } else {
    Add-Result "Imports check (Python $($py.Ver))" "WARN" ($ir.err ? $ir.err : "No output")
  }
  Remove-Item $checkPy -ErrorAction SilentlyContinue
}

# 8) Network reachability to docs (GET tiny range; 10s timeout)
function Test-Url($url) {
  try {
    $resp = Invoke-WebRequest -Uri $url -Headers @{ "Range" = "bytes=0-0" } -MaximumRedirection 5 -TimeoutSec 10 -UseBasicParsing
    $code = if ($resp.StatusCode) { $resp.StatusCode } else { 200 }
    Add-Result "URL $url" ($(if ($code -ge 200 -and $code -lt 400) {"PASS"} else {"WARN"}), "HTTP $code")
  } catch { Add-Result "URL $url" "WARN" $_.Exception.Message }
}
$urls = @(
  "https://pypi.org/project/ydata-profiling/",
  "https://docs.profiling.ydata.ai/latest/getting-started/quickstart/",
  "https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html",
  "https://openpyxl.readthedocs.io/",
  "https://pypi.org/project/pyxlsb/",
  "https://github.com/ydataai/ydata-profiling/issues/1695"
)
foreach ($u in $urls) { Test-Url $u }

# 9) Existing user shim check (optional)
if ($isWin) {
  $shim = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\xlprof.cmd"
  $exists = Test-Path $shim
  Add-Result "Existing xlprof shim" ($(if ($exists) {"PASS"} else {"INFO"}), $(if ($exists) {$shim} else {"Not found (expected if not created yet)"}))
}

# 10) Summaries
Write-Host "`n== Validation results =="
$results | Format-Table -AutoSize

# Also emit a machine-readable JSON blob to copy if needed
try {
  $jsonOut = $results | ConvertTo-Json -Depth 4
  Write-Host "`n== JSON =="
  Write-Output $jsonOut
} catch {
  # ignore JSON failures in constrained terminals
}
