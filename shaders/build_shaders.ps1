# Compile vortex.frag → vortex.frag.qsb using PySide6's bundled qsb tool.
# Run once before building the EXE, or whenever vortex.frag changes.

$ErrorActionPreference = "Stop"
$qsb = python -c "import PySide6, pathlib; print(pathlib.Path(PySide6.__file__).parent / 'qsb.exe')"
if (-not (Test-Path $qsb)) {
    Write-Error "qsb.exe not found at: $qsb  -- make sure PySide6 is installed."
    exit 1
}

$dir = $PSScriptRoot
& $qsb --glsl "100 es,120,150" --hlsl 50 --msl 12 `
       -o "$dir\vortex.frag.qsb" "$dir\vortex.frag"

Write-Host "Compiled: $dir\vortex.frag.qsb"
