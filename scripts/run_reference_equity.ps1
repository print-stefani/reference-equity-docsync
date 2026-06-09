param(
  [Parameter(Mandatory = $false)]
  [string]$DocPath = "samples\docs\reference_catalog.docx",

  [Parameter(Mandatory = $false)]
  [string]$ModelDir = "samples\models",

  [Parameter(Mandatory = $false)]
  [string]$OutDir = "output\demo",

  [Parameter(Mandatory = $false)]
  [string]$PythonPath = "",

  [Parameter(Mandatory = $false)]
  [switch]$IncludePass,

  [Parameter(Mandatory = $false)]
  [switch]$EnrichWeb
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$pipelinePath = Join-Path $projectRoot "validator_py\docsync_pipeline.py"

function Resolve-ProjectPath {
  param([string]$PathValue)
  if ([string]::IsNullOrWhiteSpace($PathValue)) { return $null }
  if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
  return Join-Path $projectRoot $PathValue
}

function Test-PythonExecutable {
  param([string]$ExePath)
  if ([string]::IsNullOrWhiteSpace($ExePath)) { return $false }
  if (-not (Test-Path -LiteralPath $ExePath)) { return $false }
  & $ExePath -c "import sys; print(sys.version)" *> $null
  return ($LASTEXITCODE -eq 0)
}

$DocPath = Resolve-ProjectPath $DocPath
$ModelDir = Resolve-ProjectPath $ModelDir
$OutDir = Resolve-ProjectPath $OutDir

if ([string]::IsNullOrWhiteSpace($PythonPath)) {
  $localVenvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
  if (Test-PythonExecutable $localVenvPython) {
    $PythonPath = $localVenvPython
  } else {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python -and (Test-PythonExecutable $python.Source)) {
      $PythonPath = $python.Source
    } else {
      $py = Get-Command py -ErrorAction SilentlyContinue
      if ($py -and (Test-PythonExecutable $py.Source)) {
        $PythonPath = $py.Source
      }
    }
  }
} else {
  $PythonPath = Resolve-ProjectPath $PythonPath
}

if (-not (Test-PythonExecutable $PythonPath)) { throw "Nenhum executável Python válido foi encontrado." }
if (-not (Test-Path -LiteralPath $pipelinePath)) { throw "Pipeline não encontrado: $pipelinePath" }
if (-not (Test-Path -LiteralPath $DocPath)) { throw "Arquivo de documentação não encontrado: $DocPath" }
if (-not (Test-Path -LiteralPath $ModelDir)) { throw "Diretório de modelos não encontrado: $ModelDir" }

$argsList = @(
  $pipelinePath,
  "--doc", $DocPath,
  "--xdmz-dir", $ModelDir,
  "--out", $OutDir,
  "--skip-web-post"
)

if ($IncludePass.IsPresent) { $argsList += "--include-pass-in-dashboard" }
if ($EnrichWeb.IsPresent) { $argsList += "--enrich-web" }

Write-Host "Executando Reference Equity DocSync"
Write-Host "  PY    : $PythonPath"
Write-Host "  DOC   : $DocPath"
Write-Host "  MODELS: $ModelDir"
Write-Host "  OUT   : $OutDir"

& $PythonPath @argsList
if ($LASTEXITCODE -ne 0) {
  throw "Pipeline falhou com exit code $LASTEXITCODE."
}
