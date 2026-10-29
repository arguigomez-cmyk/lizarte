<#
.SYNOPSIS
  Comprueba que el conector local GESTIONA_for_LZRT sigue en la configuración
  de Claude Desktop y, si ha desaparecido, lo vuelve a insertar desde una copia.

.DESCRIPTION
  - Busca claude_desktop_config.json en las dos rutas posibles:
      %APPDATA%\Claude\  (instalación normal)
      %LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\  (Microsoft Store / MSIX)
      %LOCALAPPDATA%\Claude-Data\claude_*\  (instalación del PC de mgomez)
  - Si la entrada existe: guarda una copia de ella en
      %LOCALAPPDATA%\LizarteMCP\GESTIONA_for_LZRT.json
    (carpeta local, no sincronizada con OneDrive ni perfil móvil).
  - Si la entrada falta y hay copia: hace un .bak del fichero actual y reinserta la entrada.
  - Si el JSON está roto: avisa y NO toca nada.
  - Comprueba que existe el .exe al que apunta "command".
  - Muestra las últimas líneas de los logs de MCP.

  OJO: la copia contiene la contraseña de IA_USER en texto plano (igual que el
  propio claude_desktop_config.json). No la envíes por correo ni chat.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\comprobar-gestiona-mcp.ps1

.EXAMPLE
  # Desde una tarea programada al iniciar sesión, sin mostrar logs:
  powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File .\comprobar-gestiona-mcp.ps1 -Silencioso
#>
param(
    [string]$NombreServidor = 'GESTIONA_for_LZRT',
    [switch]$Silencioso
)

$ErrorActionPreference = 'Stop'

$dirCopia  = Join-Path $env:LOCALAPPDATA 'LizarteMCP'
$copia     = Join-Path $dirCopia "$NombreServidor.json"
$registro  = Join-Path $dirCopia 'comprobar-gestiona-mcp.log'
$utf8SinBom = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path $dirCopia)) { New-Item -ItemType Directory -Path $dirCopia | Out-Null }

function Escribir([string]$texto, [string]$color = 'Gray') {
    $linea = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $texto
    Add-Content -Path $registro -Value $linea -Encoding UTF8
    if (-not $Silencioso) { Write-Host $texto -ForegroundColor $color }
}

# --- 1. Localizar los ficheros de configuración -----------------------------
$candidatos = @(Join-Path $env:APPDATA 'Claude\claude_desktop_config.json')
$candidatos += Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Packages') -Filter 'Claude_*' -Directory -ErrorAction SilentlyContinue |
    ForEach-Object { Join-Path $_.FullName 'LocalCache\Roaming\Claude\claude_desktop_config.json' }
$candidatos += Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Claude-Data') -Filter 'claude_*' -Directory -ErrorAction SilentlyContinue |
    ForEach-Object { Join-Path $_.FullName 'claude_desktop_config.json' }
$configs = @($candidatos | Where-Object { Test-Path $_ })

if ($configs.Count -eq 0) {
    Escribir "No se encuentra ningún claude_desktop_config.json. ¿Está instalado Claude Desktop?" 'Red'
    exit 1
}
if ($configs.Count -gt 1) {
    Escribir "ATENCIÓN: hay $($configs.Count) copias del fichero de configuración (causa 2). Se revisan todas:" 'Yellow'
    $configs | ForEach-Object { Escribir "  $_" 'Yellow' }
}

if ((Get-Process -Name 'Claude' -ErrorAction SilentlyContinue) -and -not $Silencioso) {
    Write-Host "Claude Desktop está abierto: si hay que reinsertar la entrada, ciérralo del todo después y vuelve a abrirlo." -ForegroundColor Yellow
}

# --- 2. Revisar cada fichero ------------------------------------------------
foreach ($ruta in $configs) {
    Escribir "`n== $ruta" 'Cyan'

    try {
        $texto = [IO.File]::ReadAllText($ruta)
        $cfg = if ($texto.Trim()) { $texto | ConvertFrom-Json } else { New-Object PSObject }
    } catch {
        Escribir "El JSON está roto (causa 1): $($_.Exception.Message)" 'Red'
        Escribir "No se modifica. Restaura un .bak o corrige la coma/comillas/barras '\' que falten." 'Red'
        continue
    }

    $servidores = $cfg.PSObject.Properties['mcpServers']
    $entrada = if ($servidores) { $servidores.Value.PSObject.Properties[$NombreServidor] } else { $null }

    if ($entrada) {
        Escribir "OK: '$NombreServidor' está en la configuración." 'Green'
        [IO.File]::WriteAllText($copia, ($entrada.Value | ConvertTo-Json -Depth 20), $utf8SinBom)
        Escribir "Copia de la entrada actualizada en $copia"
        $def = $entrada.Value
    }
    elseif (Test-Path $copia) {
        Escribir "FALTA '$NombreServidor'. Se reinserta desde la copia." 'Yellow'
        $bak = "$ruta.bak-{0:yyyyMMdd-HHmmss}" -f (Get-Date)
        Copy-Item $ruta $bak
        Escribir "Copia del fichero actual: $bak"

        $def = [IO.File]::ReadAllText($copia) | ConvertFrom-Json
        if (-not $servidores) {
            $cfg | Add-Member -NotePropertyName 'mcpServers' -NotePropertyValue (New-Object PSObject)
        }
        $cfg.mcpServers | Add-Member -NotePropertyName $NombreServidor -NotePropertyValue $def -Force
        [IO.File]::WriteAllText($ruta, ($cfg | ConvertTo-Json -Depth 20), $utf8SinBom)
        Escribir "Entrada reinsertada. Cierra Claude Desktop del todo (también en la bandeja) y vuelve a abrirlo." 'Green'
    }
    else {
        Escribir "FALTA '$NombreServidor' y no hay copia guardada en $copia." 'Red'
        Escribir "Hay que volver a configurarlo a mano una vez; a partir de ahí este script guardará la copia." 'Red'
        continue
    }

    # --- 3. Comprobar el ejecutable (causa 4) -------------------------------
    $cmd = $def.command
    if ($cmd) {
        $cmdExpandido = [Environment]::ExpandEnvironmentVariables($cmd)
        if ((Test-Path $cmdExpandido) -or (Get-Command $cmdExpandido -ErrorAction SilentlyContinue)) {
            Escribir "OK: el ejecutable existe ($cmdExpandido)." 'Green'
        } else {
            Escribir "El ejecutable NO existe: $cmdExpandido (causa 4: ¿se actualizó o reinstaló Python?)." 'Red'
        }
    }
}

# --- 4. Últimas líneas de los logs -------------------------------------------
if (-not $Silencioso) {
    $dirLogs = @(Join-Path $env:APPDATA 'Claude\logs')
    $dirLogs += Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Claude-Data') -Filter 'claude_*' -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Join-Path $_.FullName 'logs' }
    $dirLogs += Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Packages') -Filter 'Claude_*' -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Join-Path $_.FullName 'LocalCache\Roaming\Claude\logs' }
    foreach ($d in $dirLogs | Where-Object { Test-Path $_ }) {
        foreach ($f in @('mcp.log', "mcp-server-$NombreServidor.log")) {
            $p = Join-Path $d $f
            if (Test-Path $p) {
                Write-Host "`n--- Últimas 15 líneas de $p" -ForegroundColor Cyan
                Get-Content $p -Tail 15
            }
        }
    }
}
