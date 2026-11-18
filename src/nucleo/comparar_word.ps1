# Copia con control de cambios de un informe revisado: compara la versión que se revisó (con sus comentarios) con la
# versión final, muestra los cambios como revisiones de Word y responde cada comentario con un texto de un JSON.
#
# Pasos (Word por COM, como actualizar_campos_word.ps1):
#   1. abre la versión revisada y, si aún tiene cambios propuestos, los rechaza, para comparar contra el texto que se
#      revisó; los comentarios se conservan. Si la versión revisada trae muchos cambios de formato registrados, conviene
#      rechazarlos antes en el XML con `uv run python src/nucleo/rechazar_cambios.py <revisado> <base>` y pasar la base
#      como -Revisado: Revisions.RejectAll puede tardar horas en un informe largo (ADR-015);
#   2. compara ese texto con la versión final (CompareDocuments: palabra por palabra, sin formato ni campos, con notas al
#      pie y comentarios); las revisiones llevan el autor indicado;
#   3. agrega a cada comentario la respuesta del JSON ({"01": "texto", "02": ...}), por el número de la observación
#      («Observación NN»), por su orden entre los comentarios raíz («orden_NNN», el JSON de src/nucleo/observaciones.py)
#      o por la clave «sin_numero» para los comentarios sin número.
#
# Uso:
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/nucleo/comparar_word.ps1 -Revisado <docx con comentarios> `
#       -Final <docx final> -Salida <docx de salida> [-Autor "Equipo evaluador"] [-Respuestas <json>] `
#       [-AutorRespuestas "<quien firma las respuestas>"] [-Iniciales "<iniciales>"]
param(
    [Parameter(Mandatory = $true)][string]$Revisado,
    [Parameter(Mandatory = $true)][string]$Final,
    [Parameter(Mandatory = $true)][string]$Salida,
    [string]$Autor = "Equipo evaluador",
    [string]$Respuestas = "",
    [string]$AutorRespuestas = "",   # quien firma las respuestas a los comentarios; por defecto, el autor de los cambios
    [string]$Iniciales = ""          # por defecto, las iniciales de las palabras con mayúscula de quien firma
)
if ($AutorRespuestas -eq "") { $AutorRespuestas = $Autor }
if ($Iniciales -eq "") {
    $Iniciales = (($AutorRespuestas -split '\s+') | Where-Object { $_ -cmatch '^\p{Lu}' } | ForEach-Object { $_[0] }) -join ""
}
$ErrorActionPreference = "Stop"
$Revisado = (Resolve-Path -LiteralPath $Revisado).Path
$Final = (Resolve-Path -LiteralPath $Final).Path
$base = Join-Path $env:TEMP ("base_revision_" + [guid]::NewGuid().ToString() + ".docx")
$resp = @{}
if ($Respuestas -ne "") {
    $json = Get-Content -LiteralPath $Respuestas -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($p in $json.PSObject.Properties) { $resp[$p.Name] = $p.Value }
}
function Avance([string]$texto) { Write-Output ("{0:HH:mm:ss} {1}" -f (Get-Date), $texto) }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false
$usuario = $word.UserName
$iniciales = $word.UserInitials
# opciones de la instalación de Word que se restituyen al final: ortografía y gramática al escribir y paginación de fondo
$ortografia = $word.Options.CheckSpellingAsYouType
$gramatica = $word.Options.CheckGrammarAsYouType
$paginacion = $word.Options.Pagination
$datosLocales = $null
try { $datosLocales = $word.Options.UseLocalUserInfo } catch { }   # Word 2013 o posterior
$word.Options.CheckSpellingAsYouType = $false
$word.Options.CheckGrammarAsYouType = $false
$word.Options.Pagination = $false
try {
    $r = $word.Documents.Open($Revisado, $false, $true, $false)
    $propuestos = $r.Revisions.Count
    Avance ("versión revisada abierta: {0} cambios propuestos" -f $propuestos)
    if ($propuestos -gt 0) {
        $r.Close($false)
        $r = $word.Documents.Open($Revisado, $false, $false, $false)
        $r.TrackRevisions = $false
        $r.Revisions.RejectAll()
        $r.SaveAs2($base, 16)
        $r.Close($false)
        Avance "cambios propuestos rechazados"
        $orig = $word.Documents.Open($base, $false, $false, $false)
    } else {
        $orig = $r
    }
    $fin = $word.Documents.Open($Final, $false, $true, $false)
    Avance "versión final abierta; comparando"
    # CompareDocuments(Original, Revised, Destination=nuevo, Granularity=palabra, Formatting, CaseChanges, Whitespace,
    #                  Tables, Headers, Footnotes, Textboxes, Fields, Comments, Moves, RevisedAuthor, IgnoreWarnings)
    $res = $word.CompareDocuments($orig, $fin, 2, 1, $false, $true, $false, $true, $false, $true, $true, $false, $true,
                                  $false, $Autor, $true)
    $orig.Close($false)
    $fin.Close($false)
    Avance "comparación terminada; respondiendo los comentarios"
    # las respuestas llevan el autor indicado, no la cuenta con la que se inició sesión en Office
    if ($null -ne $datosLocales) { $word.Options.UseLocalUserInfo = $true }
    $word.UserName = $AutorRespuestas
    $word.UserInitials = $Iniciales
    $sin = 0
    $respondidos = 0
    $n = $res.Comments.Count
    # orden de los comentarios raíz en el documento, para las claves «orden_NNN» de src/nucleo/observaciones.py
    $raices = New-Object System.Collections.Generic.List[int]
    for ($j = 1; $j -le $n; $j++) { if ($null -eq $res.Comments.Item($j).Ancestor) { $raices.Add($j) } }
    # de atrás hacia adelante: cada respuesta se inserta detrás de su comentario y corre los índices siguientes
    for ($i = $n; $i -ge 1; $i--) {
        $c = $res.Comments.Item($i)
        if ($null -ne $c.Ancestor) { continue }   # solo los comentarios raíz
        $texto = $c.Range.Text
        $m = [regex]::Match($texto, "Observaci.n (\d+)")
        $clave = if ($m.Success) { $m.Groups[1].Value.PadLeft(2, "0") } else { "sin_numero" }
        $porOrden = "orden_{0:D3}" -f ($raices.IndexOf($i) + 1)
        if ($resp.ContainsKey($porOrden)) { $clave = $porOrden }
        if ($resp.ContainsKey($clave)) {
            [void]$c.Replies.Add($c.Scope, [string]$resp[$clave])
            $respondidos++
        } else { $sin++ }
    }
    $res.TrackRevisions = $true
    $res.SaveAs2($Salida, 16)
    Avance (("cambios propuestos rechazados en la base: {0}; revisiones en la copia: {1}; comentarios: {2}; " +
             "respondidos: {3}; sin respuesta: {4}") -f $propuestos, $res.Revisions.Count, $n, $respondidos, $sin)
    $res.Close($false)
}
finally {
    $word.UserName = $usuario
    $word.UserInitials = $iniciales
    $word.Options.CheckSpellingAsYouType = $ortografia
    $word.Options.CheckGrammarAsYouType = $gramatica
    $word.Options.Pagination = $paginacion
    if ($null -ne $datosLocales) { $word.Options.UseLocalUserInfo = $datosLocales }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base }
}
