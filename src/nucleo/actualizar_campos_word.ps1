# Actualiza con Word los campos del informe de cierre (numeración SEQ de tablas y figuras, índice general e índices de
# tablas y de figuras) con la paginación real de Word, guarda el .docx y exporta un PDF para la revisión (ronda 26).
# Así el documento abre sin el aviso de actualización de campos y con los números de página correctos. Además incrusta
# las fuentes que no son del sistema (Poppins de la portada; ronda 27).
#
# Uso (lo llama informe_cierre.py):
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/nucleo/actualizar_campos_word.ps1 -Docx <ruta> [-Pdf <ruta>]
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Pdf = ""
)
$ErrorActionPreference = "Stop"
$Docx = (Resolve-Path -LiteralPath $Docx).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $false, $false)
    $doc.Repaginate()
    [void]$doc.Fields.Update()
    foreach ($t in $doc.TablesOfContents) { [void]$t.Update() }
    foreach ($t in $doc.TablesOfFigures) { [void]$t.Update() }
    # los índices ocupan páginas: se repagina y se actualizan otra vez solo los números
    $doc.Repaginate()
    foreach ($t in $doc.TablesOfContents) { [void]$t.UpdatePageNumbers() }
    foreach ($t in $doc.TablesOfFigures) { [void]$t.UpdatePageNumbers() }
    # fuentes que no son del sistema (Poppins de la portada) incrustadas completas, para que el documento se vea igual
    # en equipos sin ellas y se pueda editar (ronda 27; Poppins tiene fsType 0, incrustación permitida)
    $doc.EmbedTrueTypeFonts = $true
    $doc.SaveSubsetFonts = $false
    $doc.DoNotEmbedSystemFonts = $true
    $doc.Save()
    if ($Pdf -ne "") {
        $doc.ExportAsFixedFormat($Pdf, 17)
    }
    Write-Output ("indices: " + $doc.TablesOfContents.Count + " generales, " + $doc.TablesOfFigures.Count +
                  " de tablas y figuras; paginas: " + $doc.ComputeStatistics(2))
    $doc.Close([ref]0)
}
finally {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
