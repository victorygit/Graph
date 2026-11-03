param(
    [Parameter(Mandatory = $true)]
    [string]$DotFile,

    [Parameter(Mandatory = $true)]
    [string]$Target,

    [string]$OutputFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DotFile -PathType Leaf)) {
    throw "DOT file not found: $DotFile"
}

$edgePattern = '^\s*"(?<from>[^"]+)"\s*->\s*"(?<to>[^"]+)"'
$edges = [System.Collections.Generic.List[object]]::new()
$nodes = [System.Collections.Generic.HashSet[string]]::new()
$lineNumber = 0
$lines = [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $DotFile).Path)
foreach ($line in $lines) {
    $lineNumber++
    if ($line -match $edgePattern) {
        $from = $Matches['from']
        $to = $Matches['to']
        [void]$nodes.Add($from)
        [void]$nodes.Add($to)
        $edges.Add([pscustomobject]@{
            LineNumber = $lineNumber
            From = $from
            To = $to
        })
    }
}

if (-not $nodes.Contains($Target)) {
    throw "Target '$Target' is not an endpoint in $DotFile"
}

$ancestors = [System.Collections.Generic.HashSet[string]]::new()
[void]$ancestors.Add($Target)
$changed = $true
while ($changed) {
    $changed = $false
    foreach ($edge in $edges) {
        if ($ancestors.Contains($edge.To) -and $ancestors.Add($edge.From)) {
            $changed = $true
        }
    }
}

$targetEdges = @($edges | Where-Object {
    $ancestors.Contains($_.From) -and $ancestors.Contains($_.To)
})

$inputPath = [System.IO.Path]::GetFullPath($DotFile)
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $safeTarget = $Target -replace '[^A-Za-z0-9._-]', '_'
    $outputName = '{0}.{1}.highlighted.dot' -f [System.IO.Path]::GetFileNameWithoutExtension($inputPath), $safeTarget
    $OutputFile = Join-Path ([System.IO.Path]::GetDirectoryName($inputPath)) $outputName
}
$outputPath = [System.IO.Path]::GetFullPath($OutputFile)
if ([string]::Equals($inputPath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputFile must not overwrite the input DOT file.'
}

$highlightedLines = [System.Collections.Generic.List[string]]::new()
$highlightedLineNumbers = [System.Collections.Generic.HashSet[int]]::new()
foreach ($edge in $targetEdges) {
    [void]$highlightedLineNumbers.Add($edge.LineNumber)
}

$graphOpeningFound = $false
for ($index = 0; $index -lt $lines.Length; $index++) {
    $line = $lines[$index]
    if ($highlightedLineNumbers.Contains($index + 1)) {
        $edgeMatch = [regex]::Match(
            $line,
            '^(?<edge>\s*"[^"]+"\s*->\s*"[^"]+")(?<attributes>\s*\[[^\]]*\])?(?<ending>\s*;?\s*(?://.*)?)$'
        )
        if (-not $edgeMatch.Success) {
            throw "Could not apply highlight to edge on line $($index + 1)."
        }

        $attributes = $edgeMatch.Groups['attributes'].Value
        if ($attributes) {
            $attributes = $attributes -replace '\]\s*$', ', color="red", penwidth=2.5]'
        }
        else {
            $attributes = ' [color="red", penwidth=2.5]'
        }
        $line = $edgeMatch.Groups['edge'].Value + $attributes + $edgeMatch.Groups['ending'].Value
    }

    $highlightedLines.Add($line)
    if (-not $graphOpeningFound -and $line -match '^\s*digraph\b.*\{\s*$') {
        $highlightedLines.Add(('    "{0}" [style=filled, fillcolor="yellow", color="red", penwidth=2];' -f $Target))
        $graphOpeningFound = $true
    }
}

if (-not $graphOpeningFound) {
    throw "Could not find the opening digraph declaration in $DotFile."
}

[System.IO.File]::WriteAllLines($outputPath, $highlightedLines.ToArray())
$graphviz = Get-Command dot.exe -ErrorAction SilentlyContinue
if (-not $graphviz) {
    Write-Warning "Graphviz dot.exe was not found. The highlighted DOT file was saved, but no image was rendered: $outputPath"
}
else {
    $imagePath = [System.IO.Path]::ChangeExtension($outputPath, '.svg')
    & $graphviz.Source -Tsvg $outputPath -o $imagePath
    if ($LASTEXITCODE -ne 0) {
        throw "Graphviz failed to render the highlighted graph (exit code $LASTEXITCODE). DOT file: $outputPath"
    }
}

Write-Output "Upstream edges for '$Target' in $DotFile ($($targetEdges.Count) edges):"
foreach ($edge in $targetEdges) {
    Write-Output ("Line {0}: {1} -> {2}" -f $edge.LineNumber, $edge.From, $edge.To)
}
Write-Output "Highlighted DOT: $outputPath"
if ($graphviz) {
    Write-Output "Highlighted graph: $imagePath"
}
