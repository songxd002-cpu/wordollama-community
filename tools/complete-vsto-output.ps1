param(
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [string[]]$SearchRoots = @()
)

# The VSTO application manifest lists Content files that the build does not copy
# into bin. Word refuses to load the add-in when a listed file is missing or its
# hash differs, so copy the missing files in and verify every listed hash.
$ErrorActionPreference = 'Stop'
$manifestPath = Join-Path $OutputDir 'WordOllama.dll.manifest'
[xml]$manifest = Get-Content -Raw -Encoding UTF8 $manifestPath
$ns = New-Object Xml.XmlNamespaceManager $manifest.NameTable
$ns.AddNamespace('asmv2', 'urn:schemas-microsoft-com:asm.v2')
$ns.AddNamespace('dsig', 'http://www.w3.org/2000/09/xmldsig#')

$entries = @()
foreach ($node in $manifest.SelectNodes('//asmv2:file', $ns)) {
    $entries += [pscustomobject]@{ Path = $node.GetAttribute('name'); Node = $node }
}
foreach ($node in $manifest.SelectNodes('//asmv2:dependentAssembly[@codebase]', $ns)) {
    $entries += [pscustomobject]@{ Path = $node.GetAttribute('codebase'); Node = $node }
}

function Get-Digest([string]$path, [string]$algorithm) {
    $hasher = if ($algorithm -match 'sha256') { [Security.Cryptography.SHA256]::Create() } else { [Security.Cryptography.SHA1]::Create() }
    $stream = [IO.File]::OpenRead($path)
    try { return [Convert]::ToBase64String($hasher.ComputeHash($stream)) } finally { $stream.Dispose() }
}

$candidates = foreach ($root in $SearchRoots) {
    if (Test-Path $root) { Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notlike "$((Resolve-Path $OutputDir).Path)*" } }
}

$errors = @()
foreach ($entry in $entries) {
    $relative = $entry.Path
    $target = Join-Path $OutputDir $relative
    $expected = $entry.Node.SelectSingleNode('asmv2:hash/dsig:DigestValue', $ns).InnerText
    $algorithm = $entry.Node.SelectSingleNode('asmv2:hash/dsig:DigestMethod', $ns).GetAttribute('Algorithm')

    if (-not (Test-Path -LiteralPath $target)) {
        $leaf = Split-Path $relative -Leaf
        $match = $candidates | Where-Object { $_.Name -eq $leaf -and (Get-Digest $_.FullName $algorithm) -eq $expected } | Select-Object -First 1
        if (-not $match) { $errors += "missing and not found: $relative"; continue }
        New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent) | Out-Null
        Copy-Item -LiteralPath $match.FullName -Destination $target
        Write-Host "Copied $relative from $($match.FullName)"
    }

    if ((Get-Digest $target $algorithm) -ne $expected) { $errors += "hash mismatch: $relative" }
}

if ($errors) { throw "VSTO output is incomplete:`n$($errors -join "`n")" }
Write-Host "Verified $($entries.Count) manifest entries in $OutputDir."
