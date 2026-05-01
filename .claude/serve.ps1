$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$port = 5173

$mime = @{
  ".html" = "text/html; charset=utf-8"
  ".htm"  = "text/html; charset=utf-8"
  ".css"  = "text/css; charset=utf-8"
  ".js"   = "application/javascript; charset=utf-8"
  ".json" = "application/json; charset=utf-8"
  ".svg"  = "image/svg+xml"
  ".png"  = "image/png"
  ".jpg"  = "image/jpeg"
  ".jpeg" = "image/jpeg"
  ".gif"  = "image/gif"
  ".ico"  = "image/x-icon"
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Serving $root at http://localhost:$port/"

try {
  while ($listener.IsListening) {
    try {
      $ctx = $listener.GetContext()
    } catch {
      break
    }

    try {
      $req = $ctx.Request
      $res = $ctx.Response

      $rel = [System.Uri]::UnescapeDataString($req.Url.AbsolutePath).TrimStart('/')
      if (-not $rel) { $rel = "index.html" }
      $file = Join-Path $root $rel

      $resolved = $null
      try { $resolved = (Resolve-Path -LiteralPath $file -ErrorAction Stop).Path } catch {}

      if ($resolved -and $resolved.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        $bytes = [System.IO.File]::ReadAllBytes($resolved)
        $ext = [System.IO.Path]::GetExtension($resolved).ToLower()
        $res.ContentType = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { "application/octet-stream" }
        $res.StatusCode = 200
        if ($req.HttpMethod -ne "HEAD") {
          $res.OutputStream.Write($bytes, 0, $bytes.Length)
        }
      } else {
        $res.StatusCode = 404
        $msg = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found: $rel")
        if ($req.HttpMethod -ne "HEAD") {
          $res.OutputStream.Write($msg, 0, $msg.Length)
        }
      }
    } catch {
      Write-Host "Request error: $($_.Exception.Message)"
      try { $ctx.Response.StatusCode = 500 } catch {}
    } finally {
      try { $ctx.Response.Close() } catch {}
    }
  }
} finally {
  $listener.Stop()
}
