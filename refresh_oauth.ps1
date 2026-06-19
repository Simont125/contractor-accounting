# OAuth flow pour clasp run - lit les credentials depuis client_secret*.json
$secretFile = Get-ChildItem -Path $PSScriptRoot -Filter "client_secret*.json" | Select-Object -First 1
if (-not $secretFile) {
  Write-Error "Fichier client_secret*.json introuvable dans le dossier du projet."
  exit 1
}

$creds       = Get-Content $secretFile.FullName | ConvertFrom-Json
$clientId     = $creds.installed.client_id
$clientSecret = $creds.installed.client_secret
$port         = 54825
$redirectUri  = "http://localhost:$port"

$scopes = @(
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.scriptapp",
  "https://www.googleapis.com/auth/script.deployments",
  "https://www.googleapis.com/auth/script.projects",
  "https://www.googleapis.com/auth/script.webapp.deploy",
  "https://www.googleapis.com/auth/drive.metadata.readonly",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/service.management",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/logging.read",
  "https://www.googleapis.com/auth/cloud-platform"
) -join " "

$scopeEncoded = [Uri]::EscapeDataString($scopes)
$authUrl = "https://accounts.google.com/o/oauth2/v2/auth?client_id=$clientId&redirect_uri=$([Uri]::EscapeDataString($redirectUri))&response_type=code&scope=$scopeEncoded&access_type=offline&prompt=consent"

Write-Host "Opening browser for OAuth approval..."
Start-Process $authUrl

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("$redirectUri/")
$listener.Start()
Write-Host "Waiting for OAuth redirect on port $port..."

$context  = $listener.GetContext()
$rawUrl   = $context.Request.RawUrl
$code     = ($rawUrl -split "[?&]" | Where-Object { $_ -like "code=*" }) -replace "code=", ""

$response = $context.Response
$html     = "<html><body><h2>Authentification réussie! Vous pouvez fermer cette fenêtre.</h2></body></html>"
$bytes    = [System.Text.Encoding]::UTF8.GetBytes($html)
$response.ContentLength64 = $bytes.Length
$response.OutputStream.Write($bytes, 0, $bytes.Length)
$response.Close()
$listener.Stop()

Write-Host "Code received. Exchanging for token..."

$body = @{
  code          = $code
  client_id     = $clientId
  client_secret = $clientSecret
  redirect_uri  = $redirectUri
  grant_type    = "authorization_code"
}

$tokenResponse = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $body
$expiryDate    = [DateTimeOffset]::UtcNow.AddSeconds($tokenResponse.expires_in).ToUnixTimeMilliseconds()

$clasprc = @{
  tokens = @{
    default = @{
      client_id     = $clientId
      client_secret = $clientSecret
      type          = "authorized_user"
      refresh_token = $tokenResponse.refresh_token
      access_token  = $tokenResponse.access_token
      token_type    = "Bearer"
      expiry_date   = $expiryDate
    }
  }
}

$json = $clasprc | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText("$env:USERPROFILE\.clasprc.json", $json, (New-Object System.Text.UTF8Encoding $false))
Write-Host "Token saved to ~/.clasprc.json. clasp run should now work!"
