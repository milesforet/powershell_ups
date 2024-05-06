$client_id = Get-Content -Path ./auth/client_id.txt
$client_secret = Get-Content -Path ./auth/client_secret.txt 

$url = "https://onlinetools.ups.com/security/v1/oauth/token"

$creds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $client_id,$client_secret)))

$body= @{
    "grant_type"= "client_credentials"
}

$headers= @{
    'Content-Type' = 'application/x-www-form-urlencoded'
    'x-merchant-id' = 'string'
    'Authorization' = 'Basic '+$creds
}

$result = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body

$result.access_token | Out-File -FilePath ./auth/bearer.txt