#Send request on Discord with content and username

Set-StrictMode -Version Latest

$webHookUrl = '{Your token Discord}'

$username = 'Bot'

$content = @'
Example
- Example 1
- Example 2
'@

$payload = [PSCustomObject]@{

    username = $username
    content = $content
}

Invoke-RestMethod -Uri $webHookUrl -Method Post -Body ($payload | ConvertTo-Json -Depth 4) -ContentType 'Application/Json'