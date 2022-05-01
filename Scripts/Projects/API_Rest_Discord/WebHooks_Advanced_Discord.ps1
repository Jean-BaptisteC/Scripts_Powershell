#Send request on Discord with content, color, title and username

Set-StrictMode -Version Latest

$webHookUrl = '{Your token Discord}'

[System.Collections.ArrayList]$embedArray = @()

$embedObject = [PSCustomObject]@{
    username = 'Bot'
    color = '1'
    title = 'Example'
    description = ('
    - Example 1
    - Example 2')
}

$embedArray.Add($embedObject)|Out-Null

$payload = [PSCustomObject]@{
    embeds = $embedObject

}

Invoke-RestMethod -Uri $webHookUrl -Body ($payload | ConvertTo-Json -Depth 4) -Method Post -ContentType 'application/json'