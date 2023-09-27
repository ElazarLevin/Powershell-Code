$WebResponse = Invoke-WebRequest -URI 'https://jsonplaceholder.typicode.com/todos'
$posts = $WebResponse | ConvertFrom-Json
$posts | FT
$JSONText = $posts | ConvertTo-Json
Write-Host $JSONText