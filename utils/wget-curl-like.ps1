# download file (wget)
$url = "url"
$output = "outputfile"
Invoke-WebRequest -Uri $url -OutFile $output

# make request (curl)
Invoke-WebRequest -Uri "http://localhost:9200" | Select-Object -Expand Content

## make rest api request example 1
$Cred = Get-Credential
$Url = "https://server.contoso.com:8089/services/search/jobs/export"
$Body = @{
    search = "search index=_internal | reverse | table index,host,source,sourcetype,_raw"
    output_mode = "csv"
    earliest_time = "-2d@d"
    latest_time = "-1d@d"
}
Invoke-RestMethod -Method 'Post' -Uri $url -Credential $Cred -Body $body -OutFile output.csv

## make rest api request example 2
$person = @{
    first='joe'
    lastname='doe'
}
$json = $person | ConvertTo-Json
$response = Invoke-RestMethod 'http://example.com/api/people/1' -Method Put -Body $json -ContentType 'application/json'


# GET with custom headers example
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("X-DATE", '9/29/2014')
$headers.Add("X-SIGNATURE", '234j123l4kl23j41l23k4j')
$headers.Add("X-API-KEY", 'testuser')

$response = Invoke-RestMethod 'http://example.com/api/people/1' -Headers $headers

# DELETE example
$response = Invoke-RestMethod 'http://example.com/api/people/1' -Method Delete