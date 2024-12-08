[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function Drink-Espresso {
    # Credits to https://den.dev/blog/caffeinate-windows/
    Write-Host "[info] Currently ordering a double shot of espresso..."

    $Signature = @"
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern void SetThreadExecutionState(uint esFlags);
"@

    $ES_SYSTEM_REQUIRED = [uint32]"0x00000001"
    $ES_CONTINUOUS = [uint32]"0x80000000"
    $ES_AWAYMODE_REQUIRED = [uint32]"0x00000040"

    $JobName = "DrinkALotOfEspresso"

    $BackgroundJob = Start-Job -Name $JobName -ScriptBlock {
        $STES = Add-Type -MemberDefinition $args[0] -Name System -Namespace Win32 -PassThru
        $STES::SetThreadExecutionState($args[1] -bor $args[2] -bor $args[3])
        while ($true) {
            Start-Sleep -s 15
        }
    } -ArgumentList $Signature, $ES_SYSTEM_REQUIRED, $ES_CONTINUOUS, $ES_AWAYMODE_REQUIRED


    return $BackgroundJob
}

function Get-Response($url, $headers) {
    try {
        $response = Invoke-WebRequest -Uri $url -Headers $headers
        # $response = Invoke-WebRequest $url
        $response.StatusCode
        Write-Output $response
        return [PSCustomObject]@{

            data    = $response
            error   = $false
            success = $true
        }
    }
    catch {
        if ($_.Exception.Response) {
            # If the Response object is present, get the status code
            $statusCode = $_.Exception.Response.StatusCode.Value__
            Write-Output "HTTP Error: $($statusCode)"
            $errorMsg = "HTTP Error Status Code: $statusCode"
        }
        else {
            # If no Response object, fallback to exception details
            Write-Output "Non-HTTP Error: $($_.Exception.Message)"
            $errorMsg = "Error: $($_.Exception.Message)"
        }

        return [PSCustomObject]@{
            data    = $false
            error   = $errorMsg
            success = $false
        }
    }
}



# Set the time interval in seconds
$interval = 10
$BackgroundJob = $null
# Start a loop to continuously check for network traffic
while ($true) {

    # Replace YOUR_IP, YOUR_PORT and YOUR_TOKEN
    $url = "127.0.0.1:8096/Sessions"
    # Set the API key in the request header
    $headers = @{ "X-Emby-Token" = "d4c40d1c32cf454c842d16df37f8081b" }

    # Send the request and get the response
    # $response = Invoke-WebRequest -Uri $url -Headers $headers
    $response = Get-Response -url $url -headers $headers
    if ($response.success -eq $false) {
        Write-Output $response.error
        continue
    }else{
        Write-Output $response.data.content
    }
    $response = $response.data
    # Parse the JSON response
    $sessions = $response.Content | ConvertFrom-Json
    $sessions = $sessions | ? { $_.NowPlayingItem -or $_.UserId } | ? { [datetime]::ParseExact(($_.LastActivityDate -replace '\.\d+Z', 'Z'), 'yyyy-MM-ddTHH:mm:ssZ', $null) -ge (Get-Date).AddMinutes(-15) }

    $hasSession = $false
    # Iterate over the list of sessions
    foreach ($session in $sessions) {
        # Check if a session is active
        if ($session.isActive -and $session.Client -ne "DLNA" -and $session.Client -ne "Home Assistant") {
            $hasSession = $true
            if ($BackgroundJob -eq $null) {
                Write-Output "$($session.DeviceName) Alive!!"
                $keepOn = $true
                $BackgroundJob = Drink-Espresso
            }
            break
        }
    }

    if (-not $hasSession -and $BackgroundJob -ne $null) {
        Stop-Job $BackgroundJob
        Remove-Job $BackgroundJob
        $BackgroundJob = $null
        Write-Output "Stop sleep prevention"
    }
    # Wait for 10 seconds before checking again
    Start-Sleep -Seconds $interval
}
