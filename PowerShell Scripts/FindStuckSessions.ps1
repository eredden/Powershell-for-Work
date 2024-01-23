# Powershell wrapper for qwinsta, ripped from StackOverflow.
# https://stackoverflow.com/questions/23445175/qwinsta-serversomesrv-equivalent-in-powershell
function Get-TSSessions {
    param(
        $ComputerName = 'localhost'
    )

    $output = qwinsta /server:$ComputerName
    if ($null -eq $output) {
        return
    }

    # Get column names and locations from fixed-width header
    $columns = [regex]::Matches($output[0],'(?<=\s)\w+')
    $output | Select-Object -Skip 1 | Foreach-Object {
        [string]$line = $_

        $session = [ordered]@{}
        for ($i=0; $i -lt $columns.Count; $i++) {
            $currentColumn = $columns[$i]
            $columnName = $currentColumn.Value

            if ($i -eq $columns.Count-1) {
                # Last column, get rest of the line
                $columnValue = $line.Substring($currentColumn.Index).Trim()
            } else {
                $lengthToNextColumn = $columns[$i+1].Index - $currentColumn.Index
                $columnValue = $line.Substring($currentColumn.Index, $lengthToNextColumn).Trim()
            }

            $session.$columnName = $columnValue.Trim()
        }

        [pscustomobject]$session
    }
}

# Find and print stuck sessions from list of sessions.
# Takes the output of QWINSTA (a.k.a Get-TSSessions) as an input.
function Get-StuckSessions {
    param(
        $Sessions
    )

    $StuckSessions = $Sessions | Where-Object State -eq "Disc" `
        | Where-Object Username -eq "" `
        | Where-Object SessionName -eq ""

    if ($StuckSessions) {
        foreach ($StuckSession in $StuckSessions) { 
			Write-Host "Session ID " -ForegroundColor Cyan -NoNewline
            Write-Host $StuckSession.Id -ForegroundColor Yellow -NoNewline
			Write-Host " is stuck!" -ForegroundColor Cyan
        }
    }

    else {
        Write-Host "$Server " -ForegroundColor Yellow -NoNewline
        Write-Host "has no stuck sessions." -ForegroundColor Cyan
    }

    Write-Host "`n" -NoNewline
}

while ($true) {
    Clear-Host

    # Create server list.
    $ShortCode = Read-Host "Enter the server prefix (ex. WESTSERV)"
    $ServerList = (Get-ADComputer -Filter "name -like '$ShortCode-*'").name

    Clear-Host

    # Get sessions from servers and find the stuck ones.
    foreach ($Server in $ServerList) {
        Write-Host "Checking for stuck sessions on server: $Server"

        if (Test-Connection -ComputerName $server -Count 1 -Quiet) {
            $Sessions = Get-TSSessions -ComputerName $Server
            Get-StuckSessions $Sessions
        }
    
        else {
            Write-Host "`n$Server " -ForegroundColor Yellow -NoNewline
            Write-Host "is unreachable." -ForegroundColor Cyan
        }
    }

    if ($null -eq $ServerList) {
        Write-Host "No servers with that server prefix could be found.`n"
    }
 
    Read-Host "Press ENTER to continue, or CTRL + C to exit"
}
