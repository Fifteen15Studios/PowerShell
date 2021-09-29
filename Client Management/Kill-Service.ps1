$computerName = Read-Host "Computer Name. (Leave blank for local machine)"
$Service = Read-Host "What service do you want to kill?"

# Use local computer if left blank
if(!$computerName) {
    $computerName = $env:COMPUTERNAME
}

# Query running services for the service in question
$Services = cmd /c "sc \\$ComputerName queryex $Service"

# Find the line with the service's PID
foreach($line in $Services) {
    if($line.Contains("PID")) {
        break
    }
}

# Find the ID
$ID = $line.Substring($line.IndexOf(":") + 2)

# If ID was found
if($ID) {
    $KillOutput = cmd /c "taskkill /S $ComputerName /PID $ID /F"

    # If it was successful, display success
    if($KillOutput.StartsWith("SUCCESS")) {
        "Successfully killed the service"
    }
    # If it wasn't successful, display output of the command
    else {
        $killOutput
    }
}
else {
    "$Service not found"
}