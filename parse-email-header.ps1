$EmailFile = "C:\Users\jadan\Scripts\email.eml"

# Read the contents of the email file
$EmailContent = Get-Content $EmailFile

# Split the email content into an array of lines
$Lines = $EmailContent -split "`n"

# Create a hashtable to store the header fields
$Header = @{}

# Set the colors for the output
$ColorNormal = [ConsoleColor]::Gray
$ColorHeader = [consoleColor]::Cyan

# Loop through the lines of the email
foreach ($Line in $Lines) {
    # Check if the line is a header field
    if ($Line -match "^(?<key>.+?):\s(?<value>.+)$") {
        # If the line is a header field, add it to the hashtable
        $Header[$Matches['key']] = $Matches['value']
    }
    # Stop processing lines once we reach the end of the header
    else {
        break
    }
}

# Output the header data
Write-Host "Header Data:" -ForegroundColor $ColorNormal
foreach ($Key in $Header.Keys) {
    Write-Host "  ${Key}: $($Header[$Key])" -ForegroundColor $ColorHeader
}

# Output additional data
Write-Host "Additional Data:" -ForegroundColor $ColorNormal

# Return-Path
if ($Header['Return-Path']) {
    Write-Host "  Return-Path: $($Header['Return-Path'])" -ForegroundColor $ColorHeader
}

# Received lines
$receivedLines = @()
foreach ($Key in $Header.Keys | Where-Object {$_ -match "^Received"}) {
    $receivedLines += $Header[$Key]
}

if ($receivedLines) {
    Write-Host "  Received lines:" -ForegroundColor $ColorHeader
    foreach ($receivedLine in $receivedLines) {
        Write-Host "    $receivedLine" -ForegroundColor $ColorValue
    }
}

# Extract the "Received" lines
$receivedLines = ($Header | Select-String -Pattern "^Received:").Line
Write-Host "Received lines:"
foreach ($receivedLine in $receivedLines) {
    Write-Host "  $receivedLine" -ForegroundColor $ColorHeader
}


# Authentication results
if ($Header['Authentication-Results']) {
    Write-Host "  Authentication results:" -ForegroundColor $ColorHeader
    Write-Host "    $($Header['Authentication-Results'])" -ForegroundColor $ColorValue
}


# DKIM signature
if ($Header['DKIM-Signature']) {
    Write-Host "  DKIM signature:" -ForegroundColor $ColorHeader
    Write-Host "    $($Header['DKIM-Signature'])" -ForegroundColor $ColorValue
}


# SPF record
if ($Header['Received-SPF']) {
    $spfRecord = ($Header['Received-SPF'] -split ' ' | Where-Object {$_ -match '^spf='}).Split('=')[1]
    Write-Host "  SPF record: $spfRecord" -ForegroundColor $ColorHeader
}


# DMARC policy
if ($Header['Authentication-Results']) {
    $dmarcPolicy = ($Header['Authentication-Results'] -split ";") |
        Where-Object {$_ -match "dmarc=.+"}
    if ($dmarcPolicy) {
        $dmarcPolicy = $dmarcPolicy -split "="
        Write-Host "  DMARC policy: $($dmarcPolicy[1])" -ForegroundColor $ColorHeader
    }
}


# X-Forwarded-For
if ($Header['X-Forwarded-For']) {
    Write-Host "  X-Forwarded-For: $($Header['X-Forwarded-For'])" -ForegroundColor $ColorHeader
}

# Sender's IP address
if ($Header['Received']) {
    Write-Host "  Sender's IP address:" -ForegroundColor $ColorHeader
    $ReceivedLines = $Header['Received'] -split "; "
    foreach ($ReceivedLine in $ReceivedLines) {
        Write-Host "    Received line: $ReceivedLine" -ForegroundColor $ColorHeader
        if ($ReceivedLine -match "^from\s(?<ip>[^\s]+?)\s") {
            Write-Host "      Matched IP: $($Matches['ip'])" -ForegroundColor $ColorHeader
        }
    }
}

# From address
if ($Header['From']) {
    Write-Host "  From: $($Header['From'])" -ForegroundColor $ColorHeader
}

# To address
if ($Header['To']) {
    Write-Host "  To: $($Header['To'])" -ForegroundColor $ColorHeader
}

# Subject
if ($Header['Subject']) {
    Write-Host "  Subject: $($Header['Subject'])" -ForegroundColor $ColorHeader
}

# Date
if ($Header['Date']) {
    Write-Host "  Date: $($Header['Date'])" -ForegroundColor $ColorHeader
}

# Message-ID
if ($Header['Message-ID']) {
    Write-Host "  Message-ID: $($Header['Message-ID'])" -ForegroundColor $ColorHeader
}

# In-Reply-To
if ($Header['In-Reply-To']) {
    Write-Host "  In-Reply-To: $($Header['In-Reply-To'])" -ForegroundColor $ColorHeader
}

# DKIM signature
if ($Header['DKIM-Signature']) {
    Write-Host "  DKIM signature: $($Header['DKIM-Signature'])" -ForegroundColor $ColorHeader
}

# Authentication results
if ($Header['Authentication-Results']) {
    Write-Host "  Authentication results: $($Header['Authentication-Results'])" -ForegroundColor $ColorHeader
}

# References
if ($Header['References']) {
    Write-Host "  References: $($Header['References'])" -ForegroundColor $ColorHeader
}

# MIME-Version
if ($Header['MIME-Version']) {
    Write-Host "  MIME-Version: $($Header['MIME-Version'])" -ForegroundColor $ColorHeader
}

# Content-Type
if ($Header['Content-Type']) {
    Write-Host "  Content-Type: $($Header['Content-Type'])" -ForegroundColor $ColorHeader
}

# Content-Transfer-Encoding
if ($Header['Content-Transfer-Encoding']) {
    Write-Host "  Content-Transfer-Encoding: $($Header['Content-Transfer-Encoding'])" -ForegroundColor $ColorHeader
}

# X-* (custom headers)
Write-Host "  Custom Headers:" -ForegroundColor $ColorHeader
foreach ($Key in $Header.Keys) {
    if ($Key -match "^X-") {
        Write-Host "    ${Key}: $($Header[$Key])" -ForegroundColor $ColorHeader
    }
}

