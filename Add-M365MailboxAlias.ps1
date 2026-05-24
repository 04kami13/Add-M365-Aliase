[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0)]
    [string]$UserEmail,

    [Parameter(Position = 1)]
    [string]$TargetDomain,

    [string]$InputFilePath,
    [string]$WorksheetName,
    [string]$EmailColumn
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:EmailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
$script:ShouldProcessInvoker = $PSCmdlet

function Read-RequiredValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    do {
        $value = Read-Host -Prompt $Prompt
    } until (-not [string]::IsNullOrWhiteSpace($value))

    return $value.Trim()
}

function Select-Choice {
    param(
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [object[]]$Choices
    )

    if (-not $Choices -or $Choices.Count -eq 0) {
        throw "No choices are available."
    }

    if ($Choices.Count -eq 1) {
        return $Choices[0].Value
    }

    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        Write-Host ""
        Write-Host $Title -ForegroundColor Cyan
    }

    for ($index = 0; $index -lt $Choices.Count; $index++) {
        Write-Host ("[{0}] {1}" -f ($index + 1), $Choices[$index].Label)
    }

    $selectedIndex = 0
    do {
        $choice = Read-Host -Prompt $Prompt
        $isValidNumber = [int]::TryParse($choice, [ref]$selectedIndex)

        if ($isValidNumber -and $selectedIndex -ge 1 -and $selectedIndex -le $Choices.Count) {
            return $Choices[$selectedIndex - 1].Value
        }

        Write-Host "Please choose a number between 1 and $($Choices.Count)." -ForegroundColor Yellow
    } while ($true)
}

function Select-ProcessingMode {
    $choices = @(
        [PSCustomObject]@{ Label = 'Process one mailbox'; Value = 'Single' }
        [PSCustomObject]@{ Label = 'Process mailbox list from Excel or CSV'; Value = 'Bulk' }
    )

    return Select-Choice -Title 'Choose how to run the script:' -Prompt 'Choose mode number' -Choices $choices
}

function Test-LooksLikeEmail {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return $Value.Trim() -match $script:EmailPattern
}

function ConvertTo-NormalizedEmail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )

    $trimmedEmail = $EmailAddress.Trim()

    try {
        $mailAddress = [System.Net.Mail.MailAddress]::new($trimmedEmail)
    }
    catch {
        throw "Invalid email address: '$EmailAddress'."
    }

    if ($mailAddress.Address -notmatch $script:EmailPattern) {
        throw "Invalid email address: '$EmailAddress'."
    }

    return $mailAddress.Address.ToLowerInvariant()
}

function Get-EmailLocalPart {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )

    return ($EmailAddress -split '@', 2)[0]
}

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    }
    catch {
        throw "File not found: $Path"
    }
}

function Ensure-ExchangeOnlineModule {
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $module) {
        throw "ExchangeOnlineManagement module is not installed. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -MinimumVersion $module.Version -ErrorAction Stop | Out-Null
}

function Test-ExchangeOnlineConnection {
    $connectionCommand = Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $connectionCommand) {
        return $false
    }

    $connections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
    if (-not $connections) {
        return $false
    }

    foreach ($connection in $connections) {
        $state = $null
        $connectionUri = $null

        try {
            $state = $connection.State
            $connectionUri = [string]$connection.ConnectionUri
        }
        catch {
            continue
        }

        if ($state -eq 'Connected' -and $connectionUri -match 'outlook\.office365\.com') {
            return $true
        }
    }

    return $false
}

function Ensure-ExchangeOnlineConnection {
    if (Test-ExchangeOnlineConnection) {
        return
    }

    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

function Select-TenantDomain {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Domains
    )

    $choices = foreach ($domain in $Domains) {
        [PSCustomObject]@{
            Label = $domain
            Value = $domain
        }
    }

    return Select-Choice -Title 'Accepted domains in this tenant:' -Prompt 'Choose domain number' -Choices $choices
}

function Get-SmtpAddresses {
    param(
        [Parameter(Mandatory = $true)]
        $Mailbox
    )

    return @(
        $Mailbox.EmailAddresses |
            ForEach-Object { $_.ToString() } |
            Where-Object { $_ -match '^(smtp|SMTP):' } |
            ForEach-Object { $_.Substring(5).ToLowerInvariant() } |
            Sort-Object -Unique
    )
}

function Assert-EmailAddressUpdateSupported {
    param(
        [Parameter(Mandatory = $true)]
        $Mailbox
    )

    $setMailboxCommand = Get-Command -Name Set-Mailbox -ErrorAction Stop
    if ($setMailboxCommand.Parameters.ContainsKey('EmailAddresses')) {
        return
    }

    $mailboxType = $null

    try {
        $mailboxType = [string]$Mailbox.RecipientTypeDetails
    }
    catch {
    }

    $message = "This Exchange session does not expose Set-Mailbox -EmailAddresses."

    if (-not [string]::IsNullOrWhiteSpace($mailboxType)) {
        $message += " Mailbox type: $mailboxType."
    }

    $message += " This usually means the signed-in account does not have permission to manage mailbox email addresses, or the mailbox is synchronized from on-premises/hybrid and the alias must be added on-premises."

    throw $message
}

function Get-UniqueColumnNames {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Values
    )

    $seen = @{}
    $headers = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $Values.Count; $index++) {
        $header = $Values[$index]

        if ([string]::IsNullOrWhiteSpace($header)) {
            $header = "Column$($index + 1)"
        }
        else {
            $header = $header.Trim()
        }

        $baseHeader = $header
        $suffix = 2

        while ($seen.ContainsKey($header.ToLowerInvariant())) {
            $header = "{0}_{1}" -f $baseHeader, $suffix
            $suffix++
        }

        $seen[$header.ToLowerInvariant()] = $true
        $headers.Add($header)
    }

    return $headers.ToArray()
}

function Test-WorksheetHasHeader {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FirstRow,

        [string[]]$SecondRow
    )

    $headerNames = @(
        'email',
        'emailaddress',
        'email address',
        'mail',
        'mailbox',
        'useremail',
        'user email',
        'userprincipalname',
        'upn',
        'primarysmtpaddress',
        'primary smtp address'
    )

    $nonEmptyFirstRow = @($FirstRow | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not $nonEmptyFirstRow) {
        return $false
    }

    foreach ($value in $nonEmptyFirstRow) {
        if ($headerNames -contains $value.Trim().ToLowerInvariant()) {
            return $true
        }
    }

    foreach ($value in $nonEmptyFirstRow) {
        if (Test-LooksLikeEmail -Value $value) {
            return $false
        }
    }

    if ($SecondRow) {
        foreach ($value in $SecondRow) {
            if (Test-LooksLikeEmail -Value $value) {
                return $true
            }
        }
    }

    return $true
}

function New-RowObject {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Headers,

        [Parameter(Mandatory = $true)]
        [string[]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$SourceRow
    )

    $properties = [ordered]@{}

    for ($index = 0; $index -lt $Headers.Count; $index++) {
        $properties[$Headers[$index]] = $Values[$index]
    }

    $properties['SourceRow'] = $SourceRow

    return [PSCustomObject]$properties
}

function Convert-TableToObjects {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[object[]]]$Rows
    )

    if ($Rows.Count -eq 0) {
        throw "The file does not contain any rows."
    }

    $firstRow = @($Rows[0] | ForEach-Object { [string]$_ })
    $secondRow = $null

    if ($Rows.Count -ge 2) {
        $secondRow = @($Rows[1] | ForEach-Object { [string]$_ })
    }

    $hasHeader = Test-WorksheetHasHeader -FirstRow $firstRow -SecondRow $secondRow
    if ($hasHeader) {
        $headers = Get-UniqueColumnNames -Values $firstRow
        $startIndex = 1
    }
    else {
        $headers = 1..$firstRow.Count | ForEach-Object { "Column$_" }
        $startIndex = 0
    }

    $objects = New-Object System.Collections.Generic.List[object]

    for ($index = $startIndex; $index -lt $Rows.Count; $index++) {
        $values = @($Rows[$index] | ForEach-Object {
                if ($null -eq $_) {
                    ''
                }
                else {
                    [string]$_
                }
            })

        $hasData = $false
        foreach ($value in $values) {
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $hasData = $true
                break
            }
        }

        if (-not $hasData) {
            continue
        }

        $sourceRow = $index + 1
        $objects.Add((New-RowObject -Headers $headers -Values $values -SourceRow $sourceRow))
    }

    return $objects.ToArray()
}

function Get-ExcelWorksheetRows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string]$WorksheetName
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($Path, 0, $true)

        $worksheetNames = New-Object System.Collections.Generic.List[string]
        for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
            $worksheetNames.Add([string]$workbook.Worksheets.Item($index).Name)
        }

        if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
            $worksheetChoices = foreach ($name in $worksheetNames) {
                [PSCustomObject]@{
                    Label = $name
                    Value = $name
                }
            }

            $WorksheetName = Select-Choice -Title 'Worksheets in the Excel file:' -Prompt 'Choose worksheet number' -Choices $worksheetChoices
        }
        elseif ($worksheetNames -notcontains $WorksheetName) {
            throw "Worksheet '$WorksheetName' was not found in file $Path"
        }

        $worksheet = $workbook.Worksheets.Item($WorksheetName)
        $usedRange = $worksheet.UsedRange

        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count

        if ($rowCount -lt 1 -or $columnCount -lt 1) {
            throw "Worksheet '$WorksheetName' is empty."
        }

        $rows = New-Object 'System.Collections.Generic.List[object[]]'

        for ($row = 1; $row -le $rowCount; $row++) {
            $values = New-Object string[] $columnCount

            for ($column = 1; $column -le $columnCount; $column++) {
                $cellText = [string]$worksheet.Cells.Item($row, $column).Text
                $values[$column - 1] = $cellText.Trim()
            }

            $rows.Add($values)
        }

        return Convert-TableToObjects -Rows $rows
    }
    catch {
        $message = $_.Exception.Message

        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = $_.ToString()
        }

        if (-not $excel) {
            throw "Excel desktop could not be opened to read '$Path'. Install Excel or save the file as CSV and use that file instead."
        }

        throw $message
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
        }

        if ($excel) {
            $excel.Quit()
        }

        foreach ($comObject in @($usedRange, $worksheet, $workbook, $excel)) {
            if ($comObject) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-DelimitedFileRows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Delimiter
    )

    Add-Type -AssemblyName Microsoft.VisualBasic

    $parser = $null

    try {
        $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters($Delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true

        $rows = New-Object 'System.Collections.Generic.List[object[]]'

        while (-not $parser.EndOfData) {
            $fields = @($parser.ReadFields())
            $rows.Add($fields)
        }

        return Convert-TableToObjects -Rows $rows
    }
    finally {
        if ($parser) {
            $parser.Close()
        }
    }
}

function Get-InputRowsFromFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string]$WorksheetName
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    switch ($extension) {
        '.xlsx' { return Get-ExcelWorksheetRows -Path $Path -WorksheetName $WorksheetName }
        '.xlsm' { return Get-ExcelWorksheetRows -Path $Path -WorksheetName $WorksheetName }
        '.xls' { return Get-ExcelWorksheetRows -Path $Path -WorksheetName $WorksheetName }
        '.csv' { return Get-DelimitedFileRows -Path $Path -Delimiter ',' }
        '.tsv' { return Get-DelimitedFileRows -Path $Path -Delimiter "`t" }
        default { throw "Unsupported file type '$extension'. Use .xlsx, .xls, .xlsm, .csv, or .tsv." }
    }
}

function Get-FirstNonEmptyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$ColumnName
    )

    foreach ($row in $Rows) {
        $value = [string]$row.PSObject.Properties[$ColumnName].Value
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    return $null
}

function Resolve-ColumnName {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [string]$RequestedColumn
    )

    $propertyNames = @(
        $Rows[0].PSObject.Properties.Name |
            Where-Object { $_ -ne 'SourceRow' }
    )

    if (-not $propertyNames) {
        throw "No columns were found in the input file."
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedColumn)) {
        foreach ($propertyName in $propertyNames) {
            if ($propertyName.ToLowerInvariant() -eq $RequestedColumn.Trim().ToLowerInvariant()) {
                return $propertyName
            }
        }

        throw "Column '$RequestedColumn' was not found in the input file."
    }

    $preferredNames = @(
        'email',
        'emailaddress',
        'email address',
        'mail',
        'mailbox',
        'useremail',
        'user email',
        'userprincipalname',
        'upn',
        'primarysmtpaddress',
        'primary smtp address'
    )

    $preferredMatches = @(
        $propertyNames |
            Where-Object { $preferredNames -contains $_.ToLowerInvariant() }
    )

    if ($preferredMatches.Count -eq 1) {
        return $preferredMatches[0]
    }

    $emailLikeColumns = New-Object System.Collections.Generic.List[string]
    foreach ($propertyName in $propertyNames) {
        $sample = Get-FirstNonEmptyValue -Rows $Rows -ColumnName $propertyName
        if (Test-LooksLikeEmail -Value $sample) {
            $emailLikeColumns.Add($propertyName)
        }
    }

    if ($emailLikeColumns.Count -eq 1) {
        return $emailLikeColumns[0]
    }

    if ($propertyNames.Count -eq 1) {
        return $propertyNames[0]
    }

    $columnChoices = foreach ($propertyName in $propertyNames) {
        $sample = Get-FirstNonEmptyValue -Rows $Rows -ColumnName $propertyName
        $label = $propertyName

        if (-not [string]::IsNullOrWhiteSpace($sample)) {
            $label = "{0} (sample: {1})" -f $propertyName, $sample
        }

        [PSCustomObject]@{
            Label = $label
            Value = $propertyName
        }
    }

    return Select-Choice -Title 'Columns found in the file:' -Prompt 'Choose email column number' -Choices $columnChoices
}

function Get-MailboxInputs {
    param(
        [string]$UserEmail,
        [string]$InputFilePath,
        [string]$WorksheetName,
        [string]$EmailColumn
    )

    if (-not [string]::IsNullOrWhiteSpace($UserEmail) -and -not [string]::IsNullOrWhiteSpace($InputFilePath)) {
        throw "Use either UserEmail or InputFilePath, not both."
    }

    $mode = $null

    if (-not [string]::IsNullOrWhiteSpace($UserEmail)) {
        $mode = 'Single'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($InputFilePath)) {
        $mode = 'Bulk'
    }
    else {
        $mode = Select-ProcessingMode
    }

    if ($mode -eq 'Single') {
        if ([string]::IsNullOrWhiteSpace($UserEmail)) {
            $UserEmail = Read-RequiredValue -Prompt 'Enter the mailbox email address'
        }

        return @(
            [PSCustomObject]@{
                SourceRow = $null
                InputEmail = $UserEmail.Trim()
            }
        )
    }

    if ([string]::IsNullOrWhiteSpace($InputFilePath)) {
        $InputFilePath = Read-RequiredValue -Prompt 'Enter the Excel or CSV file path'
    }

    $resolvedPath = Resolve-ExistingPath -Path $InputFilePath
    $rows = @(Get-InputRowsFromFile -Path $resolvedPath -WorksheetName $WorksheetName)

    if (-not $rows) {
        throw "No data rows were found in file $resolvedPath"
    }

    $resolvedColumn = Resolve-ColumnName -Rows $rows -RequestedColumn $EmailColumn

    $mailboxInputs = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $inputEmail = [string]$row.PSObject.Properties[$resolvedColumn].Value

        if ([string]::IsNullOrWhiteSpace($inputEmail)) {
            continue
        }

        $mailboxInputs.Add([PSCustomObject]@{
                SourceRow = [int]$row.SourceRow
                InputEmail = $inputEmail.Trim()
            })
    }

    if ($mailboxInputs.Count -eq 0) {
        throw "No mailbox email addresses were found in column '$resolvedColumn'."
    }

    Write-Host ("Loaded {0} mailbox value(s) from {1} using column '{2}'." -f $mailboxInputs.Count, $resolvedPath, $resolvedColumn) -ForegroundColor Cyan

    return $mailboxInputs.ToArray()
}

function Invoke-AliasUpdate {
    param(
        [Parameter(Mandatory = $true)]
        $MailboxInput,

        [Parameter(Mandatory = $true)]
        [string]$TargetDomain
    )

    $inputEmail = [string]$MailboxInput.InputEmail
    $sourceRow = $MailboxInput.SourceRow
    $normalizedUserEmail = $null
    $newAlias = $null

    try {
        $normalizedUserEmail = ConvertTo-NormalizedEmail -EmailAddress $inputEmail
        $localPart = Get-EmailLocalPart -EmailAddress $normalizedUserEmail
        $newAlias = ("{0}@{1}" -f $localPart, $TargetDomain).ToLowerInvariant()

        Write-Host ("Checking mailbox {0} ..." -f $normalizedUserEmail) -ForegroundColor Cyan
        $mailbox = Get-Mailbox -Identity $normalizedUserEmail -ErrorAction Stop
        $existingAddresses = Get-SmtpAddresses -Mailbox $mailbox

        $status = $null
        $message = $null

        if ($existingAddresses -contains $newAlias) {
            $status = 'AlreadyExists'
            $message = "Alias already exists: $newAlias"
            Write-Host $message -ForegroundColor Yellow
        }
        else {
            Assert-EmailAddressUpdateSupported -Mailbox $mailbox

            if ($script:ShouldProcessInvoker.ShouldProcess($mailbox.Identity.ToString(), "Add secondary alias $newAlias")) {
                Write-Host ("Adding alias {0} ..." -f $newAlias) -ForegroundColor Cyan
                Set-Mailbox -Identity $mailbox.Identity -EmailAddresses @{ Add = $newAlias }
                $status = 'PendingVerification'
                $message = "Alias update command completed: $newAlias"
            }
            else {
                $status = 'WhatIf'
                $message = "WhatIf: alias not added: $newAlias"
            }
        }

        $updatedMailbox = Get-Mailbox -Identity $mailbox.Identity -ErrorAction Stop
        $updatedAddresses = Get-SmtpAddresses -Mailbox $updatedMailbox

        if ($status -eq 'PendingVerification') {
            if ($updatedAddresses -contains $newAlias) {
                $status = 'Added'
                $message = "Alias added and verified: $newAlias"
            }
            else {
                $status = 'NotVerified'
                $message = "Set-Mailbox completed, but Get-Mailbox does not show $newAlias yet."
                Write-Host $message -ForegroundColor Yellow
            }
        }

        return [PSCustomObject]@{
            SourceRow   = $sourceRow
            User        = $normalizedUserEmail
            Email       = $updatedMailbox.PrimarySmtpAddress.ToString().ToLowerInvariant()
            AddedAlias  = $newAlias
            Status      = $status
            Message     = $message
            Addresses   = ($updatedAddresses -join '; ')
        }
    }
    catch {
        $message = $_.Exception.Message

        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = $_.ToString()
        }

        Write-Host ("Error for {0}: {1}" -f $inputEmail, $message) -ForegroundColor Red

        return [PSCustomObject]@{
            SourceRow   = $sourceRow
            User        = $inputEmail
            Email       = $normalizedUserEmail
            AddedAlias  = $newAlias
            Status      = 'Error'
            Message     = $message
            Addresses   = ''
        }
    }
}

try {
    $mailboxInputs = @(Get-MailboxInputs -UserEmail $UserEmail -InputFilePath $InputFilePath -WorksheetName $WorksheetName -EmailColumn $EmailColumn)

    Ensure-ExchangeOnlineModule
    Ensure-ExchangeOnlineConnection

    $acceptedDomains = @(
        Get-AcceptedDomain |
            ForEach-Object { $_.DomainName.ToString().ToLowerInvariant() } |
            Sort-Object -Unique
    )

    if (-not $acceptedDomains) {
        throw "No accepted domains were found in this tenant."
    }

    if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
        $TargetDomain = Select-TenantDomain -Domains $acceptedDomains
    }
    else {
        $TargetDomain = $TargetDomain.Trim().ToLowerInvariant()
        if ($acceptedDomains -notcontains $TargetDomain) {
            throw "Domain '$TargetDomain' is not an accepted domain in this tenant."
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($mailboxInput in $mailboxInputs) {
        $results.Add((Invoke-AliasUpdate -MailboxInput $mailboxInput -TargetDomain $TargetDomain))
    }

    $addedCount = @($results | Where-Object { $_.Status -eq 'Added' }).Count
    $alreadyExistsCount = @($results | Where-Object { $_.Status -eq 'AlreadyExists' }).Count
    $errorCount = @($results | Where-Object { $_.Status -eq 'Error' }).Count
    $whatIfCount = @($results | Where-Object { $_.Status -eq 'WhatIf' }).Count
    $notVerifiedCount = @($results | Where-Object { $_.Status -eq 'NotVerified' }).Count

    Write-Host ""
    Write-Host ("Processed: {0}" -f $results.Count) -ForegroundColor Green
    Write-Host ("Added: {0}" -f $addedCount) -ForegroundColor Green
    Write-Host ("Already exists: {0}" -f $alreadyExistsCount) -ForegroundColor Green
    Write-Host ("Errors: {0}" -f $errorCount) -ForegroundColor Green

    if ($notVerifiedCount -gt 0) {
        Write-Host ("Not verified yet: {0}" -f $notVerifiedCount) -ForegroundColor Yellow
    }

    if ($whatIfCount -gt 0) {
        Write-Host ("WhatIf only: {0}" -f $whatIfCount) -ForegroundColor Green
    }

    Write-Host ""
    $results |
        Select-Object SourceRow, User, Email, AddedAlias, Status, Message, Addresses
}
catch {
    $message = $_.Exception.Message

    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = $_.ToString()
    }

    Write-Host ""
    Write-Host ("Error: {0}" -f $message) -ForegroundColor Red
    exit 1
}
