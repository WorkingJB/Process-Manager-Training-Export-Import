<#
.SYNOPSIS
    Training Unit Import/Export Manager for Process Manager

.DESCRIPTION
    This script allows users to export training units from Process Manager to CSV
    or import/create training units from a CSV file.

.EXAMPLE
    .\TrainingUnitManager.ps1
#>

[CmdletBinding()]
param()

# Global variables
$script:BaseUrl = ""
$script:TenantName = ""
$script:BearerToken = ""
$script:ScimApiKey = ""
$script:ScimBaseUrl = "https://api.promapp.com"

#region Helper Functions

function Write-ColorOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )

    switch ($Type) {
        "Info"    { Write-Host $Message -ForegroundColor Cyan }
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error"   { Write-Host $Message -ForegroundColor Red }
    }
}

function Get-TypeLabel {
    param(
        [Parameter(Mandatory=$true)]
        [int]$TypeValue
    )

    switch ($TypeValue) {
        1 { return "Course" }
        2 { return "Online Resource" }
        3 { return "Document" }
        6 { return "Face to Face" }
        default { return "Unknown ($TypeValue)" }
    }
}

function Get-TypeValue {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TypeLabel
    )

    switch ($TypeLabel.Trim()) {
        "Course" { return 1 }
        "Online Resource" { return 2 }
        "Document" { return 3 }
        "Face to Face" { return 6 }
        default {
            # Try to parse as integer for backward compatibility
            $intValue = 0
            if ([int]::TryParse($TypeLabel, [ref]$intValue)) {
                return $intValue
            }
            Write-ColorOutput "Warning: Unknown Type label '$TypeLabel', defaulting to 1 (Course)" -Type "Warning"
            return 1
        }
    }
}

function Get-AssessmentLabel {
    param(
        [Parameter(Mandatory=$true)]
        [int]$AssessmentValue
    )

    switch ($AssessmentValue) {
        0 { return "None" }
        1 { return "Self Sign Off" }
        2 { return "Supervisor Sign Off" }
        default { return "Unknown ($AssessmentValue)" }
    }
}

function Get-AssessmentValue {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AssessmentLabel
    )

    switch ($AssessmentLabel.Trim()) {
        "None" { return 0 }
        "Self Sign Off" { return 1 }
        "Supervisor Sign Off" { return 2 }
        default {
            # Try to parse as integer for backward compatibility
            $intValue = 0
            if ([int]::TryParse($AssessmentLabel, [ref]$intValue)) {
                return $intValue
            }
            Write-ColorOutput "Warning: Unknown Assessment label '$AssessmentLabel', defaulting to 0 (None)" -Type "Warning"
            return 0
        }
    }
}

function Get-UserInput {
    Write-ColorOutput "`n=== Training Unit Manager ===" -Type "Info"
    Write-ColorOutput "Please provide the following information:`n" -Type "Info"

    # Get Site URL
    do {
        $siteUrl = Read-Host "Enter Process Manager site URL (e.g., https://us.promapp.com/sitename)"
        if ($siteUrl -match '^https?://([^/]+)/(.+?)/?$') {
            $script:BaseUrl = "https://$($matches[1])"
            $script:TenantName = $matches[2]
            Write-ColorOutput "Extracted Base URL: $script:BaseUrl" -Type "Success"
            Write-ColorOutput "Extracted Tenant: $script:TenantName" -Type "Success"
            $validUrl = $true
        } else {
            Write-ColorOutput "Invalid URL format. Please use format: https://domain.com/tenant" -Type "Error"
            $validUrl = $false
        }
    } while (-not $validUrl)

    # Get Service Account Credentials
    Write-Host ""
    $script:Username = Read-Host "Enter service account username"
    $securePassword = Read-Host "Enter service account password" -AsSecureString
    $script:Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
    )

    # Get SCIM API Key
    $secureScimKey = Read-Host "Enter SCIM API key" -AsSecureString
    $script:ScimApiKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureScimKey)
    )

    # Get Action
    Write-Host ""
    Write-ColorOutput "Select an action:" -Type "Info"
    Write-Host "1. Export training units to CSV"
    Write-Host "2. Import training units from CSV"

    do {
        $action = Read-Host "Enter your choice (1 or 2)"
    } while ($action -notin @("1", "2"))

    return $action
}

function Get-BearerToken {
    Write-ColorOutput "`nAuthenticating..." -Type "Info"

    $tokenUrl = "$script:BaseUrl/$script:TenantName/oauth2/token"

    $body = @{
        grant_type = "password"
        username = $script:Username
        password = $script:Password
        duration = 60000
    }

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        $script:BearerToken = $response.access_token
        Write-ColorOutput "Authentication successful!" -Type "Success"
        return $true
    } catch {
        Write-ColorOutput "Authentication failed: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

function Invoke-ApiRequest {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Endpoint,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Get", "Post", "Put", "Delete")]
        [string]$Method = "Get",

        [Parameter(Mandatory=$false)]
        [hashtable]$Body = $null,

        [Parameter(Mandatory=$false)]
        [bool]$UseScimApi = $false
    )

    $baseUrl = if ($UseScimApi) { $script:ScimBaseUrl } else { $script:BaseUrl }
    $uri = "$baseUrl/$Endpoint"

    $headers = @{
        "Content-Type" = "application/json"
    }

    if ($UseScimApi) {
        $headers["Authorization"] = "Bearer $script:ScimApiKey"
    } else {
        $headers["Authorization"] = "Bearer $script:BearerToken"
    }

    try {
        if ($Body) {
            $jsonBody = $Body | ConvertTo-Json -Depth 10
            $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body $jsonBody
        } else {
            $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers
        }
        return $response
    } catch {
        Write-ColorOutput "API request failed: $uri - $($_.Exception.Message)" -Type "Error"
        return $null
    }
}

function Get-AllTrainingUnits {
    Write-ColorOutput "`nFetching all training units..." -Type "Info"

    $allUnits = @()
    $page = 1
    $pageSize = 200

    do {
        Write-ColorOutput "Fetching page $page..." -Type "Info"

        $endpoint = "$script:TenantName/Training/Register/ListPage?page=$page&pageSize=$pageSize&SearchCriteria=&ListFilter=0&TrainingDue=0&StatusFilter=0"
        $response = Invoke-ApiRequest -Endpoint $endpoint -Method Get

        if ($response -and $response.success) {
            $allUnits += $response.trainingUnits
            $hasNextPage = $response.paging.HasNextPage
            $page++
        } else {
            Write-ColorOutput "Failed to fetch training units on page $page" -Type "Error"
            break
        }
    } while ($hasNextPage)

    Write-ColorOutput "Total training units fetched: $($allUnits.Count)" -Type "Success"
    return $allUnits
}

function Get-TrainingUnitDetails {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UnitUniqueId
    )

    $endpoint = "$script:TenantName/Training/Unit/GetTrainingUnitDetails?unitUniqueId=$UnitUniqueId"
    $response = Invoke-ApiRequest -Endpoint $endpoint -Method Get

    if ($response -and $response.success) {
        return $response.trainingUnit
    }
    return $null
}

function Get-ProcessDetails {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ProcessUniqueId
    )

    $endpoint = "$script:TenantName/Api/v1/Processes/$ProcessUniqueId"
    $response = Invoke-ApiRequest -Endpoint $endpoint -Method Get

    if ($response) {
        return $response.processJson
    }
    return $null
}

function Get-TrainingUnitTrainees {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UnitUniqueId
    )

    $allTrainees = @()
    $page = 1
    $pageSize = 200

    do {
        $endpoint = "$script:TenantName/Training/Trainee?unitUniqueId=$UnitUniqueId&page=$page&pageSize=$pageSize"
        $response = Invoke-ApiRequest -Endpoint $endpoint -Method Get

        if ($response -and $response.success) {
            $allTrainees += $response.trainees
            $hasNextPage = $response.paging.HasNextPage
            $page++
        } else {
            break
        }
    } while ($hasNextPage)

    return $allTrainees
}

function Get-ScimUserById {
    param(
        [Parameter(Mandatory=$true)]
        [int]$UserId
    )

    try {
        # SCIM API endpoint to get user by ID
        $endpoint = "api/scim/users/$UserId"

        $response = Invoke-ApiRequest -Endpoint $endpoint -Method Get -UseScimApi $true

        if ($response -and $response.Resources -and $response.Resources.Count -gt 0) {
            # Return the userName from the first resource
            return $response.Resources[0].userName
        }

        return $null
    } catch {
        Write-ColorOutput "Warning: Failed to lookup SCIM user for UserId $UserId : $($_.Exception.Message)" -Type "Warning"
        return $null
    }
}

#endregion

#region Export Functions

function Export-TrainingUnits {
    Write-ColorOutput "`n=== Starting Export Process ===" -Type "Info"

    # Get all training units
    $allUnits = Get-AllTrainingUnits

    if ($allUnits.Count -eq 0) {
        Write-ColorOutput "No training units found to export." -Type "Warning"
        return
    }

    $exportData = @()
    $currentUnit = 0

    foreach ($unit in $allUnits) {
        $currentUnit++
        Write-ColorOutput "Processing unit $currentUnit of $($allUnits.Count): $($unit.Title)" -Type "Info"

        # Get full details
        $details = Get-TrainingUnitDetails -UnitUniqueId $unit.UniqueId

        if (-not $details) {
            Write-ColorOutput "Failed to get details for $($unit.Title), skipping..." -Type "Warning"
            continue
        }

        # Get linked process details
        $linkedProcessTitles = @()
        $linkedProcessUniqueIds = @()

        foreach ($process in $details.LinkedProcesses) {
            $linkedProcessTitles += $process.Title
            # Extract UniqueId from URL
            if ($process.Url -match 'uniqueId=([a-f0-9\-]+)') {
                $linkedProcessUniqueIds += $matches[1]
            }
        }

        # Get linked document titles
        $linkedDocTitles = @()
        foreach ($doc in $details.LinkedDocuments) {
            if ($doc.Title) {
                $linkedDocTitles += $doc.Title
            }
        }

        # Get trainees and lookup usernames from SCIM
        $trainees = Get-TrainingUnitTrainees -UnitUniqueId $unit.UniqueId
        $traineeUsernames = @()
        foreach ($trainee in $trainees) {
            if ($trainee.UserId) {
                $username = Get-ScimUserById -UserId $trainee.UserId
                if ($username) {
                    $traineeUsernames += $username
                } else {
                    Write-ColorOutput "Warning: Could not find SCIM username for UserId $($trainee.UserId) ($($trainee.UserFullName)), skipping..." -Type "Warning"
                }
            }
        }

        # Create export object
        $exportObject = [PSCustomObject]@{
            "Title" = $details.Title
            "Description" = $details.Description
            "Type" = (Get-TypeLabel -TypeValue $details.Type)
            "Assessment Label" = (Get-AssessmentLabel -AssessmentValue $details.AssessmentMethod)
            "Renew Cycle" = $details.RenewCycle
            "Provider" = $details.Provider
            "Linked Processes: Title" = ($linkedProcessTitles -join ";")
            "Linked Processes: uniqueId" = ($linkedProcessUniqueIds -join ";")
            "Linked Documents: Titles" = ($linkedDocTitles -join ";")
            "Trainees: Usernames" = ($traineeUsernames -join ";")
        }

        $exportData += $exportObject
    }

    # Export to CSV
    $timestamp = Get-Date -Format "yyyyMMdd"
    $fileName = "TrainingUnits_Export_$timestamp.csv"

    try {
        $exportData | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8
        Write-ColorOutput "`nExport completed successfully!" -Type "Success"
        Write-ColorOutput "File saved: $fileName" -Type "Success"
        Write-ColorOutput "Total units exported: $($exportData.Count)" -Type "Success"
    } catch {
        Write-ColorOutput "Failed to save CSV file: $($_.Exception.Message)" -Type "Error"
    }
}

#endregion

#region Import Functions

function Import-TrainingUnits {
    Write-ColorOutput "`n=== Starting Import Process ===" -Type "Info"

    # Get CSV file path
    $csvPath = Read-Host "Enter the path to the CSV file"

    if (-not (Test-Path $csvPath)) {
        Write-ColorOutput "File not found: $csvPath" -Type "Error"
        return
    }

    # Read CSV
    try {
        $csvData = Import-Csv -Path $csvPath
    } catch {
        Write-ColorOutput "Failed to read CSV file: $($_.Exception.Message)" -Type "Error"
        return
    }

    Write-ColorOutput "Found $($csvData.Count) rows to process" -Type "Info"

    $successCount = 0
    $failedRows = @()
    $currentRow = 0

    foreach ($row in $csvData) {
        $currentRow++
        Write-ColorOutput "`nProcessing row $currentRow of $($csvData.Count): $($row.Title)" -Type "Info"

        try {
            # Parse linked processes
            $linkedProcesses = @()
            if ($row.'Linked Processes: uniqueIds' -and $row.'Linked Processes: uniqueIds'.Trim() -ne "") {
                $processIds = $row.'Linked Processes: uniqueIds' -split ';'

                foreach ($procId in $processIds) {
                    if ($procId.Trim() -ne "") {
                        Write-ColorOutput "Looking up process: $($procId.Trim())" -Type "Info"
                        $processDetails = Get-ProcessDetails -ProcessUniqueId $procId.Trim()

                        if ($processDetails) {
                            $processUrl = "/$script:TenantName/Process?uniqueId=$($procId.Trim())"
                            $linkedProcesses += @{
                                Id = $processDetails.Id
                                Title = $processDetails.Name
                                Url = $processUrl
                            }
                        } else {
                            Write-ColorOutput "Warning: Could not find process with ID $($procId.Trim())" -Type "Warning"
                        }
                    }
                }
            }

            # Parse linked documents (just titles for now)
            $linkedDocuments = @()
            if ($row.'Linked Documents: Titles' -and $row.'Linked Documents: Titles'.Trim() -ne "") {
                $docTitles = $row.'Linked Documents: Titles' -split ';'
                foreach ($docTitle in $docTitles) {
                    if ($docTitle.Trim() -ne "") {
                        $linkedDocuments += $docTitle.Trim()
                    }
                }
            }

            # Build request body
            $requestBody = @{
                Id = 0  # 0 for new training unit
                Title = $row.Title
                Description = $row.Description
                Type = (Get-TypeValue -TypeLabel $row.Type)
                AssessmentMethod = (Get-AssessmentValue -AssessmentLabel $row.'Assessed Label')
                RenewCycle = [int]$row.'Renew Cycle'
                Provider = $row.Provider
                LinkedProcesses = $linkedProcesses
                LinkedDocuments = $linkedDocuments
                OtherResources = @()
                RolesRequiredToComplete = @()
                Tags = @()
            }

            # Create training unit
            $endpoint = "$script:TenantName/Training/Unit/EditTrainingUnit"
            $response = Invoke-ApiRequest -Endpoint $endpoint -Method Post -Body $requestBody

            if ($response) {
                Write-ColorOutput "Successfully created: $($row.Title)" -Type "Success"
                $successCount++
            } else {
                $errorMsg = "Failed to create training unit (API returned null)"
                Write-ColorOutput $errorMsg -Type "Error"
                $failedRows += [PSCustomObject]@{
                    Row = $currentRow
                    Title = $row.Title
                    Error = $errorMsg
                }
            }

        } catch {
            $errorMsg = $_.Exception.Message
            Write-ColorOutput "Error processing row: $errorMsg" -Type "Error"
            $failedRows += [PSCustomObject]@{
                Row = $currentRow
                Title = $row.Title
                Error = $errorMsg
            }
        }
    }

    # Display summary
    Write-ColorOutput "`n=== Import Summary ===" -Type "Info"
    Write-ColorOutput "Total rows processed: $($csvData.Count)" -Type "Info"
    Write-ColorOutput "Successful creations: $successCount" -Type "Success"
    Write-ColorOutput "Failed rows: $($failedRows.Count)" -Type $(if ($failedRows.Count -gt 0) { "Error" } else { "Success" })

    if ($failedRows.Count -gt 0) {
        Write-ColorOutput "`nFailed Rows Details:" -Type "Error"
        foreach ($failed in $failedRows) {
            Write-Host "  Row $($failed.Row) - $($failed.Title): $($failed.Error)" -ForegroundColor Red
        }
    }
}

#endregion

#region Main Execution

function Main {
    # Get user input
    $action = Get-UserInput

    # Authenticate
    $authSuccess = Get-BearerToken

    if (-not $authSuccess) {
        Write-ColorOutput "`nAuthentication failed. Exiting..." -Type "Error"
        return
    }

    # Execute selected action
    switch ($action) {
        "1" { Export-TrainingUnits }
        "2" { Import-TrainingUnits }
    }

    Write-ColorOutput "`nScript completed!" -Type "Success"
}

# Run main function
Main

#endregion
