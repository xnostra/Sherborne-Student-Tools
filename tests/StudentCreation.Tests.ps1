# Offline regression checks. No Microsoft Graph connection or account writes.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
foreach ($file in Get-ChildItem $root -Filter *.ps1) {
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors)
    Assert ($errors.Count -eq 0) "Parse errors in $($file.Name): $errors"
}
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Add-M365StudentsByName.ps1'), [ref]$null, [ref]$null)
foreach ($name in @('Get-CleanLetters', 'Test-UpnTaken', 'New-StudentUpn')) {
    $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    . ([scriptblock]::Create($node.Extent.Text))
}
$Domain = 'example.invalid'
$script:lookupMode = 'missing'
function Get-MgUser {
    [CmdletBinding()] param($UserId)
    if ($script:lookupMode -eq 'exists') { return [pscustomobject]@{ Id = 'test' } }
    if ($script:lookupMode -eq 'missing') {
        $record = New-Object System.Management.Automation.ErrorRecord ([Exception]::new('Not found')), 'Request_ResourceNotFound,Microsoft.Graph.PowerShell.Cmdlets.GetMgUser_Get', ([System.Management.Automation.ErrorCategory]::ObjectNotFound), $UserId
        $PSCmdlet.ThrowTerminatingError($record)
    }
    if ($script:lookupMode -eq 'status404') {
        $exception = [Exception]::new('HTTP failure')
        $exception | Add-Member NoteProperty ResponseStatusCode 404
        throw $exception
    }
    throw '403 Access denied'
}
$used = New-Object 'System.Collections.Generic.HashSet[string]'
Assert (-not (Test-UpnTaken 'test@example.invalid' $used)) 'SDK error ID 404 should mean free'
$script:lookupMode = 'status404'
Assert (-not (Test-UpnTaken 'test@example.invalid' $used)) 'Numeric 404 should mean free'
$script:lookupMode = 'exists'
Assert (Test-UpnTaken 'test@example.invalid' $used) 'Existing UPN should be taken'
$script:lookupMode = 'forbidden'
$threw = $false
try { Test-UpnTaken 'test@example.invalid' $used | Out-Null } catch { $threw = $true }
Assert $threw 'Permission failures must not mean free'
$script:lookupMode = 'missing'
$null = $used.Add('alex26q@example.invalid')
Assert ((New-StudentUpn 'Alexander Example' '26q' $used) -eq 'alexa26q@example.invalid') 'Collision should extend name'
. (Join-Path $root 'StudentLicenses.ps1')
function Make-Sku($Name, $Enabled, $Consumed, $Warning = 0, $Status = 'Enabled', $AppliesTo = 'User') {
    [pscustomobject]@{ SkuPartNumber = $Name; ConsumedUnits = $Consumed; CapabilityStatus = $Status; AppliesTo = $AppliesTo; PrepaidUnits = @{ Enabled = $Enabled; Warning = $Warning; Suspended = 0; LockedOut = 0 } }
}
Assert ((Get-AssignableSeatCount (Make-Sku 'test' 10 7 2)) -eq 5) 'Seat calculation'
Assert ((Get-AssignableSeatCount (Make-Sku 'test' 10 12)) -eq 0) 'No negative seats'
Assert ((Get-AssignableSeatCount (Make-Sku 'test' 10 0 0 'Suspended')) -eq 0) 'Suspended cannot be selected'
Assert ((Get-AssignableSeatCount (Make-Sku 'test' 10 0 0 'Enabled' 'Company')) -eq 0) 'Company cannot be assigned to users'
function Get-MgContext { @{ Account = 'offline'; TenantId = 'test' } }
function Get-MgSubscribedSku { [CmdletBinding()] param([switch]$All) $script:testSkus }
function Read-Host { param($Prompt) if ($script:answers.Count -eq 0) { throw 'Unexpected extra prompt' }; $script:answers.Dequeue() }
$script:testSkus = @((Make-Sku 'M365EDU_A5_STUUSEBNFT' 2 2), (Make-Sku 'STANDARDWOFFPACK_STUDENT' 10 1))
$script:answers = [System.Collections.Generic.Queue[string]]::new()
foreach ($answer in @('', '-1', 'abc', '99999999999999', '99', '0', '1')) { $script:answers.Enqueue($answer) }
$selected = Select-StudentLicense
Assert ($selected.SkuPartNumber -eq 'STANDARDWOFFPACK_STUDENT') 'Invalid/unavailable selections should reprompt'
Assert ($script:answers.Count -eq 0) 'All invalid inputs should have been rejected'
$script:testSkus = @((Make-Sku 'M365EDU_A5_STUUSEBNFT' 10 1))
$script:answers.Enqueue('')
$selected = Select-StudentLicense
Assert ($selected.SkuPartNumber -eq 'M365EDU_A5_STUUSEBNFT') 'Recommendation at index zero should work'
Write-Host 'PASS: parsing, UPN lookups, collision handling, license counts and selection.'
