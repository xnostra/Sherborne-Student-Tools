# Shared license inventory and selection for both student creation workflows.
function Get-AssignableSeatCount {
    param($SubscribedSku)
    if ($SubscribedSku.AppliesTo -ne 'User' -or $SubscribedSku.CapabilityStatus -notin @('Enabled', 'Warning')) { return 0 }
    return [Math]::Max(0, ([int]$SubscribedSku.PrepaidUnits.Enabled + [int]$SubscribedSku.PrepaidUnits.Warning - [int]$SubscribedSku.ConsumedUnits))
}

function Select-StudentLicense {
    $context = Get-MgContext
    Write-Host "`nSigned in as: $($context.Account) | Tenant: $($context.TenantId)" -ForegroundColor Cyan
    $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop | Sort-Object SkuPartNumber)
    if ($skus.Count -eq 0) { throw 'Microsoft Graph returned no subscriptions for this tenant.' }
    $labels = @{
        M365EDU_A5_STUUSEBNFT = 'Microsoft 365 A5 for Students (Student Use Benefit)'
        M365EDU_A5_STUDENT = 'Microsoft 365 A5 for Students'
        STANDARDWOFFPACK_STUDENT = 'Office 365 A1 for Students'
        STANDARDWOFFPACK_FACULTY = 'Office 365 A1 for Faculty'
    }
    Write-Host "`nAll tenant subscriptions (including full or unavailable subscriptions):"
    for ($i = 0; $i -lt $skus.Count; $i++) {
        $s = $skus[$i]
        $label = $labels[$s.SkuPartNumber]
        if (-not $label) { $label = $s.SkuPartNumber }
        Write-Host "[$i] $label [$($s.SkuPartNumber)]"
        Write-Host "    Available: $(Get-AssignableSeatCount $s) | Assigned: $($s.ConsumedUnits) | Enabled: $($s.PrepaidUnits.Enabled) | Grace: $($s.PrepaidUnits.Warning) | Suspended: $($s.PrepaidUnits.Suspended) | Locked out: $($s.PrepaidUnits.LockedOut) | Status: $($s.CapabilityStatus) | Applies to: $($s.AppliesTo)"
    }
    $a5 = @(0..($skus.Count - 1) | Where-Object { $skus[$_].SkuPartNumber -match '^(M365EDU_A5|ENTERPRISEPREMIUM).*_(STUDENT|STUUSEBNFT)$' })
    $recommended = @($a5 | Where-Object { (Get-AssignableSeatCount $skus[$_]) -gt 0 })
    if ($a5.Count -eq 0) { Write-Warning 'Graph returned no A5 student subscription for the tenant shown above. Check the tenant and subscription in the admin centre.' }
    elseif ($recommended.Count -eq 0) { Write-Warning 'A5 student subscriptions are listed above, but none currently has assignable seats.' }
    if (-not @($skus | Where-Object { (Get-AssignableSeatCount $_) -gt 0 }).Count) { throw 'No user licenses with assignable seats were returned. See the subscription statuses above.' }
    if ($recommended.Count -gt 0) { Write-Host "Recommended: [$($recommended[0])] $($skus[$recommended[0]].SkuPartNumber)" -ForegroundColor Yellow }
    while ($true) {
        $choice = Read-Host 'Enter a license number (Enter uses the recommendation, if shown)'
        if ([string]::IsNullOrWhiteSpace($choice) -and $recommended.Count -gt 0) { $choice = [string]$recommended[0] }
        $index = 0
        if ($choice -notmatch '^\d+$' -or -not [int]::TryParse($choice, [ref]$index) -or $index -ge $skus.Count) {
            Write-Warning 'Enter one of the displayed license numbers.'
            continue
        }
        if ((Get-AssignableSeatCount $skus[$index]) -le 0) {
            Write-Warning 'That subscription has no assignable user seats. Choose another license.'
            continue
        }
        Write-Host "Using license: $($skus[$index].SkuPartNumber)"
        return $skus[$index]
    }
}
