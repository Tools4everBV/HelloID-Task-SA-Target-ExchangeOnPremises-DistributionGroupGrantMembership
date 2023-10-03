# HelloID-Task-SA-Target-ExchangeOnPremises-DistributionGroupGrantMembership
############################################################################
# Form mapping
$formObject = @{
    GroupIdentity = $form.GroupIdentity
    UsersToAdd    = [array]$form.Users
}

[bool]$IsConnected = $false
try {
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Add-DistributionGroupMember'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToAdd) {
        try {
            Write-Information "Executing ExchangeOnPremises action: [DistributionGroupGrantMembership] for: [$($formObject.GroupIdentity)]"
            $null = Add-DistributionGroupMember -Identity $formObject.GroupIdentity -Member $user.UserPrincipalName -Confirm:$false -ErrorAction Stop

            $auditLog = @{
                Action            = 'GrantMembership'
                System            = 'ExchangeOnPremises'
                TargetIdentifier  = $formObject.GroupIdentity
                TargetDisplayName = $formObject.GroupIdentity
                Message           = "ExchangeOnPremises action: [DistributionGroupGrantMembership] user [$($user.UserPrincipalName)] to: [$($formObject.GroupIdentity)] executed successfully"
                IsError           = $false
            }
            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information "ExchangeOnPremises action: [DistributionGroupGrantMembership] user [$($user.UserPrincipalName)] to: [$($formObject.GroupIdentity)] executed successfully"
        } catch {
            $ex = $_
            if ($ex.CategoryInfo.Reason -eq 'MemberAlreadyExistsException') {
                $auditLog = @{
                    Action            = 'GrantMembership'
                    System            = 'ExchangeOnPremises'
                    TargetIdentifier  = $formObject.GroupIdentity
                    TargetDisplayName = $formObject.GroupIdentity
                    Message           = "ExchangeOnPremises action: [DistributionGroupGrantMembership] user [$($user.UserPrincipalName)] to: [$($formObject.GroupIdentity)] executed successfully"
                    IsError           = $false
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Information "ExchangeOnPremises action: [DistributionGroupGrantMembership] user [$($user.UserPrincipalName)] to: [$($formObject.GroupIdentity)] executed successfully"
            } else {
                $auditLog = @{
                    Action            = 'GrantMembership'
                    System            = 'ExchangeOnPremises'
                    TargetIdentifier  = $formObject.GroupIdentity
                    TargetDisplayName = $formObject.GroupIdentity
                    Message           = "Could not execute ExchangeOnPremises action: [DistributionGroupGrantMembership] user [$($user.UserPrincipalName)] to: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
                    IsError           = $true
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Error "Could not execute ExchangeOnPremises action: [DistributionGroupGrantMembership] user [$($user.UserPrincipalName)] to: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
            }
        }
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'GrantMembership'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.GroupIdentity
        TargetDisplayName = $formObject.GroupIdentity
        Message           = "Could not execute ExchangeOnPremises action: [DistributionGroupGrantMembership] to: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [DistributionGroupGrantMembership] to: [$($formObject.GroupIdentity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
############################################################################
