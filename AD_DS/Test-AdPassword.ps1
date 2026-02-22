<#
.SYNOPSIS
    Validates the credentials of a PSCredential object against the local machine or Active Directory.

.DESCRIPTION
    Tests whether a PSCredential (from Get-Credential) contains a valid password by authenticating
    against either the local machine or an Active Directory domain using the
    System.DirectoryServices.AccountManagement API.

    For domain validation, the user is first resolved via FindByIdentity to support UPN-style
    usernames. Note that this approach ignores the NetBIOS domain portion of the credential,
    meaning fakedomain\user1 will be evaluated as user1.

.PARAMETER PSCredential
    A PSCredential object containing the username and password to validate.

.PARAMETER LocalUser
    When $true, validates against the local machine. When $false or omitted, validates
    against the Active Directory domain.

.PARAMETER DomainFQDN
    The fully qualified domain name to validate against. Defaults to the current user's
    domain ($env:USERDNSDOMAIN).

.EXAMPLE
    . .\Test-AdPassword.ps1
    $MyCred = Get-Credential
    Test-Credential -Verbose -PSCredential $MyCred

.EXAMPLE
    $MyCred = Get-Credential
    Test-Credential -Verbose -PSCredential $MyCred -LocalUser $true

.NOTES
    Author: Mike Crowley
    https://mikecrowley.us

    References:
        https://stackoverflow.com/questions/290548/validate-a-username-and-password-against-active-directory/499716
        https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/validating-user-account-passwords-part-1

.LINK
    https://github.com/Mike-Crowley/Public-Scripts
#>

function Test-Credential {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCredential]$PSCredential,

        [Parameter()]
        [Boolean]$LocalUser,

        [Parameter()]
        [String]$DomainFQDN = $env:USERDNSDOMAIN
    )

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement

    $ResolveUser = $PSCredential.GetNetworkCredential().domain.Length -lt 1

    If ($LocalUser) {
        $context = [DirectoryServices.AccountManagement.ContextType]::Machine
        $PrincipalContext = [DirectoryServices.AccountManagement.PrincipalContext]::new($context, $env:COMPUTERNAME)
    }
    else {
        $Context = [DirectoryServices.AccountManagement.ContextType]::Domain
        $PrincipalContext = [DirectoryServices.AccountManagement.PrincipalContext]::new($context, $DomainFQDN)
        $PasswordToValidate = $PSCredential.GetNetworkCredential().Password

        if ($ResolveUser) {
            Write-Verbose "Searching AD. Please wait...`n"
            # FindByIdentity seems to ignore the netbios domain part. This means fakedomain\user1 will be evaluated as user1.
            $ADUser = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($PrincipalContext, $PSCredential.UserName)
            $UserToValidate = $ADUser.SamAccountName
        }
        else {
            $UserToValidate = $PSCredential.GetNetworkCredential().UserName
        }
    }

    Write-Verbose "Testing credentials. Please wait...`n"

    $result = $PrincipalContext.ValidateCredentials($UserToValidate, $PasswordToValidate)

    # Guarantees a boolean output (vs null due to errors)
    return $result -eq $true
}

# Test-Credential -Verbose -PSCredential (Get-Credential)
