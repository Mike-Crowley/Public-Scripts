function Convert-GuidFormat {
    <#
    .SYNOPSIS
        Converts between standard GUID, ImmutableId (Base64), and hex byte representations.

    .DESCRIPTION
        Automatically detects the input format and returns all three representations of the same
        GUID. Useful when working with Entra Connect, Entra ID, and Active Directory, where
        the same object identifier appears in different encodings depending on the tool.

        Supported formats:
          - Standard GUID string    25b5dd02-494c-452f-bbc2-54ff54bb861c
          - ImmutableId (Base64)    At21JUxJL0W7wlT/VLuGHA==
          - Hex bytes with spaces   02 DD B5 25 4C 49 2F 45 BB C2 54 FF 54 BB 86 1C

    .PARAMETER InputValue
        A string containing a GUID in any of the three supported formats.
        Accepts pipeline input.

    .EXAMPLE
        Convert-GuidFormat "25b5dd02-494c-452f-bbc2-54ff54bb861c"

        GuidString  : 25b5dd02-494c-452f-bbc2-54ff54bb861c
        ImmutableId : At21JUxJL0W7wlT/VLuGHA==
        GuidHex     : 02 DD B5 25 4C 49 2F 45 BB C2 54 FF 54 BB 86 1C

    .EXAMPLE
        Convert-GuidFormat "At21JUxJL0W7wlT/VLuGHA=="

        Converts an ImmutableId (Base64) value to all three formats.

    .EXAMPLE
        Convert-GuidFormat "02 DD B5 25 4C 49 2F 45 BB C2 54 FF 54 BB 86 1C"

        Converts a hex byte string (as seen in LDAP editors) to all three formats.

    .EXAMPLE
        "At21JUxJL0W7wlT/VLuGHA==", "25b5dd02-494c-452f-bbc2-54ff54bb861c" | Convert-GuidFormat

        Accepts multiple values from the pipeline.

    .INPUTS
        System.String

    .OUTPUTS
        PSCustomObject with properties: GuidString, ImmutableId, GuidHex

    .NOTES
        Author:  Mike Crowley
        https://mikecrowley.us

    .LINK
        https://learn.microsoft.com/en-us/entra/identity/hybrid/connect/plan-connect-design-concepts#sourceanchor
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$InputValue
    )

    process {
        $Guid = $null

        # Method 1: Standard GUID string
        try { $Guid = [Guid]$InputValue } catch { }

        # Method 2: ImmutableId (Base64)
        if ($null -eq $Guid) {
            try {
                $bytes = [Convert]::FromBase64String($InputValue)
                if ($bytes.Length -eq 16) { $Guid = [Guid]::new($bytes) }
            }
            catch { }
        }

        # Method 3: Hex bytes with spaces
        if ($null -eq $Guid) {
            try {
                $hexValues = $InputValue.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
                if ($hexValues.Count -eq 16) {
                    $bytes = [byte[]]::new(16)
                    for ($i = 0; $i -lt 16; $i++) {
                        $bytes[$i] = [Convert]::ToByte($hexValues[$i], 16)
                    }
                    $Guid = [Guid]::new($bytes)
                }
            }
            catch { }
        }

        if ($null -ne $Guid) {
            [PSCustomObject]@{
                GuidString  = $Guid.ToString()
                ImmutableId = [Convert]::ToBase64String($Guid.ToByteArray())
                GuidHex     = [BitConverter]::ToString($Guid.ToByteArray()).Replace("-", " ")
            }
        }
        else {
            Write-Error "Cannot parse '$InputValue' as a GUID, ImmutableId, or hex byte string."
        }
    }
}
