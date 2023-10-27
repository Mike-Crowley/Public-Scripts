<#
.SYNOPSIS 
    Practicing with functions just for fun. Convert-CaesarCipher encodes or decodes case-sensitive English strings based on the Caesar Cipher.

.DESCRIPTION
    Practicing with functions just for fun. Convert-CaesarCipher encodes or decodes case-sensitive English strings based on the Caesar Cipher.
    https://en.wikipedia.org/wiki/Caesar_cipher

.EXAMPLE

    Convert-CaesarCipher -InputString "Hello World" -Shift 20
    # Encode is the default.
    # Returns: Byffi Qilfx

.EXAMPLE

    "Hello World" | Convert-CaesarCipher -Shift 20
    # Pipeline input is accepted.
    # Returns: Byffi Qilfx

.EXAMPLE

    Convert-CaesarCipher "Hello World" -Shift 20
    # InputString is a positional parameter.
    # Returns: Byffi Qilfx

.EXAMPLE

    Convert-CaesarCipher -InputString "Hello World" -Shift 20 -Encode
    # You can also specify Encode to avoid confusion.
    # Returns: Byffi Qilfx
    
.EXAMPLE

    Convert-CaesarCipher -InputString "Byffi Qilfx" -Shift 20 -Decode
    # Returns: Hello World

.LINK 

    https://mikecrowley.us
#>

function Convert-CaesarCipher {

    [CmdletBinding(DefaultParameterSetName = 'Encode')]
    param (
        [parameter(ParameterSetName = "Encode")][switch]$Encode,
        [parameter(ParameterSetName = "Decode")][switch]$Decode, 
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            HelpMessage = "Please provide an input string containing only English letters spaces."
        )]
        [ValidateScript({
                if ($_ -match '^[a-z ]+$') { $true } else { throw "Input must contain only English letters spaces." }
            })]
        [string]$InputString,
        [Parameter(
            Mandatory = $true,
            HelpMessage = "Please provide a numerical Shift value."
        )]
        [int]$Shift
    )

    $LowerAlphabet = "abcdefghijklmnopqrstuvwxyz"
    $UpperAlphabet = $LowerAlphabet.ToUpper()
    $output = ""

    foreach ($Char in $InputString.ToCharArray()) {
        [string]$StringChar = $Char
        if ($StringChar -eq " ") {
            $output += " "
        }
        elseif ($StringChar.ToUpper() -ceq $StringChar) {
            if ($Decode) {
                $index = ($UpperAlphabet.IndexOf($StringChar) - $Shift) % $UpperAlphabet.Length
                $output += $UpperAlphabet[$index]
            }
            else {
                $index = ($UpperAlphabet.IndexOf($StringChar) + $Shift) % $UpperAlphabet.Length
                $output += $UpperAlphabet[$index]
            }
        }
        else {
            if ($Decode) {
                $index = ($LowerAlphabet.IndexOf($StringChar) - $Shift) % $LowerAlphabet.Length
                $output += $LowerAlphabet[$index]
            }
            else {
                $index = ($LowerAlphabet.IndexOf($StringChar) + $Shift) % $LowerAlphabet.Length
                $output += $LowerAlphabet[$index]
            }
        }
    }
    return $output
}
