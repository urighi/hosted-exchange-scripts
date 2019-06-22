Function New-Password {
<#
.SYNOPSIS
   Generate password from random characters
.DESCRIPTION
   This script will generate a string of random characters of any length. The string can be numeric, symbolic or alphabetic
   upper and/or lower case.
.EXAMPLE

    .\New-Password -length 15 -LowerCase -Uppercase
    KXVwalbDPmreNpgh

    This example will generate and return a 15 character string comprised of lowercase and uppercase random characters.
.EXAMPLE    

    New-Password -length 8 -Numeric
    421768395

    This example will generate and return an 8 character string comprised of numeric random characters.
.EXAMPLE

    .\New-Password -length 30 -Uppercase -LowerCase -Numeric -Symbols
    ftA8Uj&Q6h1M5Zomr-)bia'9,FxgVTn

    This example will generate and return a 30 character string comprised of Uppercase, LowerCase, Numeric and Symbolic random characters.
.NOTES
    Author: Scriptimus Prime
    Version: 1.01
    Created: 5 July 2014
    Last updated: 12 April 2014
    Website: http://scriptimus.wordpress.com
#>

    [CmdletBinding()]
    [OutputType([String])]

    
    Param(

        # This parameter determines the lengh of the output string.
        [int]$length=30,

        # Use this parameter to include the uppercase characters "A -Z".
        [alias("U")]
        [Switch]$Uppercase,

        # Use this parameter to include the lowerCase characters "a - z".
        [alias("L")]
        [Switch]$LowerCase,

        # Use this parameter to include numeric characters "0 - 9".
        [alias("N")]
        [Switch]$Numeric,

        # Use this parameter to include symbolic characters ";<=>?@!"#$%&'()*+,-./".
        [alias("S")]
        [Switch]$Symbolic

    )

    Begin {}

    Process {
        
        If ($Uppercase) {$CharPool += ([char[]](64..90))}
        If ($LowerCase) {$CharPool += ([char[]](97..122))}
        If ($Numeric) {$CharPool += ([char[]](48..57))}
        If ($Symbolic) {$CharPool += ([char[]](33..47))
                       $CharPool += ([char[]](33..47))}
        
        If ($CharPool -eq $null) {
            Throw 'You must select at least one of the parameters "Uppercase" "LowerCase" "Numeric" or "Symbolic"'
        }

        [String]$Password =  (Get-Random -InputObject $CharPool -Count $length) -join ''

    }
    
    End {
        
        return $Password
    
    }
}