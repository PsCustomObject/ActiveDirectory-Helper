function Convert-ADExpirationDate
{
<#
    .SYNOPSIS
        Cmdlet is used to convert a user expiration date to human redeable format
    
    .DESCRIPTION
        Cmdlet is used to convert internal AD expiration date format to human redeable format.
    
    .PARAMETER ExpirationDate
        A string or long representing the expiration date as exposed by the accountExpires attribute in Active Directory.
    
    .EXAMPLE
        PS C:\> Convert-AdExpirationDate -ExpirationDate 132566400000000000
#>
    
    [CmdletBinding(ConfirmImpact = 'High',
                   PositionalBinding = $true,
                   SupportsPaging = $false,
                   SupportsShouldProcess = $false)]
    [OutputType([datetime])]
    param
    (
        [Parameter(Mandatory = $true,
                   ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [long]
        $ExpirationDate
    )
    
    process
    {
        foreach ($date in $ExpirationDate)
        {
            switch ($ExpirationDate)
            {
                0
                {
                    Write-Verbose -Message 'Account has no expiration date'
                }
                9223372036854775807
                {
                    Write-Verbose -Message 'Account has no expiration date'
                }
                default
                {
                    [datetime]::FromFileTime($ExpirationDate)
                }
            }
        }
    }
}