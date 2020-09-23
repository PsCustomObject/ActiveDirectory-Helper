function Convert-ADExpirationDate
{
<#
    .SYNOPSIS
        Cmdlet is used to convert a user expiration date to human readable format
    
    .DESCRIPTION
        Cmdlet is used to convert internal AD expiration date format to human readable format.
    
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
                    Write-Warning -Message 'Account has no expiration date'
                }
                9223372036854775807
                {
                    Write-Warning -Message 'Account has no expiration date'
                }
                default
                {
                    return [datetime]::FromFileTime($ExpirationDate)
                }
            }
        }
    }
}

function Get-ExpiringUsers
{
    [OutputType([array])]
    param
    (
        [Parameter(ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server
    )
    
    # Define LDAP Filter
    [string]$ldapFilter = '(&(objectCategory=person)(objectClass=user)(accountExpires>=1)(accountExpires<=9223372036854775806))'
    
    if (!($PSBoundParameters.ContainsKey('Server')))
    {
        $Server = Get-AdServer
    }
    
    $paramGetADUser = @{
        LDAPFilter = $ldapFilter
        Server     = $Server
        Properties = 'accountExpires'
    }
    
    return Get-ADUser @paramGetADUser
}

function Get-ExpiringUsersReport
{
<#
    .SYNOPSIS
        A brief description of the Get-ExpiringUsers function.
    
    .DESCRIPTION
        A detailed description of the Get-ExpiringUsers function.
    
    .PARAMETER Server
        A string representing the Domain Controller to query. 
        
        If not specified a random domain controller in the local site will be used.
    
    .EXAMPLE
        		PS C:\> Get-ExpiringUsers
    
    .NOTES
        Additional information about the function.
#>
    
    [OutputType([array])]
    param
    (
        [Parameter(ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server
    )
    
    begin
    {
        # Define LDAP Filter
        [string]$ldapFilter = '(&(objectCategory=person)(objectClass=user)(accountExpires>=1)(accountExpires<=9223372036854775806))'
        
        # Initialize return array
        [System.Collections.ArrayList]$returnArray = @()
        
        if (!($PSBoundParameters.ContainsKey('Server')))
        {
            $Server = Get-AdServer
        }
    }
    process
    {
        $paramGetADUser = @{
            LDAPFilter = $ldapFilter
            Server     = $Server
            Properties = 'accountExpires'
        }
        
        [array]$matchingUsers = Get-ADUser @paramGetADUser
        
        foreach ($user in $matchingUsers)
        {
            # Format data
            $returnData = [pscustomobject]@{
                'User DN' = $user.'DistinguishedName'
                'Sam Account Name' = $user.'SamAccountName'
                'User UPN' = $user.'UserPrincipalName'
                'Object Guid' = $user.'ObjectGUID'
                'Account Enabled' = $user.'Enabled'
                'Expiration Date' = (Convert-ADExpirationDate -ExpirationDate ($user.'accountExpires'))
            }
            
            [void]($returnArray.Add($returnData))
        }
        
        # Return relevant objects
        return $returnArray
    }
}

function Get-AdServer
{
	<#
	.SYNOPSIS
		Return a string containing the name of a domain controller.
	
	.DESCRIPTION
		Function will return a string containing the name of a domain controller.
		
		If Active Directory module is installed on the machine it will be used to locate the closest domain controller, if module is not found function will use standard ADSI interface.
		
		Before returning the domain controller name it will be checked if server is available.
	
	.EXAMPLE
		PS C:\> Get-AdServer
	#>
    
    [OutputType([string])]
    param ()
    
    if (!(Get-Module -Name ActiveDirectory -ListAvailable))
    {
        # Control variable
        [bool]$isAvailable = $false
        
        do
        {
            # Get all Global Catalogs
            [array]$allAdGC = [DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GlobalCatalogs.Name
            
            # Get max number of GC
            [int]$arrayLength = $allAdGC.Length
            
            # Generate a random index
            [int]$arrayIndex = Get-Random -Minimum 0 -Maximum ($arrayLength - 1)
            
            # Save server
            $Server = $allAdGC[$arrayIndex]
            
            $isAvailable = Test-Connection -ComputerName $Server -Count 1 -Quiet
        }
        
        while ($isAvailable -ne $true)
        
        return $Server
    }
    else
    {
        do
        {
            # Pick closest Global Catalog
            $paramGetADDomainController = @{
                Discover        = $true
                NextClosestSite = $true
                Service         = 'GlobalCatalog'
            }
            
            [string]$Server = Get-ADDomainController @paramGetADDomainController
            
            $isAvailable = Test-Connection -ComputerName $Server -Count 1 -Quiet
        }
        
        while ($isAvailable -ne $true)
        
        return $Server
    }
}