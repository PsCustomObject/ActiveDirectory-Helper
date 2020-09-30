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
    
        Monday, February 1, 2021 9:00:00 AM
    
    .EXAMPLE
    
        PS C:\> Convert-AdExpirationDate -ExpirationDate 9223372036854775807
    
        WARNING: Account has no expiration date
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

function Test-IsUniqueAddress
{
<#
    .SYNOPSIS
        Cmdlet will check if a given email address is unisque in the forest.
    
    .DESCRIPTION
        Cmdlet will check if a given email address is unisque in the forest checking both ProxyAddresses and Mail attributes.
    
    .PARAMETER Server
        A string representing the domain controller that will be used to run the query.
    
        If parameter is not specified a Domain Controller will be dynamically searched and used among the ones available in the closest site.
    
    .PARAMETER MailAddress
        A string or array of strings containing email addresses to be checked for uniqueness.
    
        If supplied email address is not a valid email address a warning will be printed on screen and nothing will be returned.
    
    .EXAMPLE
        PS C:\> Test-IsUniqueAddress -MailAddress 'value1'
#>
    [CmdletBinding(ConfirmImpact = 'High',
                   SupportsPaging = $false)]
    param
    (
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory = $true,
                   ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $MailAddress
    )
    
    begin
    {
        if (!($PSBoundParameters.ContainsKey('Server')))
        {
            $Server = Get-AdServer
        }
    }
    process
    {
        foreach ($mail in $MailAddress)
        {
            if (Test-IsEmail -EmailAddress $mail)
            {
                # Define LDAP Filter
                [string]$ldapFilter = "(&(objectCategory=person)(objectClass=user)(|(proxyAddresses=*:$mail)(mail=$mail)))"
                
                # Get any matching user
                $paramGetADUser = @{
                    LDAPFilter = $ldapFilter
                    Server     = $Server
                }
                
                [array]$matchingUsers = Get-ADUser @paramGetADUser
                
                if ($matchingUsers.Count -gt 0)
                {
                    return $false
                }
                else
                {
                    return $true
                }
            }
            else
            {
                Write-Warning -Message "$mail is not a valid email address"
            }
        }
    }
}

function Test-IsEmail
{
	<#
	.SYNOPSIS
		Cmdlet will check if a string is an RFC email address.
	
	.DESCRIPTION
		Cmdlet will check if an input string is an RFC complient email address. 
	
	.PARAMETER EmailAddress
		A string representing the email address to be checked
	
	.EXAMPLE
		PS C:\> Test-IsEmail -EmailAddress 'value1'
	
	.OUTPUTS
		System.Boolean
	
	.LINK
		Restrictions on email addresses
		https://tools.ietf.org/html/rfc3696#section-3
	#>
    
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Email', 'Mail', 'Address')]
        [string]
        $EmailAddress
    )
    
    try
    {
        # Check if address is RFC compliant	
        [void]([mailaddress]$EmailAddress)
        
        Write-Verbose -Message "Address $EmailAddress is an RFC compliant address"
        
        return $true
    }
    catch
    {
        Write-Verbose -Message "Address $EmailAddress is not an RFC compliant address"
        
        return $false
    }
}

function Test-IsGuid
{
    <#
        .SYNOPSIS
            Cmdlet will check if input string is a valid GUID.
        
        .DESCRIPTION
            Cmdlet will check if input string is a valid GUID.
        
        .PARAMETER ObjectGuid
            A string representing the GUID to be tested.
        
        .EXAMPLE
            PS C:\> Test-IsGuid -ObjectGuid 'value1'
        
            # Output
            $False
        
        .EXAMPLE
            PS C:\> Test-IsGuid -ObjectGuid '7761bf39-9a9f-42c8-869f-7c6e2689811a'
        
            # Output
            $True
        
        .OUTPUTS
            System.Boolean
        
        .NOTES
            Additional information about the function.
    #>
    
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $ObjectGuid
    )
    
    # Define verification regex
    [regex]$guidRegex = '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'
    
    # Check guid against regex
    return $ObjectGuid -match $guidRegex
}

function Format-Base64ToGuid
{
<#
    .SYNOPSIS
        Cmdlet will convert a Base64 encoded string to GUID format.
    
    .DESCRIPTION
        Cmdlet will convert a Base64 encoded string to GUID format.
    
    .PARAMETER InputString
        A string representing the encoded input to be converted to GUID.
    
    .EXAMPLE
        PS C:\> Format-Base64ToGuid -InputString 'value1'
#>
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $InputString
    )
        
    return New-Object -TypeName System.Guid -ArgumentList (, (([System.Convert]::FromBase64String($InputString))))
}

function Format-GuidToOctectString
{
<#
    .SYNOPSIS
        Cmdlet will con
    
    .DESCRIPTION
        A detailed description of the Format-GuidToOctectString function.
    
    .PARAMETER InputGuid
        A string representing the GUID to be converted.
    
    .EXAMPLE
        PS C:\> Format-GuidToOctectString -InputGuid '6f577363-83aa-4a18-9b68-3575ad19292b'
   
        # Output
        7552b5c081b20e43ac79a13f5048e3ea
#>
    
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $InputGuid
    )
    
    if (Test-IsGuid -ObjectGuid $InputGuid)
    {
        [System.String]::Join('', ((New-Object -TypeName System.Guid($InputGuid)).ToByteArray() | ForEach-Object { $_.ToString('x2') }))
    }
    else
    {
        Write-Warning -Message 'Input string is not a valid GUID'
    }
}

function Format-GuidToBase64
{
<#
    .SYNOPSIS
        Cmdlet will convert a GUID to Base64 Encoded string.
    
    .DESCRIPTION
        Cmdlet will convert a GUID to Base64 Encoded string.
    
    .PARAMETER InputGuid
        A string representing the GUID to be converted.
    
    .EXAMPLE
        PS C:\> Format-GuidToOctectString -InputGuid '6f577363-83aa-4a18-9b68-3575ad19292b'
   
        # Output
        dVK1wIGyDkOseaE/UEjj6g==
    
    .OUTPUTS
        System.String
#>
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $InputGuid
    )
    
    if (Test-IsGuid -ObjectGuid $InputGuid)
    {
        return [System.Convert]::ToBase64String((New-Object System.Guid($InputGuid)).ToByteArray())
    }
    else
    {
        Write-Warning -Message 'Input string is not a valid GUID'
    }
}