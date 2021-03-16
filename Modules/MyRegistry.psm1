# functions for registry management
function Search-Registry { 
    <# 
    .SYNOPSIS 
    Searches registry key names, value names, and value data (limited). 
    
    .DESCRIPTION 
    This function can search registry key names, value names, and value data (in a limited fashion). It outputs custom objects that contain the key and the first match type (KeyName, ValueName, or ValueData). 
    
    .EXAMPLE 
    Search-Registry -Path HKLM:\SYSTEM\CurrentControlSet\Services\* -SearchRegex "svchost" -ValueData 
    
    .EXAMPLE 
    Search-Registry -Path HKLM:\SOFTWARE\Microsoft -Recurse -ValueNameRegex "ValueName1|ValueName2" -ValueDataRegex "ValueData" -KeyNameRegex "KeyNameToFind1|KeyNameToFind2" 
    
    #> 

    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory, Position=0, ValueFromPipelineByPropertyName)] 
        [Alias("PsPath")] 
        # Registry path to search 
        [string[]] $Path, 
        # Specifies whether or not all subkeys should also be searched 
        [switch] $Recurse, 
        [Parameter(ParameterSetName="SingleSearchString", Mandatory)] 
        # A regular expression that will be checked against key names, value names, and value data (depending on the specified switches) 
        [string] $SearchRegex, 
        [Parameter(ParameterSetName="SingleSearchString")] 
        # When the -SearchRegex parameter is used, this switch means that key names will be tested (if none of the three switches are used, keys will be tested) 
        [switch] $KeyName, 
        [Parameter(ParameterSetName="SingleSearchString")] 
        # When the -SearchRegex parameter is used, this switch means that the value names will be tested (if none of the three switches are used, value names will be tested) 
        [switch] $ValueName, 
        [Parameter(ParameterSetName="SingleSearchString")] 
        # When the -SearchRegex parameter is used, this switch means that the value data will be tested (if none of the three switches are used, value data will be tested) 
        [switch] $ValueData, 
        [Parameter(ParameterSetName="MultipleSearchStrings")] 
        # Specifies a regex that will be checked against key names only 
        [string] $KeyNameRegex, 
        [Parameter(ParameterSetName="MultipleSearchStrings")] 
        # Specifies a regex that will be checked against value names only 
        [string] $ValueNameRegex, 
        [Parameter(ParameterSetName="MultipleSearchStrings")] 
        # Specifies a regex that will be checked against value data only 
        [string] $ValueDataRegex 
    ) 

    begin { 
        switch ($PSCmdlet.ParameterSetName) { 
            SingleSearchString { 
                $NoSwitchesSpecified = -not ($PSBoundParameters.ContainsKey("KeyName") -or $PSBoundParameters.ContainsKey("ValueName") -or $PSBoundParameters.ContainsKey("ValueData")) 
                if ($KeyName -or $NoSwitchesSpecified) { $KeyNameRegex = $SearchRegex } 
                if ($ValueName -or $NoSwitchesSpecified) { $ValueNameRegex = $SearchRegex } 
                if ($ValueData -or $NoSwitchesSpecified) { $ValueDataRegex = $SearchRegex } 
            } 
            MultipleSearchStrings { 
                # No extra work needed 
            } 
        } 
    } 

    process { 
        foreach ($CurrentPath in $Path) { 
            Get-ChildItem $CurrentPath -Recurse:$Recurse |  
                ForEach-Object { 
                    $Key = $_ 

                    if ($KeyNameRegex) {  
                        Write-Verbose ("{0}: Checking KeyNamesRegex" -f $Key.Name)  

                        if ($Key.PSChildName -match $KeyNameRegex) {  
                            Write-Verbose "  -> Match found!" 
                            return [PSCustomObject] @{ 
                                Key = $Key 
                                Reason = "KeyName" 
                            } 
                        }  
                    } 

                    if ($ValueNameRegex) {  
                        Write-Verbose ("{0}: Checking ValueNamesRegex" -f $Key.Name) 

                        if ($Key.GetValueNames() -match $ValueNameRegex) {  
                            Write-Verbose "  -> Match found!" 
                            return [PSCustomObject] @{ 
                                Key = $Key 
                                Reason = "ValueName" 
                            } 
                        }  
                    } 

                    if ($ValueDataRegex) {  
                        Write-Verbose ("{0}: Checking ValueDataRegex" -f $Key.Name) 

                        if (($Key.GetValueNames() | % { $Key.GetValue($_) }) -match $ValueDataRegex) {  
                            Write-Verbose "  -> Match!" 
                            return [PSCustomObject] @{ 
                                Key = $Key 
                                Reason = "ValueData" 
                            } 
                        } 
                    } 
                } 
        } 
    } 
} 
Export-ModuleMember -Function Search-Registry