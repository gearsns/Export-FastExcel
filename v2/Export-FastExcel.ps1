function Export-FastExcel {
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DataArray')]
        [object[]]$Data,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'InputObject')]
        [psobject]$InputObject, # For pipeline input of single objects

        [Parameter(Mandatory = $true)]
        [string]$Filename
    )
    begin {
        $script:excelData = [System.Collections.Generic.List[object[]]]::new()
    }
    process {
        if ($PSCmdlet.ParameterSetName -eq 'InputObject') {
            $currentRowData = $null
            if ($InputObject -is [System.Array] -or $InputObject -is [System.Data.DataRow]) {
                $currentRowData = $InputObject
            } else {
                $currentRowData = @($InputObject)
            }
            if ($currentRowData) {
                $script:excelData.Add($currentRowData)
            }
        } elseif ($PSCmdlet.ParameterSetName -eq 'DataArray') {
            # If data is an array of arrays (e.g., from CSV import)
            if ($Data -is [System.Array] -and $Data.Length -gt 0) {
                if ($Data[0] -is [System.Array] -or $Data[0] -is [System.Data.DataRow]){
                    foreach ($item in $Data) {
                        # AddRange ではなく Add を使う
                        $script:excelData.Add($item)
                    }
                } else {
                    $script:excelData.AddRange($Data)
                }
            } else {
                # Handle single object or array of objects
                foreach ($item in $Data) {
                    if ($item -is [System.Array]) {
                        $script:excelData.Add($item)
                    } else {
                        $script:excelData.Add(@($item))
                    }
                }
            }
        }
    }
    end {
        if ($script:excelData.Count -eq 0) {
            Write-Warning "No data to export. Exiting."
            return
        }
        Add-Type -Path "$PSScriptRoot/Export-FastExcel.cs" -ReferencedAssemblies @(
            "System.Console", "System.Collections", 
            "System.Data", "System.Data.Common",
            "System.IO.Compression", "System.IO.Compression.ZipFile", 
            "System.IO.Compression.FileSystem",
            "System.Xml", "System.Text.RegularExpressions", "System.Management.Automation"
        )
        # Export
        [ExportFastExcelCS]::Export("$PSScriptRoot/Template/", $Filename, $script:excelData)
     }
}
