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
        $script:columnRefCache = @{} # Cache for column references
        $script:excelStartTime = [datetime]::SpecifyKind((Get-Date "1899-12-30 00:00:00"), [DateTimeKind]::Utc)
        # Helper function to escape special characters for XML
        function ReplaceSpecialCharacter
        {
            param($str)
            $html_conv_char=@{'&'='&amp;';'<'='&lt;';'>'='&gt;';"\r\n"='<br>';"\n"='<br>';"\r"='<br>'}
            [regex]::Replace($str, "[&<>]", {
                param($match)
                $html_conv_char[$match.Value]
            })
        }
        # Helper function to convert column index to Excel column reference (A, B, AA, etc.)
        function Convert-ColumnIndexToRef {
            param (
                [int]$ColumnIndex
            )
            $ref = ''
            $col = $ColumnIndex - 1
            while ($true) {
                $x = $col % 26
                $ref = [char](65 + $x) + $ref
                $col = [Math]::Floor($col / 26) - 1
                if ($col -lt 0) {
                    break
                }
            }
            $ref
        }
        # Helper function to get cell reference (e.g., A1, B2)
        function Get-CellRef {
            param (
                [int]$Row,
                [int]$Col
            )
            if (-not $script:columnRefCache.ContainsKey($Col)) {
                $script:columnRefCache[$Col] = Convert-ColumnIndexToRef $Col
            }
            "$($script:columnRefCache[$Col])$($Row + 1)"
        }
    }
    process {
        if ($PSCmdlet.ParameterSetName -eq 'InputObject') {
            $currentRowData = $null
            if ($InputObject -is [System.Array]) {
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
                if ($Data[0] -is [System.Array]){
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
        # Determine headers
        $headers = @()
        $startRowIndex  = 0
        $firstRowData = $script:excelData[0]
        if ($firstRowData -is [System.Array]) {
            if ($firstRowData[0] -is [psobject]) {
                $headers = $firstRowData[0].psobject.Properties.Name | ForEach-Object { $_ }
            } else {
                $headers = $firstRowData
                $startRowIndex  = 1
            }
        } elseif ($firstRowData -is [psobject]) {
            $headers = $firstRowData.psobject.Properties.Name | ForEach-Object { _ }
        } else {
            $headers = 1..$firstRowData.Length | ForEach-Object { $_ }
        }

        # Generate XML parts for headers
        $used = @{}
        $headerMap = ""
        $tableColumn = ""
        $index = 0
        $tmpHeaders = [System.Collections.Generic.List[string]]::new()
        foreach ($header in $headers){
            $name = $header
            $index++
            while ($tmpHeaders.Contains($name)){
                if (!$used.Contains($header)){
                    $used[$header] = 0
                }
                $used[$header] += 1
                $name = "$header ($($used[$header]))"
            }
            $headers[$index-1] = $name
            $tmpHeaders.Add($name)
            $script:columnRefCache[$index-1] = Convert-ColumnIndexToRef $index
            $xml_str = $name -replace '"', '&quot;'
            $tableColumn += "<tableColumn id=""$index"" name=""$xml_str""/>"
            $headerMap += "<si><t>$(ReplaceSpecialCharacter $name)</t></si>" # Already escaped in $headers
        }
        #
        $rangeReference = Get-CellRef ($script:excelData.Count-$startRowIndex) $headers.Length
        $replacementRules = @{
            'header.size' = $headers.Length
            'header.map' = $headerMap
            'rangeReference' = $rangeReference
            'tableColumn' = $tableColumn
            'dimension' = "<dimension ref=""A1:$rangeReference""/>"
        }
        # Function to build cell elements for a row
        function Build-CellElements {
            param (
                [int]$RowIndex,
                [object[]]$Values
            )
            $ret = ""
            for ($colIndex = 0; $colIndex -lt $Values.Length; $colIndex++){
                $cellRef  = Get-CellRef $RowIndex $colIndex
                $value = $Values[$colIndex]
                $ret += switch -Wildcard ($value.GetType().FullName) {
                    "System.*Int*"   { "<c r=""$cellRef"" s=""0"" t=""n""><v>$value</v></c>" }
                    "System.Decimal" { "<c r=""$cellRef"" s=""0"" t=""n""><v>$value</v></c>" }
                    "System.Double"  { "<c r=""$cellRef"" s=""0"" t=""n""><v>$value</v></c>" }
                    "System.Boolean" { "<c r=""$cellRef"" s=""0"" t=""n""><v>$([int]$value)</v></c>" }
                    "System.DateTime" {
                        "<c r=""$cellRef"" s=""1""><v>$(($value - $script:excelStartTime).TotalSeconds / 86400)</v></c>"
                    }
                    Default { "<c r=""$cellRef"" s=""0"" t=""str""><v>$(ReplaceSpecialCharacter ([string]$value))</v></c>" }
                }
            }
            # Excel row index starts from 1, so add 1 to $RowIndex
            "<row r=""$($RowIndex+1)"" spans=""1:$($values.Length)"" customFormat=""false"">$ret</row>"
        }

        # Create Zip file and add content
        try {
            if (Test-Path $Filename) {
                Remove-Item $Filename -ErrorAction Stop
            }
            Add-Type -AssemblyName System.IO.Compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem

            $workDir = "$PSScriptRoot/Template/"
            $sourceFolderFullPath = [System.IO.Path]::GetFullPath($workDir)
            $zipArchive = [System.IO.Compression.ZipFile]::Open($Filename, [System.IO.Compression.ZipArchiveMode]::Create)

            Get-ChildItem -Path $workDir -File -Recurse | Sort-Object FullName | ForEach-Object {
                $filePath = $_.fullname
                $fileContent = (Get-Content -LiteralPath $filePath -Encoding "utf8" | Out-String)
                $fileContent = [regex]::Replace($fileContent, "#{@([^}]+)}", {
                    param($match)
                    $replacementRules[$match.Groups[1].Value]
                })
                $relativePath = $_.FullName.Substring($sourceFolderFullPath.Length).TrimStart('\')
                $relativePath = $relativePath -replace "\\", "/"
                # ファイルを追加
                $entry = $zipArchive.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Optimal) # SmallestSize)
                $entryStream = $entry.Open()
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($fileContent)
                $entryStream.Write($bytes, 0, $bytes.Length)
                if ($filePath.EndsWith("sheet1.xml")){
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes((Build-CellElements 0 $headers))
                    $entryStream.Write($bytes, 0, $bytes.Length)
                    $rowIndex = 0
                    for ($row = $startRowIndex; $row -lt $script:excelData.Count; $row++){
                        $rowData = $script:excelData[$row]
                        $rowIndex++
                        # If the input was PSObject, we need to extract values in header order
                        if ($rowData -is [System.Array] -and $rowData[0] -is [psobject]) {
                            $orderedValues = @()
                            foreach ($headerProp in $headers) {
                                $orderedValues += $rowData[0].($headerProp)
                            }
                            $bytes = [System.Text.Encoding]::UTF8.GetBytes((Build-CellElements ($rowIndex) $orderedValues))
                        } else {
                            $bytes = [System.Text.Encoding]::UTF8.GetBytes((Build-CellElements ($rowIndex) $rowData))
                        }
                        $entryStream.Write($bytes, 0, $bytes.Length)
                    }
                    $closeContent = "</sheetData><tableParts count=""1""><tablePart r:id=""rId1""/></tableParts></worksheet>"
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($closeContent)
                    $entryStream.Write($bytes, 0, $bytes.Length)
                }
                $entryStream.Close()
            }
        } catch {
            Write-Error "Failed to export Excel file: $($_.Exception.Message),$($_.Exception.StackTrace)"
        } finally {
            if ($zipArchive) {
                $zipArchive.Dispose()
            }
        }
    }
}
