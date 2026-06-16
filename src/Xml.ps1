    #region XML and Namespace Helpers

    function Initialize-Namespaces {
        $script:NsMgr = New-Object System.Xml.XmlNamespaceManager([System.Xml.XmlDocument]::new().NameTable)
        $script:NsMgr.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        $script:NsMgr.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $script:NsMgr.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        $script:NsMgr.AddNamespace('x', 'http://schemas.openxmlformats.org/package/2006/relationships')
        $script:NsMgr.AddNamespace('asvg', 'http://schemas.microsoft.com/office/drawing/2016/SVG/main')
        $script:NsMgr.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
        $script:NsMgr.AddNamespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main')
        $script:NsMgr.AddNamespace('p159', 'http://schemas.microsoft.com/office/powerpoint/2015/09/main')
    }

    function Get-XmlDocument {
        param([string]$path)

        if (-not (Test-Path -LiteralPath $path)) {
            Write-Warning "XML file not found: $path"
            return $null
        }

        try {
            $xml = New-Object System.Xml.XmlDocument
            $xml.XmlResolver = $null
            $xml.PreserveWhitespace = $false
            $xml.Load($path)
            return $xml
        } catch {
            Write-Warning "Failed to load XML from $path : $_"
            return $null
        }
    }

    function Save-XmlDocument {
        param(
            [System.Xml.XmlDocument]$doc,
            [string]$path
        )

        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = [System.Text.UTF8Encoding]::new($false)  # No BOM
        $settings.Indent = $false
        $settings.NewLineHandling = [System.Xml.NewLineHandling]::None

        $writer = [System.Xml.XmlWriter]::Create($path, $settings)
        try {
            $doc.Save($writer)
            $writer.Flush()
        } finally {
            if ($writer) {
                $writer.Dispose()  # Dispose instead of Close for proper cleanup
            }
        }
    }

    #endregion
