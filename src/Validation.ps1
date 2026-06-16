    # Load the DocumentFormat.OpenXml SDK if available. Returns $true when the
    # PresentationDocument type is usable, $false otherwise. The SDK is optional
    # tooling (only ImageMagick is a hard dependency), so absence is not an error.
    function Import-OpenXmlSdk {
        if ('DocumentFormat.OpenXml.Packaging.PresentationDocument' -as [type]) {
            return $true
        }
        try {
            $fwPkg = Get-Package -Name DocumentFormat.OpenXml.Framework -ErrorAction Stop
            $mainPkg = Get-Package -Name DocumentFormat.OpenXml -ErrorAction Stop
            $fwDll = Get-ChildItem (Split-Path $fwPkg.Source -Parent) -Recurse -Filter '*.dll' |
                Where-Object { $_.FullName -match 'netstandard' } | Select-Object -First 1
            $mainDll = Get-ChildItem (Split-Path $mainPkg.Source -Parent) -Recurse -Filter '*.dll' |
                Where-Object { $_.FullName -match 'netstandard' } | Select-Object -First 1
            if (-not $fwDll -or -not $mainDll) { return $false }
            Add-Type -Path $fwDll.FullName
            Add-Type -Path $mainDll.FullName
            return [bool]('DocumentFormat.OpenXml.Packaging.PresentationDocument' -as [type])
        } catch {
            Write-Verbose "Open XML SDK not available: $_"
            return $false
        }
    }

    # Optional post-repack validation (opt-in via -ValidateOutput). Runs the Open XML
    # SDK validator against the repacked file and reports any errors. Non-fatal: it
    # surfaces problems but does not abort the run or discard the output.
    function Test-OutputValidity {
        param([string]$pptxPath)

        Write-Host "`n[VALIDATE] Validating output with Open XML SDK..." -ForegroundColor Yellow

        if (-not (Import-OpenXmlSdk)) {
            Write-Warning "Open XML SDK (DocumentFormat.OpenXml) not found; skipping validation. Install it to enable -ValidateOutput."
            return
        }

        $doc = $null
        try {
            $doc = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Open($pptxPath, $false)
            $validator = [DocumentFormat.OpenXml.Validation.OpenXmlValidator]::new()
            $errors = @($validator.Validate($doc))

            if ($errors.Count -eq 0) {
                Write-Host "[OK] Output is valid (0 errors)" -ForegroundColor Green
            } else {
                Write-Warning "Output validation reported $($errors.Count) issue(s):"
                foreach ($e in ($errors | Select-Object -First 10)) {
                    Write-Warning "  [$($e.ErrorType)] $($e.Description)"
                }
                if ($errors.Count -gt 10) {
                    Write-Warning "  ... and $($errors.Count - 10) more"
                }
            }
        } catch {
            Write-Warning "Validation could not run: $_"
        } finally {
            if ($doc) { $doc.Dispose() }
        }
    }

    #endregion
