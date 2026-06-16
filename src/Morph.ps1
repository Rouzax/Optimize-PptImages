    #region Morph Detection

    function Find-MorphPairs {
        param([System.Collections.Generic.List[ImageUsage]]$usages)
        
        $pairs = @{}
        $morphSlideCount = 0
        $totalPairsFound = 0
        $unmatchedImages = 0
        
        $slideUsages = $usages | Where-Object { $_.ContextType -eq [ContextType]::Slide } | 
            Group-Object -Property SlideNumber
        
        foreach ($group in $slideUsages) {
            $slideNum = $group.Name
            if ($slideNum -le 1) { continue }
            
            $currentSlide = $group.Group
            $currentMorph = $currentSlide | Where-Object { $_.IsMorphSlide } | Select-Object -First 1
            if (-not $currentMorph) { continue }
            
            $morphSlideCount++
            $prevSlideNum = [int]$slideNum - 1
            $prevSlide = $usages | Where-Object { 
                $_.ContextType -eq [ContextType]::Slide -and $_.SlideNumber -eq $prevSlideNum 
            }
            
            if (Test-VerboseMode) {
                Write-Verbose "Analyzing Morph transition: Slide $prevSlideNum -> Slide $slideNum"
            }
            
            # Match by physical image path
            $slideMatchCount = 0
            foreach ($curr in $currentSlide) {
                $match = $prevSlide | Where-Object { $_.ImagePhysicalPath -eq $curr.ImagePhysicalPath } | 
                    Select-Object -First 1
                
                if ($match) {
                    $curr.MorphPair = $match
                    $match.MorphPair = $curr
                    
                    $pairKey = "$($match.ImagePhysicalPath)|$prevSlideNum-$slideNum"
                    $pairs[$pairKey] = @($match, $curr)
                    $slideMatchCount++
                    $totalPairsFound++
                    
                    if (Test-VerboseMode) {
                        $fileName = Split-Path $match.ImagePhysicalPath -Leaf
                        Write-Verbose "  Linked pair: $fileName across slides $prevSlideNum <-> $slideNum"
                    }
                } else {
                    $unmatchedImages++
                }
            }
            
            if (Test-VerboseMode -and $slideMatchCount -eq 0) {
                Write-Verbose "  Warning: No matching images found between slides $prevSlideNum and $slideNum"
            }
        }
        
        if ($morphSlideCount -gt 0) {
            Write-Host "[LINK] Morph detection: $morphSlideCount transition(s), $totalPairsFound image pair(s) linked" -ForegroundColor Cyan
            
            if (Test-VerboseMode) {
                if ($unmatchedImages -gt 0) {
                    Write-Verbose "Note: $unmatchedImages image(s) on morph slides had no match on previous slide"
                }
                Write-Verbose "Morph pairs will be protected from conflicting crop operations"
            }
        }
        
        return $pairs
    }

    #endregion
