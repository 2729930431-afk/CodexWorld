$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$script:WordApplication = $null
$script:ObjectHandles = @{}
$script:NextHandle = 1

function Has-Arg {
    param($Arguments, [string]$Name)
    return ($null -ne $Arguments -and $null -ne $Arguments.PSObject.Properties[$Name])
}

function Get-Arg {
    param($Arguments, [string]$Name, $Default = $null)
    if (Has-Arg $Arguments $Name) {
        return $Arguments.PSObject.Properties[$Name].Value
    }
    return $Default
}

function Has-Property {
    param($Object, [string]$Name)
    return ($null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name])
}

function Get-WordApplication {
    if ($null -ne $script:WordApplication) {
        try {
            $null = $script:WordApplication.Version
            return $script:WordApplication
        } catch {
            $script:WordApplication = $null
        }
    }

    try {
        $script:WordApplication = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    } catch {
        throw "Microsoft Word is not running in this Windows user session. Open desktop Word and the target document first."
    }
    return $script:WordApplication
}

function Get-DocumentFullName {
    param($Document)
    try {
        if ([string]::IsNullOrWhiteSpace([string]$Document.Path)) {
            return [string]$Document.Name
        }
        return [string]$Document.FullName
    } catch {
        return [string]$Document.Name
    }
}

function Resolve-Document {
    param($Arguments)
    $app = Get-WordApplication
    if ($app.Documents.Count -lt 1) {
        throw "Word is running, but no document is open."
    }

    $selector = [string](Get-Arg $Arguments "document" "")
    if ([string]::IsNullOrWhiteSpace($selector)) {
        return $app.ActiveDocument
    }

    $matches = @()
    for ($i = 1; $i -le $app.Documents.Count; $i++) {
        $doc = $app.Documents.Item($i)
        $fullName = Get-DocumentFullName $doc
        if ($doc.Name -ieq $selector -or $fullName -ieq $selector) {
            $matches += $doc
        }
    }

    if ($matches.Count -eq 1) {
        return $matches[0]
    }
    if ($matches.Count -gt 1) {
        throw "More than one open document matches '$selector'. Pass the exact full path."
    }
    throw "No open Word document matches '$selector'. Call word_status to list open documents."
}

function Get-DocumentInfo {
    param($Document, $Application)
    $path = ""
    try { $path = [string]$Document.Path } catch {}
    $fullName = Get-DocumentFullName $Document
    $active = $false
    try { $active = ((Get-DocumentFullName $Application.ActiveDocument) -ieq $fullName) } catch {}

    return [ordered]@{
        name = [string]$Document.Name
        full_name = $fullName
        path = $path
        active = $active
        saved = [bool]$Document.Saved
        read_only = [bool]$Document.ReadOnly
        track_changes = [bool]$Document.TrackRevisions
        compatibility_mode = [int]$Document.CompatibilityMode
        characters = [int]$Document.Characters.Count
        words = [int]$Document.Words.Count
        paragraphs = [int]$Document.Paragraphs.Count
        tables = [int]$Document.Tables.Count
        comments = [int]$Document.Comments.Count
        revisions = [int]$Document.Revisions.Count
    }
}

function Activate-Document {
    param($Document)
    try { $Document.Activate() } catch {}
}

function Get-TargetRange {
    param($Document, $Arguments, [string]$DefaultScope = "selection")
    $scope = [string](Get-Arg $Arguments "scope" $DefaultScope)
    $app = Get-WordApplication

    switch ($scope.ToLowerInvariant()) {
        "document" {
            return $Document.Content.Duplicate
        }
        "selection" {
            Activate-Document $Document
            return $app.Selection.Range.Duplicate
        }
        "current_paragraph" {
            Activate-Document $Document
            if ($app.Selection.Paragraphs.Count -lt 1) {
                throw "The current selection does not contain a paragraph."
            }
            return $app.Selection.Paragraphs.Item(1).Range.Duplicate
        }
        "range" {
            if (-not (Has-Arg $Arguments "start") -or -not (Has-Arg $Arguments "end")) {
                throw "scope='range' requires integer start and end offsets."
            }
            $start = [int](Get-Arg $Arguments "start")
            $end = [int](Get-Arg $Arguments "end")
            $maxEnd = [int]$Document.Content.End
            if ($start -lt 0 -or $end -lt $start -or $end -gt $maxEnd) {
                throw "Invalid Word range [$start,$end). Document bounds are [0,$maxEnd)."
            }
            return $Document.Range($start, $end)
        }
        default {
            throw "Unsupported scope '$scope'. Use document, selection, current_paragraph, or range."
        }
    }
}

function Invoke-WithTrackChanges {
    param($Document, $Arguments, [scriptblock]$Action)
    $override = Has-Arg $Arguments "track_changes"
    $previous = [bool]$Document.TrackRevisions
    if ($override) {
        $Document.TrackRevisions = [bool](Get-Arg $Arguments "track_changes")
    }
    try {
        return & $Action
    } finally {
        if ($override) {
            $Document.TrackRevisions = $previous
        }
    }
}

function Save-InPlace {
    param($Document, $Arguments)
    $saveAfter = [bool](Get-Arg $Arguments "save_after" $true)
    if (-not $saveAfter) {
        return [ordered]@{ requested = $false; saved = [bool]$Document.Saved; full_name = Get-DocumentFullName $Document }
    }
    if ([bool]$Document.ReadOnly) {
        throw "The document is read-only and cannot be saved in place."
    }
    if ([string]::IsNullOrWhiteSpace([string]$Document.Path)) {
        return [ordered]@{
            requested = $true
            saved = $false
            full_name = [string]$Document.Name
            reason = "The document has never been saved. World did not invent a path or create a copy."
        }
    }
    $Document.Save()
    return [ordered]@{ requested = $true; saved = [bool]$Document.Saved; full_name = Get-DocumentFullName $Document }
}

function Limit-Text {
    param([string]$Text, [int]$Maximum)
    if ($null -eq $Text) { return "" }
    if ($Maximum -lt 1) { $Maximum = 1 }
    if ($Text.Length -le $Maximum) { return $Text }
    return $Text.Substring(0, $Maximum)
}

function Clean-WordText {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return $Text.Replace([char]7, "`t")
}

function Get-NearestTableContext {
    param($Document, [int]$Position, [int]$RequestedIndex = 0)

    if ($RequestedIndex -gt 0) {
        if ($RequestedIndex -gt $Document.Tables.Count) {
            throw "Invalid template_table_index $RequestedIndex."
        }
        return [ordered]@{
            index = $RequestedIndex
            table = $Document.Tables.Item($RequestedIndex)
        }
    }

    $nearest = $null
    $nearestIndex = 0
    $nearestDistance = [int]::MaxValue
    for ($i = 1; $i -le $Document.Tables.Count; $i++) {
        $candidate = $Document.Tables.Item($i)
        $start = [int]$candidate.Range.Start
        $end = [int]$candidate.Range.End
        if ($Position -ge $start -and $Position -le $end) {
            continue
        }
        $distance = if ($Position -lt $start) { $start - $Position } else { $Position - $end }
        if ($distance -lt $nearestDistance) {
            $nearest = $candidate
            $nearestIndex = $i
            $nearestDistance = $distance
        }
    }
    if ($null -eq $nearest) { return $null }
    return [ordered]@{ index = $nearestIndex; table = $nearest }
}

function Copy-RangeVisualFormat {
    param($SourceRange, $TargetRange)

    try { $TargetRange.Font.Name = $SourceRange.Font.Name } catch {}
    try { $TargetRange.Font.NameFarEast = $SourceRange.Font.NameFarEast } catch {}
    try { $TargetRange.Font.Size = $SourceRange.Font.Size } catch {}
    try { $TargetRange.Font.Bold = $SourceRange.Font.Bold } catch {}
    try { $TargetRange.Font.Italic = $SourceRange.Font.Italic } catch {}
    try { $TargetRange.Font.Underline = $SourceRange.Font.Underline } catch {}
    try { $TargetRange.Font.Color = $SourceRange.Font.Color } catch {}
    try { $TargetRange.ParagraphFormat.Alignment = $SourceRange.ParagraphFormat.Alignment } catch {}
    try { $TargetRange.ParagraphFormat.LeftIndent = $SourceRange.ParagraphFormat.LeftIndent } catch {}
    try { $TargetRange.ParagraphFormat.RightIndent = $SourceRange.ParagraphFormat.RightIndent } catch {}
    try { $TargetRange.ParagraphFormat.FirstLineIndent = $SourceRange.ParagraphFormat.FirstLineIndent } catch {}
    try { $TargetRange.ParagraphFormat.SpaceBefore = $SourceRange.ParagraphFormat.SpaceBefore } catch {}
    try { $TargetRange.ParagraphFormat.SpaceAfter = $SourceRange.ParagraphFormat.SpaceAfter } catch {}
    try { $TargetRange.ParagraphFormat.LineSpacing = $SourceRange.ParagraphFormat.LineSpacing } catch {}
    try { $TargetRange.Shading.BackgroundPatternColor = $SourceRange.Shading.BackgroundPatternColor } catch {}
    try { $TargetRange.Shading.ForegroundPatternColor = $SourceRange.Shading.ForegroundPatternColor } catch {}
    try { $TargetRange.Shading.Texture = $SourceRange.Shading.Texture } catch {}
}

function Copy-CellVisualFormat {
    param($SourceCell, $TargetCell)

    Copy-RangeVisualFormat $SourceCell.Range $TargetCell.Range
    try { $TargetCell.Shading.BackgroundPatternColor = $SourceCell.Shading.BackgroundPatternColor } catch {}
    try { $TargetCell.Shading.ForegroundPatternColor = $SourceCell.Shading.ForegroundPatternColor } catch {}
    try { $TargetCell.Shading.Texture = $SourceCell.Shading.Texture } catch {}
    try { $TargetCell.VerticalAlignment = $SourceCell.VerticalAlignment } catch {}
}

function Copy-TableContextFormat {
    param($SourceTable, $TargetTable)

    try { $TargetTable.Style = [string]$SourceTable.Style.NameLocal } catch {}
    foreach ($property in @(
        "ApplyStyleHeadingRows",
        "ApplyStyleFirstColumn",
        "ApplyStyleLastColumn",
        "ApplyStyleRowBands",
        "ApplyStyleColumnBands",
        "LeftPadding",
        "RightPadding",
        "TopPadding",
        "BottomPadding",
        "Spacing"
    )) {
        try { $TargetTable.$property = $SourceTable.$property } catch {}
    }

    foreach ($borderType in @(-1, -2, -3, -4, -5, -6)) {
        try {
            $sourceBorder = $SourceTable.Borders.Item($borderType)
            $targetBorder = $TargetTable.Borders.Item($borderType)
            foreach ($property in @("LineStyle", "LineWidth", "Color", "Visible")) {
                try { $targetBorder.$property = $sourceBorder.$property } catch {}
            }
        } catch {}
    }

    $sourceTotalWidth = 0.0
    for ($column = 1; $column -le $SourceTable.Columns.Count; $column++) {
        try { $sourceTotalWidth += [double]$SourceTable.Columns.Item($column).Width } catch {}
    }
    $targetTotalWidth = 0.0
    for ($column = 1; $column -le $TargetTable.Columns.Count; $column++) {
        try { $targetTotalWidth += [double]$TargetTable.Columns.Item($column).Width } catch {}
    }
    if ($sourceTotalWidth -gt 0 -and $targetTotalWidth -gt 0) {
        try { $TargetTable.AllowAutoFit = $false } catch {}
        for ($column = 1; $column -le $TargetTable.Columns.Count; $column++) {
            try {
                $currentWidth = [double]$TargetTable.Columns.Item($column).Width
                $scaledWidth = [single]($sourceTotalWidth * ($currentWidth / $targetTotalWidth))
                $TargetTable.Columns.Item($column).Width = $scaledWidth
            } catch {}
        }
    } else {
        try { $TargetTable.AllowAutoFit = $SourceTable.AllowAutoFit } catch {}
    }

    $sourceRows = [int]$SourceTable.Rows.Count
    $sourceColumns = [int]$SourceTable.Columns.Count
    $targetRows = [int]$TargetTable.Rows.Count
    $targetColumns = [int]$TargetTable.Columns.Count
    for ($row = 1; $row -le $targetRows; $row++) {
        $sourceRow = if ($row -eq 1 -or $sourceRows -lt 2) {
            1
        } else {
            2 + (($row - 2) % ($sourceRows - 1))
        }
        for ($column = 1; $column -le $targetColumns; $column++) {
            $sourceColumn = [int][Math]::Min(
                $sourceColumns,
                [Math]::Floor((($column - 1) * $sourceColumns) / $targetColumns) + 1
            )
            try {
                $sourceCell = $SourceTable.Cell($sourceRow, $sourceColumn)
                $targetCell = $TargetTable.Cell($row, $column)
                Copy-CellVisualFormat $sourceCell $targetCell
            } catch {}
        }
    }

    $styleName = ""
    try { $styleName = [string]$SourceTable.Style.NameLocal } catch {}
    $headerColor = $null
    $bodyColor = $null
    $sourceHeaderColor = $null
    $sourceBodyColor = $null
    try { $sourceHeaderColor = [int]$SourceTable.Cell(1, 1).Shading.BackgroundPatternColor } catch {}
    if ($sourceRows -gt 1) {
        try { $sourceBodyColor = [int]$SourceTable.Cell(2, 1).Shading.BackgroundPatternColor } catch {}
    }
    try { $headerColor = [int]$TargetTable.Cell(1, 1).Shading.BackgroundPatternColor } catch {}
    if ($targetRows -gt 1) {
        try { $bodyColor = [int]$TargetTable.Cell(2, 1).Shading.BackgroundPatternColor } catch {}
    }
    return [ordered]@{
        style = $styleName
        total_width = [single]$sourceTotalWidth
        source_header_color = $sourceHeaderColor
        source_body_color = $sourceBodyColor
        header_color = $headerColor
        body_color = $bodyColor
    }
}

function Set-TableText {
    param($Table, [string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $lines = [regex]::Split($Text, "\r\n|\n|\r")
    $written = 0
    for ($row = 1; $row -le [Math]::Min($Table.Rows.Count, $lines.Count); $row++) {
        $cells = @($lines[$row - 1].Split("`t"))
        for ($column = 1; $column -le [Math]::Min($Table.Columns.Count, $cells.Count); $column++) {
            $Table.Cell($row, $column).Range.Text = [string]$cells[$column - 1]
            $written++
        }
    }
    return $written
}

function Convert-Rgb {
    param([string]$Hex)
    $value = $Hex.Trim().TrimStart("#")
    if ($value.Length -ne 6) {
        throw "font.color must be a six-digit RGB value such as #185ABD."
    }
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return ($r + ($g * 256) + ($b * 65536))
}

function Get-ComProperty {
    param($Object, [string]$Name)
    $flags = [Reflection.BindingFlags]::GetProperty -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::IgnoreCase
    return $Object.GetType().InvokeMember($Name, $flags, $null, $Object, @())
}

function Set-ComProperty {
    param($Object, [string]$Name, $Value)
    $flags = [Reflection.BindingFlags]::SetProperty -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::IgnoreCase
    $null = $Object.GetType().InvokeMember($Name, $flags, $null, $Object, @($Value))
}

function Invoke-ComMethod {
    param($Object, [string]$Name, [object[]]$Arguments)
    $flags = [Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::IgnoreCase
    return $Object.GetType().InvokeMember($Name, $flags, $null, $Object, $Arguments)
}

function Resolve-ComPathParent {
    param($Root, [string]$Path)
    $parts = @($Path.Split(".", [StringSplitOptions]::RemoveEmptyEntries))
    if ($parts.Count -lt 1) { throw "A non-empty member path is required." }
    $current = $Root
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        $current = Get-ComProperty $current $parts[$i]
    }
    return [ordered]@{ parent = $current; member = $parts[$parts.Count - 1] }
}

function Register-ComHandle {
    param($Object)
    $id = "h$($script:NextHandle)"
    $script:NextHandle++
    $script:ObjectHandles[$id] = $Object
    return $id
}

function Convert-ComResult {
    param($Value)
    if ($null -eq $Value) {
        return $null
    }
    if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $handle = Register-ComHandle $Value
        $summary = [ordered]@{ kind = "com_object"; handle = $handle }
        foreach ($propertyName in @("Name", "FullName", "Count", "Text", "Start", "End", "Type")) {
            try {
                $propertyValue = Get-ComProperty $Value $propertyName
                if ($null -eq $propertyValue -or -not [Runtime.InteropServices.Marshal]::IsComObject($propertyValue)) {
                    $summary[$propertyName.ToLowerInvariant()] = $propertyValue
                }
            } catch {}
        }
        return $summary
    }
    if ($Value -is [DateTime]) {
        return $Value.ToString("o")
    }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        $count = 0
        foreach ($item in $Value) {
            if ($count -ge 200) { break }
            $items += Convert-ComResult $item
            $count++
        }
        return $items
    }
    return $Value
}

function Get-WordStatus {
    $app = Get-WordApplication
    $documents = @()
    for ($i = 1; $i -le $app.Documents.Count; $i++) {
        $documents += Get-DocumentInfo $app.Documents.Item($i) $app
    }
    $selection = $null
    try {
        $selection = [ordered]@{
            document = Get-DocumentFullName $app.Selection.Document
            start = [int]$app.Selection.Start
            end = [int]$app.Selection.End
            text = Limit-Text ([string]$app.Selection.Text) 4000
        }
    } catch {}

    return [ordered]@{
        connected = $true
        application = "Microsoft Word"
        version = [string]$app.Version
        visible = [bool]$app.Visible
        document_count = [int]$app.Documents.Count
        active_document = if ($app.Documents.Count -gt 0) { Get-DocumentFullName $app.ActiveDocument } else { $null }
        selection = $selection
        documents = $documents
        object_handles = [int]$script:ObjectHandles.Count
    }
}

function Invoke-WordTool {
    param([string]$Name, $Arguments)

    switch ($Name) {
        "word_status" {
            return Get-WordStatus
        }

        "word_read" {
            $doc = Resolve-Document $Arguments
            $range = Get-TargetRange $doc $Arguments "selection"
            $maximum = [int](Get-Arg $Arguments "max_chars" 20000)
            $text = Clean-WordText ([string]$range.Text)
            return [ordered]@{
                document = Get-DocumentFullName $doc
                scope = [string](Get-Arg $Arguments "scope" "selection")
                start = [int]$range.Start
                end = [int]$range.End
                truncated = ($text.Length -gt $maximum)
                text = Limit-Text $text $maximum
            }
        }

        "word_edit" {
            $doc = Resolve-Document $Arguments
            $range = Get-TargetRange $doc $Arguments "selection"
            $mode = [string](Get-Arg $Arguments "mode" "replace")
            $text = [string](Get-Arg $Arguments "text" "")
            $expected = Get-Arg $Arguments "expected_text" $null
            if ($null -ne $expected -and [string]$range.Text -cne [string]$expected) {
                throw "The target range no longer matches expected_text. Re-read the range before editing."
            }
            $beforeStart = [int]$range.Start
            $beforeEnd = [int]$range.End

            Invoke-WithTrackChanges $doc $Arguments {
                switch ($mode.ToLowerInvariant()) {
                    "replace" { $range.Text = $text }
                    "insert_before" { $range.InsertBefore($text) }
                    "insert_after" { $range.InsertAfter($text) }
                    default { throw "Unsupported edit mode '$mode'." }
                }
            } | Out-Null

            $save = Save-InPlace $doc $Arguments
            return [ordered]@{
                document = Get-DocumentFullName $doc
                mode = $mode
                original_range = [ordered]@{ start = $beforeStart; end = $beforeEnd }
                resulting_range = [ordered]@{ start = [int]$range.Start; end = [int]$range.End }
                characters_written = $text.Length
                save = $save
            }
        }

        "word_find_replace" {
            $doc = Resolve-Document $Arguments
            $range = Get-TargetRange $doc $Arguments "document"
            $findText = [string](Get-Arg $Arguments "find_text")
            $replaceText = [string](Get-Arg $Arguments "replace_text" "")
            if ([string]::IsNullOrEmpty($findText)) {
                throw "find_text must not be empty."
            }
            $replaceMode = [string](Get-Arg $Arguments "replace" "all")
            $replaceCode = if ($replaceMode -ieq "one") { 1 } else { 2 }
            $matchCase = [bool](Get-Arg $Arguments "match_case" $false)
            $wholeWord = [bool](Get-Arg $Arguments "whole_word" $false)
            $wildcards = [bool](Get-Arg $Arguments "wildcards" $false)
            $performed = $false

            $performed = Invoke-WithTrackChanges $doc $Arguments {
                $find = $range.Find
                $find.ClearFormatting()
                $find.Replacement.ClearFormatting()
                return [bool]$find.Execute($findText, $matchCase, $wholeWord, $wildcards, $false, $false, $true, 0, $false, $replaceText, $replaceCode)
            }

            $save = Save-InPlace $doc $Arguments
            return [ordered]@{
                document = Get-DocumentFullName $doc
                find_text = $findText
                replace_text = $replaceText
                replace = $replaceMode
                performed = $performed
                save = $save
            }
        }

        "word_format" {
            $doc = Resolve-Document $Arguments
            $range = Get-TargetRange $doc $Arguments "selection"
            $applied = @()
            $style = [string](Get-Arg $Arguments "style" "")
            if (-not [string]::IsNullOrWhiteSpace($style)) {
                $range.Style = $style
                $applied += "style"
            }
            $font = Get-Arg $Arguments "font" $null
            if ($null -ne $font) {
                if (Has-Property $font "name") { $range.Font.Name = [string]$font.name; $applied += "font.name" }
                if (Has-Property $font "size") { $range.Font.Size = [single]$font.size; $applied += "font.size" }
                if (Has-Property $font "bold") { $range.Font.Bold = if ([bool]$font.bold) { -1 } else { 0 }; $applied += "font.bold" }
                if (Has-Property $font "italic") { $range.Font.Italic = if ([bool]$font.italic) { -1 } else { 0 }; $applied += "font.italic" }
                if (Has-Property $font "underline") { $range.Font.Underline = if ([bool]$font.underline) { 1 } else { 0 }; $applied += "font.underline" }
                if (Has-Property $font "color") { $range.Font.Color = Convert-Rgb ([string]$font.color); $applied += "font.color" }
            }
            $paragraph = Get-Arg $Arguments "paragraph" $null
            if ($null -ne $paragraph) {
                $format = $range.ParagraphFormat
                if (Has-Property $paragraph "alignment") {
                    $map = @{ left = 0; center = 1; right = 2; justify = 3 }
                    $alignment = [string]$paragraph.alignment
                    if (-not $map.ContainsKey($alignment.ToLowerInvariant())) { throw "Unsupported paragraph alignment '$alignment'." }
                    $format.Alignment = $map[$alignment.ToLowerInvariant()]
                    $applied += "paragraph.alignment"
                }
                if (Has-Property $paragraph "left_indent") { $format.LeftIndent = [single]$paragraph.left_indent; $applied += "paragraph.left_indent" }
                if (Has-Property $paragraph "right_indent") { $format.RightIndent = [single]$paragraph.right_indent; $applied += "paragraph.right_indent" }
                if (Has-Property $paragraph "first_line_indent") { $format.FirstLineIndent = [single]$paragraph.first_line_indent; $applied += "paragraph.first_line_indent" }
                if (Has-Property $paragraph "space_before") { $format.SpaceBefore = [single]$paragraph.space_before; $applied += "paragraph.space_before" }
                if (Has-Property $paragraph "space_after") { $format.SpaceAfter = [single]$paragraph.space_after; $applied += "paragraph.space_after" }
                if (Has-Property $paragraph "line_spacing") { $format.LineSpacing = [single]$paragraph.line_spacing; $applied += "paragraph.line_spacing" }
            }
            if ($applied.Count -eq 0) {
                throw "No style, font, or paragraph properties were supplied."
            }
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; start = [int]$range.Start; end = [int]$range.End; applied = $applied; save = $save }
        }

        "word_comment" {
            $doc = Resolve-Document $Arguments
            $action = [string](Get-Arg $Arguments "action" "list")
            if ($action -ieq "list") {
                $comments = @()
                for ($i = 1; $i -le $doc.Comments.Count; $i++) {
                    $comment = $doc.Comments.Item($i)
                    $done = $null
                    try { $done = [bool]$comment.Done } catch {}
                    $comments += [ordered]@{
                        index = $i
                        author = [string]$comment.Author
                        initials = [string]$comment.Initial
                        date = ([DateTime]$comment.Date).ToString("o")
                        text = [string]$comment.Range.Text
                        scope_text = Limit-Text (Clean-WordText ([string]$comment.Scope.Text)) 1000
                        done = $done
                    }
                }
                return [ordered]@{ document = Get-DocumentFullName $doc; comments = $comments }
            }
            if ($action -ieq "add") {
                $range = Get-TargetRange $doc $Arguments "selection"
                $text = [string](Get-Arg $Arguments "text")
                if ([string]::IsNullOrWhiteSpace($text)) { throw "text is required when adding a comment." }
                $comment = $doc.Comments.Add($range, $text)
                $save = Save-InPlace $doc $Arguments
                return [ordered]@{ document = Get-DocumentFullName $doc; index = [int]$comment.Index; save = $save }
            }
            $index = [int](Get-Arg $Arguments "index")
            if ($index -lt 1 -or $index -gt $doc.Comments.Count) { throw "Invalid comment index $index." }
            $target = $doc.Comments.Item($index)
            if ($action -ieq "delete") {
                $target.Delete()
            } elseif ($action -ieq "resolve") {
                $target.Done = $true
            } elseif ($action -ieq "reopen") {
                $target.Done = $false
            } else {
                throw "Unsupported comment action '$action'."
            }
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; action = $action; index = $index; save = $save }
        }

        "word_track_changes" {
            $doc = Resolve-Document $Arguments
            $action = [string](Get-Arg $Arguments "action" "get")
            if ($action -ieq "get") {
                return [ordered]@{ document = Get-DocumentFullName $doc; enabled = [bool]$doc.TrackRevisions; revisions = [int]$doc.Revisions.Count }
            }
            if ($action -ine "set") { throw "Unsupported track-changes action '$action'." }
            if (-not (Has-Arg $Arguments "enabled")) { throw "enabled is required for action='set'." }
            $doc.TrackRevisions = [bool](Get-Arg $Arguments "enabled")
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; enabled = [bool]$doc.TrackRevisions; save = $save }
        }

        "word_review_changes" {
            $doc = Resolve-Document $Arguments
            $action = [string](Get-Arg $Arguments "action" "list")
            $scope = [string](Get-Arg $Arguments "scope" "document")
            $revisions = if ($scope -ieq "selection") { (Get-TargetRange $doc $Arguments "selection").Revisions } else { $doc.Revisions }
            if ($action -ieq "list") {
                $items = @()
                $limit = [int](Get-Arg $Arguments "limit" 200)
                for ($i = 1; $i -le [Math]::Min($revisions.Count, $limit); $i++) {
                    $revision = $revisions.Item($i)
                    $items += [ordered]@{
                        index = $i
                        type = [int]$revision.Type
                        author = [string]$revision.Author
                        date = ([DateTime]$revision.Date).ToString("o")
                        start = [int]$revision.Range.Start
                        end = [int]$revision.Range.End
                        text = Limit-Text (Clean-WordText ([string]$revision.Range.Text)) 1000
                    }
                }
                return [ordered]@{ document = Get-DocumentFullName $doc; count = [int]$revisions.Count; revisions = $items }
            }
            if ($revisions.Count -lt 1) {
                return [ordered]@{ document = Get-DocumentFullName $doc; action = $action; changed = 0; save = Save-InPlace $doc $Arguments }
            }
            $before = [int]$revisions.Count
            switch ($action.ToLowerInvariant()) {
                "accept_all" { $revisions.AcceptAll() }
                "reject_all" { $revisions.RejectAll() }
                "accept_next" { $revisions.Item(1).Accept() }
                "reject_next" { $revisions.Item(1).Reject() }
                default { throw "Unsupported revision action '$action'." }
            }
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; action = $action; changed = if ($action -like "*_all") { $before } else { 1 }; save = $save }
        }

        "word_table" {
            $doc = Resolve-Document $Arguments
            $action = [string](Get-Arg $Arguments "action" "list")
            if ($action -ieq "list") {
                $tables = @()
                for ($i = 1; $i -le $doc.Tables.Count; $i++) {
                    $table = $doc.Tables.Item($i)
                    $tables += [ordered]@{
                        index = $i
                        rows = [int]$table.Rows.Count
                        columns = [int]$table.Columns.Count
                        start = [int]$table.Range.Start
                        end = [int]$table.Range.End
                        text = Limit-Text (Clean-WordText ([string]$table.Range.Text)) 4000
                    }
                }
                return [ordered]@{ document = Get-DocumentFullName $doc; tables = $tables }
            }
            if ($action -ieq "insert") {
                $range = Get-TargetRange $doc $Arguments "selection"
                $range.Collapse(1)
                $insertPosition = [int]$range.Start
                $rows = [int](Get-Arg $Arguments "rows" 2)
                $columns = [int](Get-Arg $Arguments "columns" 2)
                if ($rows -lt 1 -or $columns -lt 1) { throw "rows and columns must be positive." }
                $matchContext = [bool](Get-Arg $Arguments "match_context" $true)
                $templateIndex = [int](Get-Arg $Arguments "template_table_index" 0)
                $context = if ($matchContext) { Get-NearestTableContext $doc $insertPosition $templateIndex } else { $null }
                $table = $doc.Tables.Add($range, $rows, $columns)
                if ([bool](Get-Arg $Arguments "auto_fit" $true)) { $table.AutoFitBehavior(1) }
                $cellsWritten = Set-TableText $table ([string](Get-Arg $Arguments "text" ""))
                $format = $null
                if ($null -ne $context) {
                    $format = Copy-TableContextFormat $context.table $table
                }
                $tableIndex = 0
                for ($i = 1; $i -le $doc.Tables.Count; $i++) {
                    if ([int]$doc.Tables.Item($i).Range.Start -eq [int]$table.Range.Start) {
                        $tableIndex = $i
                        break
                    }
                }
                $save = Save-InPlace $doc $Arguments
                return [ordered]@{
                    document = Get-DocumentFullName $doc
                    index = $tableIndex
                    rows = $rows
                    columns = $columns
                    cells_written = $cellsWritten
                    context_matched = ($null -ne $context)
                    template_table_index = if ($null -ne $context) { [int]$context.index } else { $null }
                    format = $format
                    save = $save
                }
            }
            $tableIndex = [int](Get-Arg $Arguments "table_index")
            if ($tableIndex -lt 1 -or $tableIndex -gt $doc.Tables.Count) { throw "Invalid table_index $tableIndex." }
            $table = $doc.Tables.Item($tableIndex)
            if ($action -ieq "get_cell") {
                $row = [int](Get-Arg $Arguments "row")
                $column = [int](Get-Arg $Arguments "column")
                $cell = $table.Cell($row, $column)
                return [ordered]@{ document = Get-DocumentFullName $doc; table_index = $tableIndex; row = $row; column = $column; text = (Clean-WordText ([string]$cell.Range.Text)).TrimEnd("`r", "`t") }
            }
            switch ($action.ToLowerInvariant()) {
                "set_cell" {
                    $row = [int](Get-Arg $Arguments "row")
                    $column = [int](Get-Arg $Arguments "column")
                    $table.Cell($row, $column).Range.Text = [string](Get-Arg $Arguments "text" "")
                }
                "add_row" { $null = $table.Rows.Add() }
                "add_column" { $null = $table.Columns.Add() }
                "delete" { $table.Delete() }
                default { throw "Unsupported table action '$action'." }
            }
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; action = $action; table_index = $tableIndex; save = $save }
        }

        "word_image" {
            $doc = Resolve-Document $Arguments
            $path = [string](Get-Arg $Arguments "file_path")
            if (-not [IO.Path]::IsPathRooted($path)) { throw "file_path must be absolute." }
            if (-not [IO.File]::Exists($path)) { throw "Image file does not exist: $path" }
            $range = Get-TargetRange $doc $Arguments "selection"
            $range.Collapse(1)
            $link = [bool](Get-Arg $Arguments "link_to_file" $false)
            $saveWithDocument = [bool](Get-Arg $Arguments "save_with_document" $true)
            $shape = $doc.InlineShapes.AddPicture($path, $link, $saveWithDocument, $range)
            if (Has-Arg $Arguments "width") { $shape.Width = [single](Get-Arg $Arguments "width") }
            if (Has-Arg $Arguments "height") { $shape.Height = [single](Get-Arg $Arguments "height") }
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; type = [int]$shape.Type; width = [single]$shape.Width; height = [single]$shape.Height; save = $save }
        }

        "word_document" {
            $app = Get-WordApplication
            $action = [string](Get-Arg $Arguments "action" "list")
            if ($action -ieq "list") { return Get-WordStatus }
            if ($action -ieq "open") {
                $path = [string](Get-Arg $Arguments "file_path")
                if (-not [IO.Path]::IsPathRooted($path) -or -not [IO.File]::Exists($path)) { throw "file_path must be an existing absolute path." }
                $doc = $app.Documents.Open($path, $false, [bool](Get-Arg $Arguments "read_only" $false))
                return Get-DocumentInfo $doc $app
            }
            if ($action -ieq "new") {
                $doc = $app.Documents.Add()
                return Get-DocumentInfo $doc $app
            }
            if ($action -ieq "save_all") {
                $items = @()
                for ($i = 1; $i -le $app.Documents.Count; $i++) {
                    $doc = $app.Documents.Item($i)
                    $items += Save-InPlace $doc $Arguments
                }
                return [ordered]@{ action = "save_all"; documents = $items }
            }
            $doc = Resolve-Document $Arguments
            switch ($action.ToLowerInvariant()) {
                "activate" { $doc.Activate(); return Get-DocumentInfo $doc $app }
                "save" { return [ordered]@{ document = Get-DocumentFullName $doc; save = Save-InPlace $doc $Arguments } }
                "export_pdf" {
                    $path = [string](Get-Arg $Arguments "file_path")
                    if (-not [IO.Path]::IsPathRooted($path)) { throw "file_path must be absolute." }
                    $doc.ExportAsFixedFormat($path, 17)
                    return [ordered]@{ document = Get-DocumentFullName $doc; exported = $path }
                }
                "print" {
                    $doc.PrintOut()
                    return [ordered]@{ document = Get-DocumentFullName $doc; print_submitted = $true }
                }
                "close" {
                    $saveChanges = [bool](Get-Arg $Arguments "save_changes" $true)
                    $name = Get-DocumentFullName $doc
                    $doc.Close($(if ($saveChanges) { -1 } else { 0 }))
                    return [ordered]@{ document = $name; closed = $true; saved = $saveChanges }
                }
                default { throw "Unsupported document action '$action'." }
            }
        }

        "word_undo_redo" {
            $action = [string](Get-Arg $Arguments "action" "undo")
            $steps = [int](Get-Arg $Arguments "steps" 1)
            if ($steps -lt 1 -or $steps -gt 100) { throw "steps must be between 1 and 100." }
            $doc = Resolve-Document $Arguments
            $ok = if ($action -ieq "undo") {
                [bool]$doc.Undo($steps)
            } elseif ($action -ieq "redo") {
                [bool]$doc.Redo($steps)
            } else {
                throw "Unsupported action '$action'."
            }
            $completed = if ($ok) { $steps } else { 0 }
            $save = Save-InPlace $doc $Arguments
            return [ordered]@{ document = Get-DocumentFullName $doc; action = $action; requested = $steps; completed = $completed; save = $save }
        }

        "word_com" {
            $operation = [string](Get-Arg $Arguments "operation" "get")
            if ($operation -ieq "release_all") {
                foreach ($value in @($script:ObjectHandles.Values)) {
                    try { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch {}
                }
                $script:ObjectHandles.Clear()
                return [ordered]@{ released_all = $true }
            }
            if ($operation -ieq "release") {
                $handle = [string](Get-Arg $Arguments "handle")
                if (-not $script:ObjectHandles.ContainsKey($handle)) { throw "Unknown COM handle '$handle'." }
                try { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:ObjectHandles[$handle]) } catch {}
                $script:ObjectHandles.Remove($handle)
                return [ordered]@{ released = $handle }
            }

            $rootName = [string](Get-Arg $Arguments "root" "document")
            switch ($rootName.ToLowerInvariant()) {
                "application" { $root = Get-WordApplication }
                "document" { $root = Resolve-Document $Arguments }
                "selection" { $root = (Get-WordApplication).Selection }
                "handle" {
                    $handle = [string](Get-Arg $Arguments "handle")
                    if (-not $script:ObjectHandles.ContainsKey($handle)) { throw "Unknown COM handle '$handle'." }
                    $root = $script:ObjectHandles[$handle]
                }
                default { throw "Unsupported COM root '$rootName'." }
            }

            $path = [string](Get-Arg $Arguments "path")
            $target = Resolve-ComPathParent $root $path
            if ($operation -ieq "get") {
                $result = Get-ComProperty $target.parent $target.member
            } elseif ($operation -ieq "set") {
                if (-not (Has-Arg $Arguments "value")) { throw "value is required for operation='set'." }
                Set-ComProperty $target.parent $target.member (Get-Arg $Arguments "value")
                $result = Get-ComProperty $target.parent $target.member
            } elseif ($operation -ieq "call") {
                $callArguments = @()
                $rawArguments = Get-Arg $Arguments "arguments" @()
                foreach ($argument in @($rawArguments)) { $callArguments += $argument }
                $result = Invoke-ComMethod $target.parent $target.member $callArguments
            } else {
                throw "Unsupported COM operation '$operation'."
            }

            $save = $null
            if (($operation -ieq "set" -or $operation -ieq "call") -and [bool](Get-Arg $Arguments "save_after" $true)) {
                try {
                    $doc = Resolve-Document $Arguments
                    $save = Save-InPlace $doc $Arguments
                } catch {}
            }
            return [ordered]@{
                root = $rootName
                operation = $operation
                path = $path
                result = Convert-ComResult $result
                save = $save
            }
        }

        default {
            throw "Unknown World tool '$Name'."
        }
    }
}

function New-ObjectSchema {
    param($Properties, [string[]]$Required = @())
    $schema = [ordered]@{ type = "object"; additionalProperties = $false; properties = $Properties }
    if ($Required.Count -gt 0) { $schema.required = $Required }
    return $schema
}

$documentProperty = [ordered]@{ type = "string"; description = "Exact open document title or full path. Defaults to ActiveDocument." }
$scopeProperty = [ordered]@{ type = "string"; enum = @("selection", "current_paragraph", "range", "document"); default = "selection" }
$rangeProperties = [ordered]@{
    document = $documentProperty
    scope = $scopeProperty
    start = [ordered]@{ type = "integer"; minimum = 0; description = "Zero-based Word character start when scope=range." }
    end = [ordered]@{ type = "integer"; minimum = 0; description = "Exclusive Word character end when scope=range." }
}
$saveProperty = [ordered]@{ type = "boolean"; default = $true; description = "Save back to the same document after mutation. Never creates a copy." }
$trackProperty = [ordered]@{ type = "boolean"; description = "Temporary track-changes override for this operation; omitted preserves the current setting." }

$script:Tools = @(
    [ordered]@{
        name = "word_status"
        description = "Connect to desktop Microsoft Word and list open documents, active document, selection, save state, and revision counts. Always call this before the first edit."
        inputSchema = New-ObjectSchema ([ordered]@{})
    },
    [ordered]@{
        name = "word_read"
        description = "Read text and absolute Word character offsets from the active or named open document."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            scope = $scopeProperty
            start = $rangeProperties.start
            end = $rangeProperties.end
            max_chars = [ordered]@{ type = "integer"; minimum = 1; maximum = 200000; default = 20000 }
        })
    },
    [ordered]@{
        name = "word_edit"
        description = "Directly replace or insert text in the live Word document. Saves in place by default and never creates a copy."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            scope = $scopeProperty
            start = $rangeProperties.start
            end = $rangeProperties.end
            mode = [ordered]@{ type = "string"; enum = @("replace", "insert_before", "insert_after"); default = "replace" }
            text = [ordered]@{ type = "string" }
            expected_text = [ordered]@{ type = "string"; description = "Optional exact stale-write guard." }
            track_changes = $trackProperty
            save_after = $saveProperty
        }) @("text")
    },
    [ordered]@{
        name = "word_find_replace"
        description = "Run Word-native find and replace in a document, selection, or range."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            scope = [ordered]@{ type = "string"; enum = @("document", "selection", "current_paragraph", "range"); default = "document" }
            start = $rangeProperties.start
            end = $rangeProperties.end
            find_text = [ordered]@{ type = "string"; minLength = 1 }
            replace_text = [ordered]@{ type = "string"; default = "" }
            replace = [ordered]@{ type = "string"; enum = @("one", "all"); default = "all" }
            match_case = [ordered]@{ type = "boolean"; default = $false }
            whole_word = [ordered]@{ type = "boolean"; default = $false }
            wildcards = [ordered]@{ type = "boolean"; default = $false }
            track_changes = $trackProperty
            save_after = $saveProperty
        }) @("find_text")
    },
    [ordered]@{
        name = "word_format"
        description = "Apply a Word style, font properties, or paragraph formatting to a selection or range."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            scope = $scopeProperty
            start = $rangeProperties.start
            end = $rangeProperties.end
            style = [ordered]@{ type = "string" }
            font = [ordered]@{
                type = "object"
                additionalProperties = $false
                properties = [ordered]@{
                    name = [ordered]@{ type = "string" }
                    size = [ordered]@{ type = "number"; minimum = 1 }
                    bold = [ordered]@{ type = "boolean" }
                    italic = [ordered]@{ type = "boolean" }
                    underline = [ordered]@{ type = "boolean" }
                    color = [ordered]@{ type = "string"; pattern = "^#?[0-9A-Fa-f]{6}$" }
                }
            }
            paragraph = [ordered]@{
                type = "object"
                additionalProperties = $false
                properties = [ordered]@{
                    alignment = [ordered]@{ type = "string"; enum = @("left", "center", "right", "justify") }
                    left_indent = [ordered]@{ type = "number" }
                    right_indent = [ordered]@{ type = "number" }
                    first_line_indent = [ordered]@{ type = "number" }
                    space_before = [ordered]@{ type = "number" }
                    space_after = [ordered]@{ type = "number" }
                    line_spacing = [ordered]@{ type = "number" }
                }
            }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_comment"
        description = "List, add, delete, resolve, or reopen Word comments."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            action = [ordered]@{ type = "string"; enum = @("list", "add", "delete", "resolve", "reopen"); default = "list" }
            scope = $scopeProperty
            start = $rangeProperties.start
            end = $rangeProperties.end
            text = [ordered]@{ type = "string" }
            index = [ordered]@{ type = "integer"; minimum = 1 }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_track_changes"
        description = "Get or set Track Changes on the active or named Word document."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            action = [ordered]@{ type = "string"; enum = @("get", "set"); default = "get" }
            enabled = [ordered]@{ type = "boolean" }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_review_changes"
        description = "List, accept, or reject tracked revisions in the document or current selection."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            action = [ordered]@{ type = "string"; enum = @("list", "accept_all", "reject_all", "accept_next", "reject_next"); default = "list" }
            scope = [ordered]@{ type = "string"; enum = @("document", "selection"); default = "document" }
            limit = [ordered]@{ type = "integer"; minimum = 1; maximum = 1000; default = 200 }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_table"
        description = "List, insert, read/write cells, add rows/columns, or delete Word tables. New tables match nearby document formatting by default."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            action = [ordered]@{ type = "string"; enum = @("list", "insert", "get_cell", "set_cell", "add_row", "add_column", "delete"); default = "list" }
            scope = $scopeProperty
            start = $rangeProperties.start
            end = $rangeProperties.end
            table_index = [ordered]@{ type = "integer"; minimum = 1 }
            row = [ordered]@{ type = "integer"; minimum = 1 }
            column = [ordered]@{ type = "integer"; minimum = 1 }
            rows = [ordered]@{ type = "integer"; minimum = 1; default = 2 }
            columns = [ordered]@{ type = "integer"; minimum = 1; default = 2 }
            text = [ordered]@{ type = "string"; description = "Optional tab-separated cells and newline-separated rows to populate during insertion." }
            auto_fit = [ordered]@{ type = "boolean"; default = $true }
            match_context = [ordered]@{ type = "boolean"; default = $true; description = "When inserting, copy the nearest table's visual formatting unless the user requested a different appearance." }
            template_table_index = [ordered]@{ type = "integer"; minimum = 1; description = "Optional explicit table whose visual formatting should be copied during insertion." }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_image"
        description = "Insert an inline image from an absolute local path into the live Word document."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            scope = $scopeProperty
            start = $rangeProperties.start
            end = $rangeProperties.end
            file_path = [ordered]@{ type = "string" }
            link_to_file = [ordered]@{ type = "boolean"; default = $false }
            save_with_document = [ordered]@{ type = "boolean"; default = $true }
            width = [ordered]@{ type = "number"; minimum = 1; description = "Width in points." }
            height = [ordered]@{ type = "number"; minimum = 1; description = "Height in points." }
            save_after = $saveProperty
        }) @("file_path")
    },
    [ordered]@{
        name = "word_document"
        description = "List, activate, open, create, save, export PDF, print, or close Word documents. Export, print, close, and new should only be used on explicit user request."
        inputSchema = New-ObjectSchema ([ordered]@{
            action = [ordered]@{ type = "string"; enum = @("list", "activate", "open", "new", "save", "save_all", "export_pdf", "print", "close"); default = "list" }
            document = $documentProperty
            file_path = [ordered]@{ type = "string"; description = "Absolute path for open or export_pdf." }
            read_only = [ordered]@{ type = "boolean"; default = $false }
            save_changes = [ordered]@{ type = "boolean"; default = $true }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_undo_redo"
        description = "Undo or redo live Word actions, then save the same document in place."
        inputSchema = New-ObjectSchema ([ordered]@{
            document = $documentProperty
            action = [ordered]@{ type = "string"; enum = @("undo", "redo"); default = "undo" }
            steps = [ordered]@{ type = "integer"; minimum = 1; maximum = 100; default = 1 }
            save_after = $saveProperty
        })
    },
    [ordered]@{
        name = "word_com"
        description = "Generic bridge to nearly the complete Word COM object model. Get/set properties, call methods, retain COM handles, and release handles when focused tools are insufficient."
        inputSchema = New-ObjectSchema ([ordered]@{
            root = [ordered]@{ type = "string"; enum = @("application", "document", "selection", "handle"); default = "document" }
            document = $documentProperty
            handle = [ordered]@{ type = "string" }
            operation = [ordered]@{ type = "string"; enum = @("get", "set", "call", "release", "release_all"); default = "get" }
            path = [ordered]@{ type = "string"; description = "Dot-separated property path; final segment is the method for call." }
            value = [ordered]@{ description = "Property value for set." }
            arguments = [ordered]@{ type = "array"; items = [ordered]@{}; default = @(); description = "Positional COM method arguments." }
            save_after = $saveProperty
        })
    }
)

function Write-JsonMessage {
    param($Object)
    $json = $Object | ConvertTo-Json -Depth 40 -Compress
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

function Write-ToolResult {
    param($Id, $Value, [bool]$IsError = $false)
    $text = if ($Value -is [string]) { $Value } else { $Value | ConvertTo-Json -Depth 30 -Compress }
    $result = [ordered]@{
        content = @([ordered]@{ type = "text"; text = $text })
        isError = $IsError
    }
    if (-not $IsError -and $Value -isnot [string]) {
        $result.structuredContent = $Value
    }
    Write-JsonMessage ([ordered]@{ jsonrpc = "2.0"; id = $Id; result = $result })
}

while ($null -ne ($line = [Console]::In.ReadLine())) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $message = $null
    try {
        $message = $line | ConvertFrom-Json
        $method = [string]$message.method
        $id = $message.id

        switch ($method) {
            "initialize" {
                $requestedVersion = [string](Get-Arg $message.params "protocolVersion" "2025-06-18")
                Write-JsonMessage ([ordered]@{
                    jsonrpc = "2.0"
                    id = $id
                    result = [ordered]@{
                        protocolVersion = $requestedVersion
                        capabilities = [ordered]@{ tools = [ordered]@{ listChanged = $false } }
                        serverInfo = [ordered]@{ name = "world"; version = "0.1.0" }
                        instructions = "Controls the locally running Microsoft Word application and edits open documents in place."
                    }
                })
            }
            "notifications/initialized" {}
            "notifications/cancelled" {}
            "ping" {
                Write-JsonMessage ([ordered]@{ jsonrpc = "2.0"; id = $id; result = [ordered]@{} })
            }
            "tools/list" {
                Write-JsonMessage ([ordered]@{ jsonrpc = "2.0"; id = $id; result = [ordered]@{ tools = $script:Tools } })
            }
            "tools/call" {
                try {
                    $toolName = [string]$message.params.name
                    $toolArguments = $message.params.arguments
                    $value = Invoke-WordTool $toolName $toolArguments
                    Write-ToolResult $id $value $false
                } catch {
                    Write-ToolResult $id ([ordered]@{
                        error = $_.Exception.Message
                        type = $_.Exception.GetType().FullName
                    }) $true
                }
            }
            default {
                if ($null -ne $id) {
                    Write-JsonMessage ([ordered]@{
                        jsonrpc = "2.0"
                        id = $id
                        error = [ordered]@{ code = -32601; message = "Method not found: $method" }
                    })
                }
            }
        }
    } catch {
        $id = $null
        if ($null -ne $message) { $id = $message.id }
        Write-JsonMessage ([ordered]@{
            jsonrpc = "2.0"
            id = $id
            error = [ordered]@{ code = -32603; message = $_.Exception.Message }
        })
    }
}
