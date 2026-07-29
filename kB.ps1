<#
.SYNOPSIS
    Extracts Agent Copilot knowledge article suggestions per conversation for a
    specific division in Genesys Cloud (Australia region), to audit KB article
    effectiveness.

.DESCRIPTION
    Flow:
      1. OAuth client credentials token (Basic auth header built with a pure
         PowerShell Base64 encoder - Constrained Language Mode safe).
      2. POST /api/v2/analytics/conversations/details/query filtered by
         divisionId, paged.
      3. GET /api/v2/conversations/{id}/suggestions for each conversation
         (Agent Copilot knowledge suggestions).
      4. Exports two CSVs:
         - Detail: one row per suggestion per conversation.
         - Summary: per KB article - presented count, accepted count,
           acceptance rate, average confidence.

    CLM-safe: no .NET static methods, no ::new(), no [pscustomobject] casts.
    Windows PowerShell 5.1 / Constrained Language Mode compatible.

.NOTES
    Required OAuth client (Client Credentials grant) role permissions:
      - analytics > conversationDetail > view
      - conversation > suggestion > view   (Agent Copilot suggestions)
      - knowledge > document > view        (optional - title lookup fallback)
    Division-aware permissions must include the target division.

    Copilot suggestion data has limited retention - run this regularly
    (e.g. daily/weekly scheduled task) rather than for long historic ranges.
#>

# =============================================================================
# CONFIGURATION - EDIT THESE
# =============================================================================
$ClientId     = "YOUR_CLIENT_ID_HERE"
$ClientSecret = "YOUR_CLIENT_SECRET_HERE"
$DivisionId   = "YOUR_DIVISION_ID_HERE"

# Australia (Sydney) region endpoints
$LoginBase = "https://login.mypurecloud.com.au"
$ApiBase   = "https://api.mypurecloud.com.au"

# Reporting window (days back from now, UTC)
$DaysBack = 7

# Media types to include in the analytics query (voice + digital)
$MediaTypes = @("voice", "message", "email", "chat")

# Debug / isolation switches
$IncludeMediaFilter = $true    # set $false to test with ONLY the division filter
$ShowRequestBody    = $true    # echoes the JSON body sent to the analytics query (set $false once working)

# Output
$OutputFolder = "C:\Temp\GenesysKB"
$RunStamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$DetailCsv    = $OutputFolder + "\KnowledgeSuggestions_Detail_"  + $RunStamp + ".csv"
$SummaryCsv   = $OutputFolder + "\KnowledgeSuggestions_Summary_" + $RunStamp + ".csv"

# =============================================================================
# CLM-SAFE BASE64 ENCODER (pure PowerShell, bit operators only)
# =============================================================================
function ConvertTo-Base64Clm {
    param([string]$Text)

    $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    $bytes = @()
    foreach ($ch in $Text.ToCharArray()) {
        $bytes += [int][char]$ch    # client id/secret are ASCII-safe
    }

    $result = ""
    $i = 0
    while ($i -lt $bytes.Count) {
        $b0 = $bytes[$i]
        $b1 = -1
        $b2 = -1
        if (($i + 1) -lt $bytes.Count) { $b1 = $bytes[$i + 1] }
        if (($i + 2) -lt $bytes.Count) { $b2 = $bytes[$i + 2] }

        $result += $alphabet[(($b0 -shr 2) -band 63)]

        if ($b1 -ge 0) {
            $result += $alphabet[((($b0 -shl 4) -bor ($b1 -shr 4)) -band 63)]
            if ($b2 -ge 0) {
                $result += $alphabet[((($b1 -shl 2) -bor ($b2 -shr 6)) -band 63)]
                $result += $alphabet[($b2 -band 63)]
            }
            else {
                $result += $alphabet[(($b1 -shl 2) -band 63)]
                $result += "="
            }
        }
        else {
            $result += $alphabet[(($b0 -shl 4) -band 63)]
            $result += "=="
        }
        $i += 3
    }
    return $result
}

# =============================================================================
# CLM-SAFE HTTP HELPER (429 retry + status extraction)
# =============================================================================
function Invoke-GcApi {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Body = $null,
        [int]$MaxRetries = 3
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $Body -ContentType "application/json" -ErrorAction Stop
            }
            else {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ErrorAction Stop
            }
        }
        catch {
            $status = 0
            try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
            if ($status -eq 0) {
                $msg = [string]$_.Exception.Message
                if ($msg -like "*429*") { $status = 429 }
                elseif ($msg -like "*404*") { $status = 404 }
                elseif ($msg -like "*401*") { $status = 401 }
                elseif ($msg -like "*403*") { $status = 403 }
            }

            if ($status -eq 429 -and $attempt -le $MaxRetries) {
                Write-Warning ("Rate limited (429) on " + $Uri + " - waiting 60s (attempt " + $attempt + " of " + $MaxRetries + ")")
                Start-Sleep -Seconds 60
                continue
            }
            if ($status -eq 404) {
                # No suggestions / resource not found - treat as empty, not fatal
                return $null
            }

            # Surface the Genesys error body - it states the exact rejection reason
            $errBody = ""
            try {
                if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                    $errBody = [string]$_.ErrorDetails.Message
                }
            } catch { $errBody = "" }

            Write-Warning ("API call failed [" + $status + "] " + $Method + " " + $Uri + " :: " + $_.Exception.Message)
            if ($errBody -ne "") {
                Write-Warning ("Genesys error body: " + $errBody)
            }
            return $null
        }
    }
}

# =============================================================================
# 1. AUTHENTICATE
# =============================================================================
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

Write-Host "Authenticating to Genesys Cloud (AU region)..." -ForegroundColor Cyan

$basicToken = ConvertTo-Base64Clm ($ClientId + ":" + $ClientSecret)
$authHeaders = @{ "Authorization" = "Basic " + $basicToken }
$authBody = "grant_type=client_credentials"

$tokenResponse = $null
try {
    $tokenResponse = Invoke-RestMethod -Method Post -Uri ($LoginBase + "/oauth/token") -Headers $authHeaders -Body $authBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
}
catch {
    $status = 0
    try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
    Write-Error ("Authentication request failed (HTTP " + $status + "): " + $_.Exception.Message)
    exit 1
}

if (-not $tokenResponse -or -not $tokenResponse.access_token) {
    Write-Error "Authentication failed - no access_token in response. Check client id/secret and region."
    exit 1
}

$headers = @{
    "Authorization" = "Bearer " + $tokenResponse.access_token
    "Content-Type"  = "application/json"
}
Write-Host "Authenticated OK." -ForegroundColor Green

# =============================================================================
# 1b. VALIDATE DIVISION ID (most common cause of 400 on the analytics query)
# =============================================================================
$DivisionId = $DivisionId.Trim()
$guidPattern = "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"

if ($DivisionId -notmatch $guidPattern) {
    Write-Warning ("DivisionId '" + $DivisionId + "' is not a valid GUID. Here are the divisions visible to this OAuth client:")
    $divList = Invoke-GcApi -Method "Get" -Uri ($ApiBase + "/api/v2/authorization/divisions?pageSize=100") -Headers $headers
    if ($divList -and $divList.entities) {
        foreach ($d in $divList.entities) {
            Write-Host ("  " + [string]$d.id + "   " + [string]$d.name)
        }
    }
    else {
        Write-Warning "Could not list divisions (check 'authorization > division > view' permission). Get the ID from Admin > Directory > Divisions."
    }
    exit 1
}

# =============================================================================
# 2. GET CONVERSATIONS FOR THE DIVISION (analytics details query, paged)
# =============================================================================
$endUtc   = Get-Date
$endUtc   = $endUtc.ToUniversalTime()
$startUtc = $endUtc.AddDays(-1 * $DaysBack)

$intervalString = $startUtc.ToString("yyyy-MM-ddTHH:mm:ss") + ".000Z/" + $endUtc.ToString("yyyy-MM-ddTHH:mm:ss") + ".000Z"
Write-Host ("Query interval (UTC): " + $intervalString) -ForegroundColor Cyan

$conversationList = @()
$pageNumber = 1
$pageSize   = 100

while ($true) {
    $queryBody = @{
        interval = $intervalString
        order    = "asc"
        orderBy  = "conversationStart"
        paging   = @{
            pageSize   = $pageSize
            pageNumber = $pageNumber
        }
        segmentFilters = @(
            @{
                type = "or"
                predicates = @(
                    @{
                        type      = "dimension"
                        dimension = "divisionId"
                        operator  = "matches"
                        value     = $DivisionId
                    }
                )
            }
        )
        segmentFilters2 = $null
    }
    # Remove helper null key before serialising
    $queryBody.Remove("segmentFilters2")

    # Add media type filter as a segment filter (any of the listed media types)
    if ($IncludeMediaFilter) {
        $mediaPredicates = @()
        foreach ($mt in $MediaTypes) {
            $mediaPredicates += @{
                type      = "dimension"
                dimension = "mediaType"
                operator  = "matches"
                value     = $mt
            }
        }
        $queryBody.segmentFilters += @{
            type       = "or"
            predicates = $mediaPredicates
        }
    }

    $jsonBody = $queryBody | ConvertTo-Json -Depth 10
    if ($ShowRequestBody -and $pageNumber -eq 1) {
        Write-Host "---- Analytics request body ----" -ForegroundColor DarkGray
        Write-Host $jsonBody -ForegroundColor DarkGray
        Write-Host "--------------------------------" -ForegroundColor DarkGray
    }

    $page = Invoke-GcApi -Method "Post" -Uri ($ApiBase + "/api/v2/analytics/conversations/details/query") -Headers $headers -Body $jsonBody
    if (-not $page) {
        Write-Warning ("Analytics query returned nothing on page " + $pageNumber + " - stopping pagination.")
        break
    }

    $count = 0
    if ($page.conversations) {
        $count = @($page.conversations).Count
        foreach ($conv in $page.conversations) {
            $queueId = ""
            # Grab first queueId found across participants/sessions/segments (best effort)
            foreach ($p in $conv.participants) {
                foreach ($s in $p.sessions) {
                    foreach ($seg in $s.segments) {
                        if ($seg.queueId -and $queueId -eq "") { $queueId = [string]$seg.queueId }
                    }
                }
            }
            $conversationList += New-Object PSObject -Property @{
                ConversationId    = [string]$conv.conversationId
                ConversationStart = [string]$conv.conversationStart
                ConversationEnd   = [string]$conv.conversationEnd
                QueueId           = $queueId
            }
        }
    }

    Write-Host ("Page " + $pageNumber + ": " + $count + " conversations (running total " + $conversationList.Count + ")")

    if ($count -lt $pageSize) { break }
    $pageNumber++
    if ($pageNumber -gt 100) {
        Write-Warning "Pagination safety cap hit (100 pages / 10,000 conversations). Narrow the interval."
        break
    }
}

Write-Host ("Total conversations in division for interval: " + $conversationList.Count) -ForegroundColor Green

if ($conversationList.Count -eq 0) {
    Write-Host "No conversations found - nothing to do." -ForegroundColor Yellow
    exit 0
}

# =============================================================================
# 3. GET COPILOT KNOWLEDGE SUGGESTIONS PER CONVERSATION
# =============================================================================
$detailRows   = @()
$titleCache   = @{}    # documentId -> title (avoid repeat knowledge API lookups)
$processed    = 0

foreach ($conv in $conversationList) {
    $processed++
    if (($processed % 25) -eq 0) {
        Write-Host ("Processed " + $processed + " of " + $conversationList.Count + " conversations...")
    }

    $sugUri = $ApiBase + "/api/v2/conversations/" + $conv.ConversationId + "/suggestions?pageSize=100"
    $sugResponse = Invoke-GcApi -Method "Get" -Uri $sugUri -Headers $headers
    if (-not $sugResponse) { continue }

    $entities = @()
    if ($sugResponse.entities) { $entities = @($sugResponse.entities) }
    if ($entities.Count -eq 0) { continue }

    foreach ($sug in $entities) {
        $sugType = ""
        if ($sug.type) { $sugType = [string]$sug.type }

        # Only knowledge-related suggestions are relevant to the KB audit
        if (($sugType -notlike "*Knowledge*") -and ($sugType -notlike "*knowledge*")) { continue }

        $state       = ""
        $confidence  = ""
        $kbId        = ""
        $docId       = ""
        $docVersion  = ""
        $title       = ""
        $utterance   = ""

        if ($sug.state)      { $state      = [string]$sug.state }
        if ($sug.confidence) { $confidence = [string]$sug.confidence }

        # Knowledge article payload - schema-defensive extraction
        $k = $null
        if ($sug.knowledgeArticleSuggestion) { $k = $sug.knowledgeArticleSuggestion }
        elseif ($sug.knowledgeSearchSuggestion) { $k = $sug.knowledgeSearchSuggestion }
        elseif ($sug.knowledge) { $k = $sug.knowledge }

        if ($k) {
            if ($k.knowledgeBase -and $k.knowledgeBase.id) { $kbId = [string]$k.knowledgeBase.id }
            elseif ($k.knowledgeBaseId) { $kbId = [string]$k.knowledgeBaseId }

            if ($k.document -and $k.document.id) { $docId = [string]$k.document.id }
            elseif ($k.documentId) { $docId = [string]$k.documentId }

            if ($k.document -and $k.document.version) { $docVersion = [string]$k.document.version }

            if ($k.document -and $k.document.title) { $title = [string]$k.document.title }
            elseif ($k.title) { $title = [string]$k.title }

            if ($k.query) { $utterance = [string]$k.query }
        }

        # Triggering utterance / message that caused the suggestion (best effort)
        if ($utterance -eq "" -and $sug.triggeringMessage -and $sug.triggeringMessage.text) {
            $utterance = [string]$sug.triggeringMessage.text
        }
        if ($utterance -eq "" -and $sug.utterance) {
            $utterance = [string]$sug.utterance
        }

        # Title fallback via Knowledge API (cached per document)
        if ($title -eq "" -and $kbId -ne "" -and $docId -ne "") {
            $cacheKey = $kbId + "|" + $docId
            if ($titleCache.ContainsKey($cacheKey)) {
                $title = $titleCache[$cacheKey]
            }
            else {
                $docUri = $ApiBase + "/api/v2/knowledge/knowledgebases/" + $kbId + "/documents/" + $docId
                $doc = Invoke-GcApi -Method "Get" -Uri $docUri -Headers $headers
                if ($doc -and $doc.title) { $title = [string]$doc.title }
                $titleCache[$cacheKey] = $title
            }
        }

        $detailRows += New-Object PSObject -Property @{
            ConversationId    = $conv.ConversationId
            ConversationStart = $conv.ConversationStart
            QueueId           = $conv.QueueId
            SuggestionId      = [string]$sug.id
            SuggestionType    = $sugType
            State             = $state
            Confidence        = $confidence
            KnowledgeBaseId   = $kbId
            DocumentId        = $docId
            DocumentVersion   = $docVersion
            ArticleTitle      = $title
            TriggeringText    = $utterance
        }
    }
}

Write-Host ("Knowledge suggestion rows collected: " + $detailRows.Count) -ForegroundColor Green

if ($detailRows.Count -eq 0) {
    Write-Host "No knowledge suggestions found for these conversations. Check that Agent Copilot is enabled on the relevant queues, the OAuth client has 'conversation > suggestion > view', and the interval is within suggestion retention." -ForegroundColor Yellow
    exit 0
}

# =============================================================================
# 4. EXPORT DETAIL CSV (explicit column order for CLM property hashtables)
# =============================================================================
$detailRows |
    Select-Object ConversationId, ConversationStart, QueueId, SuggestionId, SuggestionType, State, Confidence, KnowledgeBaseId, DocumentId, DocumentVersion, ArticleTitle, TriggeringText |
    Export-Csv -Path $DetailCsv -NoTypeInformation -Encoding UTF8

Write-Host ("Detail CSV written: " + $DetailCsv) -ForegroundColor Green

# =============================================================================
# 5. BUILD PER-ARTICLE KB IMPROVEMENT SUMMARY
# =============================================================================
$summaryRows = @()
$grouped = $detailRows | Group-Object DocumentId

foreach ($g in $grouped) {
    if ($g.Name -eq "") { continue }

    $presented  = $g.Count
    $accepted   = 0
    $rejected   = 0
    $confSum    = 0.0
    $confCount  = 0
    $firstTitle = ""
    $firstKb    = ""
    $convIds    = @{}

    foreach ($row in $g.Group) {
        if ($firstTitle -eq "" -and $row.ArticleTitle -ne "") { $firstTitle = $row.ArticleTitle }
        if ($firstKb -eq "" -and $row.KnowledgeBaseId -ne "") { $firstKb = $row.KnowledgeBaseId }
        $convIds[$row.ConversationId] = $true

        if ($row.State -like "*Accept*") { $accepted++ }
        if ($row.State -like "*Reject*" -or $row.State -like "*Dismiss*") { $rejected++ }

        if ($row.Confidence -ne "") {
            $c = 0.0
            $parsed = $false
            try { $c = [double]$row.Confidence; $parsed = $true } catch { $parsed = $false }
            if ($parsed) {
                $confSum = $confSum + $c
                $confCount++
            }
        }
    }

    $avgConf = ""
    if ($confCount -gt 0) {
        $avgConf = "{0:N3}" -f ($confSum / $confCount)
    }

    $acceptRate = "0.0%"
    if ($presented -gt 0) {
        $acceptRate = "{0:N1}%" -f (($accepted / $presented) * 100)
    }

    $summaryRows += New-Object PSObject -Property @{
        DocumentId          = $g.Name
        ArticleTitle        = $firstTitle
        KnowledgeBaseId     = $firstKb
        TimesPresented      = $presented
        UniqueConversations = $convIds.Keys.Count
        TimesAccepted       = $accepted
        TimesRejected       = $rejected
        AcceptanceRate      = $acceptRate
        AvgConfidence       = $avgConf
    }
}

$summaryRows |
    Sort-Object TimesPresented -Descending |
    Select-Object ArticleTitle, DocumentId, KnowledgeBaseId, TimesPresented, UniqueConversations, TimesAccepted, TimesRejected, AcceptanceRate, AvgConfidence |
    Export-Csv -Path $SummaryCsv -NoTypeInformation -Encoding UTF8

Write-Host ("Summary CSV written: " + $SummaryCsv) -ForegroundColor Green
Write-Host ""
Write-Host "Done. Articles with high TimesPresented but low Accepta
