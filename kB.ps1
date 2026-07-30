<#
.SYNOPSIS
    Extracts, per conversation in a specific Genesys Cloud division (Australia
    region), the agent, queue, conversation transcript and the knowledge base
    articles that were presented (Agent Copilot suggestions) - to audit and
    improve KB article effectiveness.

.DESCRIPTION
    Flow:
      1. OAuth client credentials token (Basic auth header built with a pure
         PowerShell Base64 encoder - Constrained Language Mode safe).
      2. POST /api/v2/analytics/conversations/details/query filtered by
         divisionId, paged. Captures agent userId, queue id, customer/agent
         session ids per conversation.
      3. GET /api/v2/conversations/{id}/suggestions for each conversation
         (Agent Copilot knowledge suggestions - which article, when, why).
      4. GET /api/v2/speechandtextanalytics/.../transcripturl per conversation
         to pull the transcript text (customer + agent phrases).
      5. Resolves agent userIds -> names (GET /api/v2/users/{id}) and
         queue ids -> names (GET /api/v2/routing/queues/{id}), cached.
      6. Exports four CSVs:
         - Overview : ONE ROW PER CONVERSATION - agent, queue, transcript
                      text, articles presented and why. (The main deliverable.)
         - Detail   : one row per article suggestion (drill-down).
         - Transcript: one row per transcript phrase (timeline join).
         - Summary  : per KB article - presented/accepted counts, acceptance
                      rate, average confidence (rewrite candidates).

    CLM-safe: no .NET static method calls, no ::new(), no [pscustomobject]
    casts, no [uri]:: encoding. Windows PowerShell 5.1 / Constrained Language
    Mode compatible.

.NOTES
    Required OAuth client (Client Credentials grant) role permissions:
      - analytics > conversationDetail > view
      - conversation > suggestion > view      (Agent Copilot suggestions)
      - speechAndTextAnalytics > data > view  (transcripts)
      - routing > queue > view                (queue name lookup)
      - directory > user > view               (agent name lookup)
      - knowledge > document > view           (optional - title fallback)
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
$IncludeMediaFilter       = $true   # set $false to test with ONLY the division filter
$ShowRequestBody          = $false  # echoes the JSON body sent to the analytics query
$OnlyCopilotConversations = $true   # only pull conversations where an agent assistant (Copilot) was present
$DumpFirstSuggestion      = $true   # dumps the first raw non-empty suggestions response (schema check)

# Transcript scope:
#   $false = only fetch transcripts for conversations where a KB article was
#            presented (fewer API calls - recommended)
#   $true  = fetch transcripts for EVERY conversation in the division/window
$TranscriptAllConversations = $false

# Excel cell hard limit is 32,767 chars - transcripts longer than this are
# truncated in the overview CSV (full phrases remain in the transcript CSV)
$TranscriptMaxChars = 30000

# Output
$OutputFolder  = "C:\Temp\GenesysKB"
$RunStamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$OverviewCsv   = $OutputFolder + "\KB_PerConversation_"              + $RunStamp + ".csv"
$DetailCsv     = $OutputFolder + "\KnowledgeSuggestions_Detail_"     + $RunStamp + ".csv"
$SummaryCsv    = $OutputFolder + "\KnowledgeSuggestions_Summary_"    + $RunStamp + ".csv"
$TranscriptCsv = $OutputFolder + "\KnowledgeSuggestions_Transcript_" + $RunStamp + ".csv"

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
# CLM-SAFE RECURSIVE PROPERTY FINDERS
# Walk a deserialised JSON object tree and return the first match by name,
# regardless of nesting depth - so schema differences don't break extraction.
# =============================================================================
function Test-IsScalar {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [string] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [bool]) { return $true }
    return $false
}

function Find-ScalarProp {
    param($Obj, [string[]]$Names, [int]$MaxDepth = 7)
    if ($null -eq $Obj -or $MaxDepth -lt 0) { return "" }
    if (Test-IsScalar $Obj) { return "" }
    if ($Obj -is [array]) {
        foreach ($item in $Obj) {
            $v = Find-ScalarProp $item $Names ($MaxDepth - 1)
            if ($v -ne "") { return $v }
        }
        return ""
    }
    $props = $null
    try { $props = $Obj.PSObject.Properties } catch { return "" }
    if (-not $props) { return "" }
    # Pass 1: direct name match with a scalar value
    foreach ($p in $props) {
        foreach ($n in $Names) {
            if ($p.Name -eq $n) {
                if (Test-IsScalar $p.Value) {
                    if (([string]$p.Value) -ne "") { return [string]$p.Value }
                }
            }
        }
    }
    # Pass 2: recurse into children
    foreach ($p in $props) {
        if (-not (Test-IsScalar $p.Value)) {
            $v = Find-ScalarProp $p.Value $Names ($MaxDepth - 1)
            if ($v -ne "") { return $v }
        }
    }
    return ""
}

function Find-ObjectProp {
    param($Obj, [string]$Name, [int]$MaxDepth = 7)
    if ($null -eq $Obj -or $MaxDepth -lt 0) { return $null }
    if (Test-IsScalar $Obj) { return $null }
    if ($Obj -is [array]) {
        foreach ($item in $Obj) {
            $v = Find-ObjectProp $item $Name ($MaxDepth - 1)
            if ($null -ne $v) { return $v }
        }
        return $null
    }
    $props = $null
    try { $props = $Obj.PSObject.Properties } catch { return $null }
    if (-not $props) { return $null }
    foreach ($p in $props) {
        if ($p.Name -eq $Name -and $null -ne $p.Value -and -not (Test-IsScalar $p.Value)) { return $p.Value }
    }
    foreach ($p in $props) {
        if (-not (Test-IsScalar $p.Value)) {
            $v = Find-ObjectProp $p.Value $Name ($MaxDepth - 1)
            if ($null -ne $v) { return $v }
        }
    }
    return $null
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
# 1c. NAME RESOLUTION HELPERS (agent userId -> name, queueId -> name), cached
# =============================================================================
$userNameCache  = @{}
$queueNameCache = @{}

function Resolve-AgentName {
    param([string]$UserId, [hashtable]$AuthHeaders)
    if ($UserId -eq "") { return "" }
    if ($script:userNameCache.ContainsKey($UserId)) { return $script:userNameCache[$UserId] }
    $name = ""
    $u = Invoke-GcApi -Method "Get" -Uri ($script:ApiBase + "/api/v2/users/" + $UserId + "?state=any") -Headers $AuthHeaders
    if ($u -and $u.name) { $name = [string]$u.name }
    if ($name -eq "") { $name = $UserId }   # fall back to the raw id
    $script:userNameCache[$UserId] = $name
    return $name
}

function Resolve-QueueName {
    param([string]$QueueId, [hashtable]$AuthHeaders)
    if ($QueueId -eq "") { return "" }
    if ($script:queueNameCache.ContainsKey($QueueId)) { return $script:queueNameCache[$QueueId] }
    $name = ""
    $q = Invoke-GcApi -Method "Get" -Uri ($script:ApiBase + "/api/v2/routing/queues/" + $QueueId) -Headers $AuthHeaders
    if ($q -and $q.name) { $name = [string]$q.name }
    if ($name -eq "") { $name = $QueueId }  # fall back to the raw id
    $script:queueNameCache[$QueueId] = $name
    return $name
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
        conversationFilters = @(
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
    }

    # Add media type filter as a segment filter (any of the listed media types)
    $segFilters = @()
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
        $segFilters += @{
            type       = "or"
            predicates = $mediaPredicates
        }
    }

    # Only conversations where Agent Copilot / Agent Assist participated -
    # avoids querying suggestions for thousands of conversations that never had it
    if ($OnlyCopilotConversations) {
        $segFilters += @{
            type = "or"
            predicates = @(
                @{
                    type      = "dimension"
                    dimension = "agentAssistantId"
                    operator  = "exists"
                }
            )
        }
    }

    if ($segFilters.Count -gt 0) {
        $queryBody["segmentFilters"] = $segFilters
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
            $queueId     = ""
            $agentUserId = ""
            $custSessId  = ""
            $agentSessId = ""
            $convMedia   = ""
            foreach ($p in $conv.participants) {
                $purpose = ""
                if ($p.purpose) { $purpose = [string]$p.purpose }

                # First agent participant with a userId = the handling agent
                if ($agentUserId -eq "" -and $purpose -eq "agent" -and $p.userId) {
                    $agentUserId = [string]$p.userId
                }

                foreach ($s in $p.sessions) {
                    if ($convMedia -eq "" -and $s.mediaType) { $convMedia = [string]$s.mediaType }
                    # Session ids double as communicationId for transcript retrieval:
                    # customer side preferred, agent side kept as fallback
                    if ($custSessId -eq "" -and ($purpose -eq "customer" -or $purpose -eq "external") -and $s.sessionId) {
                        $custSessId = [string]$s.sessionId
                    }
                    if ($agentSessId -eq "" -and $purpose -eq "agent" -and $s.sessionId) {
                        $agentSessId = [string]$s.sessionId
                    }
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
                AgentUserId       = $agentUserId
                MediaType         = $convMedia
                CustomerSessionId = $custSessId
                AgentSessionId    = $agentSessId
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
$statNoData   = 0      # 404 / null responses
$statEmpty    = 0      # 200 but no entities
$statHasData  = 0      # conversations that returned suggestion entities
$statNonKb    = 0      # suggestion entities that were not knowledge-related
$dumpedSample = $false

foreach ($conv in $conversationList) {
    $processed++
    if (($processed % 25) -eq 0) {
        Write-Host ("Processed " + $processed + " of " + $conversationList.Count + " conversations...")
    }

    $sugUri = $ApiBase + "/api/v2/conversations/" + $conv.ConversationId + "/suggestions?pageSize=100"
    $sugResponse = Invoke-GcApi -Method "Get" -Uri $sugUri -Headers $headers
    if (-not $sugResponse) { $statNoData++; continue }

    $entities = @()
    if ($sugResponse.entities) { $entities = @($sugResponse.entities) }
    if ($entities.Count -eq 0) { $statEmpty++; continue }

    $statHasData++

    # One-off raw dump so we can verify the actual schema in your org
    if ($DumpFirstSuggestion -and (-not $dumpedSample)) {
        $rawPath = $OutputFolder + "\RawSuggestionSample_" + $RunStamp + ".json"
        ($sugResponse | ConvertTo-Json -Depth 15) | Out-File -FilePath $rawPath -Encoding UTF8
        Write-Host ("Raw suggestions sample saved to: " + $rawPath) -ForegroundColor Magenta
        $dumpedSample = $true
    }

    foreach ($sug in $entities) {
        $sugType = ""
        if ($sug.type) { $sugType = [string]$sug.type }

        # Knowledge-related if the type says so OR any knowledge payload property exists
        $isKnowledge = $false
        if ($sugType -like "*nowledge*") { $isKnowledge = $true }
        if ($sug.knowledgeArticleSuggestion) { $isKnowledge = $true }
        if ($sug.knowledgeSearchSuggestion)  { $isKnowledge = $true }
        if ($sug.knowledge)                  { $isKnowledge = $true }
        if (-not $isKnowledge) { $statNonKb++; continue }

        $state       = ""
        $confidence  = ""
        $kbId        = ""
        $docId       = ""
        $docVersion  = ""
        $title       = ""
        $utterance   = ""
        $sugTime     = ""

        if ($sug.state) { $state = [string]$sug.state }
        else { $state = Find-ScalarProp $sug @("state") }

        # Confidence can sit at top level, inside the trigger, or inside the article payload
        $confidence = Find-ScalarProp $sug @("confidence")

        # WHY was it presented - the trigger that fired the Copilot rule.
        # trigger.type tells us what trigger.value means:
        #   type like "Intent"    -> value is the detected intent NAME
        #   type like "Utterance" -> value is the spoken/typed TEXT itself
        # (The old code shoved value into the intent column regardless - that
        #  is why TriggeringText came out empty for utterance triggers.)
        $intentName = ""
        $trigTime   = ""
        $trigObj = Find-ObjectProp $sug "trigger"
        if ($trigObj) {
            $trigType  = Find-ScalarProp $trigObj @("type")
            $trigValue = Find-ScalarProp $trigObj @("value")
            $trigTime  = Find-ScalarProp $trigObj @("eventTime", "dateCreated", "timestamp")
            if ($trigType -like "*ntent*") {
                $intentName = $trigValue
            }
            elseif ($trigValue -ne "") {
                $utterance = $trigValue
            }
            if ($intentName -eq "") { $intentName = Find-ScalarProp $trigObj @("intent", "intentName", "name") }
        }
        if ($intentName -eq "") { $intentName = Find-ScalarProp $sug @("intent", "intentName") }

        # When was the article surfaced - needed to line up with the transcript
        $sugTime = Find-ScalarProp $sug @("dateCreated", "timestamp", "dateModified", "dateStart")

        # Knowledge identifiers - wherever they are nested
        $kbId = Find-ScalarProp $sug @("knowledgeBaseId")
        if ($kbId -eq "") {
            $kbObj = Find-ObjectProp $sug "knowledgeBase"
            if ($kbObj) { $kbId = Find-ScalarProp $kbObj @("id") }
        }

        $docId = Find-ScalarProp $sug @("documentId")
        $docObj = Find-ObjectProp $sug "document"
        if ($docId -eq "" -and $docObj) { $docId = Find-ScalarProp $docObj @("id") }
        if ($docObj) {
            $docVersion = Find-ScalarProp $docObj @("version", "versionId")
            $title      = Find-ScalarProp $docObj @("title")
        }
        if ($title -eq "") {
            $artObj = Find-ObjectProp $sug "article"
            if ($artObj) {
                $title = Find-ScalarProp $artObj @("title")
                if ($docId -eq "") { $docId = Find-ScalarProp $artObj @("documentId", "id") }
                if ($docVersion -eq "") { $docVersion = Find-ScalarProp $artObj @("versionId", "version") }
            }
        }
        if ($title -eq "") { $title = Find-ScalarProp $sug @("title") }

        # The phrase / search query that caused the suggestion = the "why"
        if ($utterance -eq "") {
            $utterance = Find-ScalarProp $sug @("query", "searchQuery", "utterance", "transcriptionText")
        }
        if ($utterance -eq "" -and $trigObj) {
            $utterance = Find-ScalarProp $trigObj @("text", "message", "phrase", "utterance")
        }
        $utteranceSource = ""
        if ($utterance -ne "") { $utteranceSource = "copilot-api" }

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
            AgentUserId       = $conv.AgentUserId
            MediaType         = $conv.MediaType
            SuggestionId      = [string]$sug.id
            SuggestionType    = $sugType
            SuggestedAtUtc    = $sugTime
            TriggerTimeUtc    = $trigTime
            State             = $state
            Confidence        = $confidence
            TriggerIntent     = $intentName
            KnowledgeBaseId   = $kbId
            DocumentId        = $docId
            DocumentVersion   = $docVersion
            ArticleTitle      = $title
            TriggeringText    = $utterance
            TriggerTextSource = $utteranceSource
        }
    }
}

Write-Host ""
Write-Host ("Suggestion endpoint results across " + $conversationList.Count + " conversations:") -ForegroundColor Cyan
Write-Host ("  With suggestion data : " + $statHasData)
Write-Host ("  200 but empty        : " + $statEmpty)
Write-Host ("  404 / no data / error: " + $statNoData)
Write-Host ("  Non-knowledge entries skipped: " + $statNonKb)
Write-Host ("Knowledge suggestion rows collected: " + $detailRows.Count) -ForegroundColor Green

if ($detailRows.Count -eq 0) {
    Write-Host ""
    Write-Host "Diagnosis guide:" -ForegroundColor Yellow
    Write-Host " - Mostly 404/no data + warnings showing 403: OAuth client is missing 'conversation > suggestion > view' - add it to the role." -ForegroundColor Yellow
    Write-Host " - Mostly 200-but-empty: suggestion data has likely aged out (short retention, similar to Copilot summaries ~10 days) - re-run with DaysBack = 2 or 3." -ForegroundColor Yellow
    Write-Host " - Zero conversations found at all: your Copilot may be recorded under a different dimension - set OnlyCopilotConversations = `$false and retry." -ForegroundColor Yellow
    Write-Host " - Non-knowledge entries skipped > 0 but rows = 0: only canned response/script suggestions fired - check Copilot rules and KB confidence threshold." -ForegroundColor Yellow
    Write-Host "The per-conversation overview will still be written with ArticlePresented = NO for every row." -ForegroundColor Yellow
}

# =============================================================================
# 4. TRANSCRIPTS - what was said in each conversation
#    Fetched BEFORE the overview so the transcript text can sit on the same
#    row as agent / queue / articles presented.
#    Requires: speechAndTextAnalytics > data > view, and voice transcription
#    or digital transcripts enabled on the relevant queues.
# =============================================================================
Write-Host ""
Write-Host "Fetching transcripts..." -ForegroundColor Cyan

# Which conversations get a transcript?
$transcriptConvIds = @{}
if ($TranscriptAllConversations) {
    foreach ($c in $conversationList) { $transcriptConvIds[$c.ConversationId] = $true }
}
else {
    foreach ($d in $detailRows) { $transcriptConvIds[$d.ConversationId] = $true }
}
Write-Host ("Transcript scope: " + $transcriptConvIds.Keys.Count + " conversation(s)")

$transcriptRows = @()     # one row per phrase (for the transcript CSV)
$convTranscripts = @{}    # conversationId -> single joined text block (for the overview CSV)
$convPhrases = @{}        # conversationId -> phrase objects with parsed times (for trigger-text backfill)
$epochBase = [datetime]"1970-01-01T00:00:00Z"

foreach ($c in $conversationList) {
    if (-not $transcriptConvIds.ContainsKey($c.ConversationId)) { continue }

    # Try the customer-side communication first, then the agent side
    $commIds = @()
    if ($c.CustomerSessionId -ne "") { $commIds += $c.CustomerSessionId }
    if ($c.AgentSessionId -ne "")    { $commIds += $c.AgentSessionId }
    if ($commIds.Count -eq 0) {
        Write-Warning ("No session id for conversation " + $c.ConversationId + " - skipping transcript.")
        continue
    }

    $tdoc = $null
    foreach ($commId in $commIds) {
        $turlUri = $ApiBase + "/api/v2/speechandtextanalytics/conversations/" + $c.ConversationId + "/communications/" + $commId + "/transcripturl"
        $turl = Invoke-GcApi -Method "Get" -Uri $turlUri -Headers $headers
        if (-not $turl -or -not $turl.url) { continue }

        # The returned URL is pre-signed - fetch without auth headers
        try {
            $tdoc = Invoke-RestMethod -Method Get -Uri ([string]$turl.url) -ErrorAction Stop
        }
        catch {
            Write-Warning ("Transcript download failed for " + $c.ConversationId + " :: " + $_.Exception.Message)
            $tdoc = $null
        }
        if ($tdoc) { break }
    }
    if (-not $tdoc) { continue }

    # Transcript JSON: transcripts[] -> phrases[] with text / participantPurpose / startTimeMs
    $tsets = @()
    if ($tdoc.transcripts) { $tsets = @($tdoc.transcripts) }
    elseif ($tdoc.phrases) { $tsets = @($tdoc) }

    $joined = ""
    $phraseObjs = @()
    foreach ($tset in $tsets) {
        $phrases = @()
        if ($tset.phrases) { $phrases = @($tset.phrases) }
        foreach ($ph in $phrases) {
            $text = ""
            if ($ph.text) { $text = [string]$ph.text }
            elseif ($ph.decoratedText) { $text = [string]$ph.decoratedText }
            if ($text -eq "") { continue }

            $speaker = ""
            if ($ph.participantPurpose) { $speaker = [string]$ph.participantPurpose }
            if ($speaker -eq "") { $speaker = "unknown" }

            # startTimeMs may be epoch milliseconds (absolute) or an offset
            $phTimeUtc = ""
            $parsedTime = $null
            $rawMs = 0
            $gotMs = $false
            try {
                if ($ph.startTimeMs) { $rawMs = [double]$ph.startTimeMs; $gotMs = $true }
            } catch { $gotMs = $false }
            if ($gotMs) {
                if ($rawMs -gt 1000000000000) {
                    # epoch ms -> UTC datetime via instance methods (CLM-safe).
                    # The [datetime]"...Z" cast converts to LOCAL time, so
                    # normalise back to true UTC to match the API's UTC strings.
                    $dt = $epochBase.AddMilliseconds($rawMs)
                    $dt = $dt.ToUniversalTime()
                    $phTimeUtc = $dt.ToString("yyyy-MM-ddTHH:mm:ss") + "Z"
                    $parsedTime = $dt
                }
                else {
                    # offset ms from start of communication
                    $phTimeUtc = "offset:" + [string][int]$rawMs + "ms"
                }
            }

            $transcriptRows += New-Object PSObject -Property @{
                ConversationId = $c.ConversationId
                PhraseTimeUtc  = $phTimeUtc
                Speaker        = $speaker
                Text           = $text
            }

            $phraseObjs += New-Object PSObject -Property @{
                ParsedTime = $parsedTime
                Speaker    = $speaker
                Text       = $text
            }

            if ($joined -ne "") { $joined = $joined + "`n" }
            $joined = $joined + "[" + $speaker + "] " + $text
        }
    }

    if ($phraseObjs.Count -gt 0) {
        $convPhrases[$c.ConversationId] = $phraseObjs
    }

    if ($joined -ne "") {
        if ($joined.Length -gt $TranscriptMaxChars) {
            $joined = $joined.Substring(0, $TranscriptMaxChars) + " ...[TRUNCATED - full text in transcript CSV]"
        }
        $convTranscripts[$c.ConversationId] = $joined
    }
}

Write-Host ("Transcripts retrieved for " + $convTranscripts.Keys.Count + " conversation(s), " + $transcriptRows.Count + " phrases total.") -ForegroundColor Green
if ($transcriptConvIds.Keys.Count -gt 0 -and $convTranscripts.Keys.Count -eq 0) {
    Write-Host "No transcripts retrieved. Check 'speechAndTextAnalytics > data > view' permission and that voice transcription / digital transcripts are enabled for these queues." -ForegroundColor Yellow
}

# =============================================================================
# 4b. BACKFILL TriggeringText FROM THE TRANSCRIPT
#     The Copilot suggestions API frequently returns only an intent reference
#     in the trigger - not the spoken/typed text. When TriggeringText is
#     empty, derive it: the customer phrases spoken just BEFORE the suggestion
#     fired are what triggered the article.
# =============================================================================
$backfilled = 0
foreach ($d in $detailRows) {
    if ($d.TriggeringText -ne "") { continue }
    if (-not $convPhrases.ContainsKey($d.ConversationId)) { continue }

    # Customer-side phrases only - those are what Copilot listens to
    $custPhrases = @()
    foreach ($ph in $convPhrases[$d.ConversationId]) {
        if ($ph.Speaker -eq "customer" -or $ph.Speaker -eq "external") { $custPhrases += $ph }
    }
    if ($custPhrases.Count -eq 0) { continue }

    # Reference time: trigger eventTime if present, else when the suggestion was
    # created. Normalised to UTC so it compares correctly with phrase times.
    $refTime = $null
    $refStr = [string]$d.TriggerTimeUtc
    if ($refStr -eq "") { $refStr = [string]$d.SuggestedAtUtc }
    if ($refStr -match "^[0-9]{12,}$") {
        # epoch milliseconds delivered as a bare number
        try {
            $refTime = $epochBase.AddMilliseconds([double]$refStr)
            $refTime = $refTime.ToUniversalTime()
        } catch { $refTime = $null }
    }
    elseif ($refStr -ne "") {
        try {
            $refTime = [datetime]$refStr
            $refTime = $refTime.ToUniversalTime()
        } catch { $refTime = $null }
    }

    $picked = @()
    if ($null -ne $refTime) {
        # Last 2 timestamped customer phrases at/before the suggestion moment
        foreach ($ph in $custPhrases) {
            if ($null -ne $ph.ParsedTime -and $ph.ParsedTime -le $refTime) { $picked += $ph.Text }
        }
        if ($picked.Count -gt 2) {
            $picked = @($picked[($picked.Count - 2)], $picked[($picked.Count - 1)])
        }
    }
    if ($picked.Count -eq 0) {
        # No usable timestamps (e.g. offset-only digital transcripts) - fall back
        # to the customer's opening phrases, the usual reason for contact
        $take = 2
        if ($custPhrases.Count -lt $take) { $take = $custPhrases.Count }
        for ($k = 0; $k -lt $take; $k++) { $picked += $custPhrases[$k].Text }
    }

    if ($picked.Count -gt 0) {
        $d.TriggeringText    = ($picked -join " ")
        $d.TriggerTextSource = "transcript-derived"
        $backfilled++
    }
}
if ($backfilled -gt 0) {
    Write-Host ("TriggeringText backfilled from transcripts for " + $backfilled + " suggestion row(s).") -ForegroundColor Green
}

# =============================================================================
# 5. PRIMARY OUTPUT: ONE ROW PER CONVERSATION
#    Conversation -> agent -> queue -> transcript -> article(s) presented -> why
# =============================================================================
Write-Host ""
Write-Host "Resolving agent and queue names..." -ForegroundColor Cyan

$overviewRows = @()
foreach ($c in $conversationList) {
    $convSuggestions = @()
    foreach ($d in $detailRows) {
        if ($d.ConversationId -eq $c.ConversationId) { $convSuggestions += $d }
    }

    $presented  = "NO"
    $articles   = ""
    $whyParts   = @()
    if ($convSuggestions.Count -gt 0) {
        $presented = "YES"
        $titles = @()
        foreach ($m in $convSuggestions) {
            $t = $m.ArticleTitle
            if ($t -eq "") { $t = $m.DocumentId }
            if ($t -eq "") { $t = "(unidentified " + $m.SuggestionType + " suggestion " + $m.SuggestionId + ")" }
            if ($titles -notcontains $t) { $titles += $t }

            $why = ""
            if ($m.TriggerIntent -ne "") { $why = "intent [" + $m.TriggerIntent + "]" }
            if ($m.TriggeringText -ne "") {
                if ($why -ne "") { $why = $why + " from " }
                $why = $why + "phrase: '" + $m.TriggeringText + "'"
            }
            if ($why -eq "") { $why = "auto-suggest (no trigger detail returned)" }
            $why = $t + " <= " + $why
            if ($whyParts -notcontains $why) { $whyParts += $why }
        }
        $articles = $titles -join "; "
    }

    $agentName = Resolve-AgentName -UserId $c.AgentUserId -AuthHeaders $headers
    $queueName = Resolve-QueueName -QueueId $c.QueueId    -AuthHeaders $headers

    $transcriptText = ""
    if ($convTranscripts.ContainsKey($c.ConversationId)) {
        $transcriptText = $convTranscripts[$c.ConversationId]
    }

    $overviewRows += New-Object PSObject -Property @{
        ConversationId    = $c.ConversationId
        ConversationStart = $c.ConversationStart
        Agent             = $agentName
        Queue             = $queueName
        MediaType         = $c.MediaType
        ArticlePresented  = $presented
        ArticleCount      = $convSuggestions.Count
        Articles          = $articles
        WhyPresented      = ($whyParts -join " || ")
        Transcript        = $transcriptText
    }
}

$overviewRows |
    Sort-Object ConversationStart |
    Select-Object ConversationId, ConversationStart, Agent, Queue, MediaType, ArticlePresented, ArticleCount, Articles, WhyPresented, Transcript |
    Export-Csv -Path $OverviewCsv -NoTypeInformation -Encoding UTF8

Write-Host ("PER-CONVERSATION overview written: " + $OverviewCsv) -ForegroundColor Green

# =============================================================================
# 5a. EXPORT DETAIL CSV (one row per suggestion - the drill-down)
# =============================================================================
if ($detailRows.Count -eq 0) {
    Write-Host "No suggestion rows - skipping detail, transcript and summary CSVs." -ForegroundColor Yellow
    exit 0
}

# Stamp resolved names onto the detail rows too (cached - no extra API calls)
$detailOut = @()
foreach ($d in $detailRows) {
    $detailOut += New-Object PSObject -Property @{
        ConversationId    = $d.ConversationId
        ConversationStart = $d.ConversationStart
        Agent             = (Resolve-AgentName -UserId $d.AgentUserId -AuthHeaders $headers)
        Queue             = (Resolve-QueueName -QueueId $d.QueueId    -AuthHeaders $headers)
        MediaType         = $d.MediaType
        SuggestionId      = $d.SuggestionId
        SuggestionType    = $d.SuggestionType
        SuggestedAtUtc    = $d.SuggestedAtUtc
        TriggerTimeUtc    = $d.TriggerTimeUtc
        State             = $d.State
        Confidence        = $d.Confidence
        TriggerIntent     = $d.TriggerIntent
        KnowledgeBaseId   = $d.KnowledgeBaseId
        DocumentId        = $d.DocumentId
        DocumentVersion   = $d.DocumentVersion
        ArticleTitle      = $d.ArticleTitle
        TriggeringText    = $d.TriggeringText
        TriggerTextSource = $d.TriggerTextSource
    }
}

$detailOut |
    Select-Object ConversationId, ConversationStart, Agent, Queue, MediaType, SuggestionId, SuggestionType, SuggestedAtUtc, TriggerTimeUtc, State, Confidence, TriggerIntent, KnowledgeBaseId, DocumentId, DocumentVersion, ArticleTitle, TriggeringText, TriggerTextSource |
    Export-Csv -Path $DetailCsv -NoTypeInformation -Encoding UTF8

Write-Host ("Detail CSV written: " + $DetailCsv) -ForegroundColor Green

# =============================================================================
# 5b. EXPORT TRANSCRIPT CSV (one row per phrase - the timeline)
# =============================================================================
if ($transcriptRows.Count -gt 0) {
    $transcriptRows |
        Select-Object ConversationId, PhraseTimeUtc, Speaker, Text |
        Export-Csv -Path $TranscriptCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("Transcript CSV written: " + $TranscriptCsv + " (" + $transcriptRows.Count + " phrases)") -ForegroundColor Green
    Write-Host "Join it to the detail CSV on ConversationId, then compare PhraseTimeUtc against SuggestedAtUtc - the customer phrases just BEFORE the suggestion are what triggered the article." -ForegroundColor Cyan
}

# =============================================================================
# 6. BUILD PER-ARTICLE KB IMPROVEMENT SUMMARY
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
Write-Host "Done. Articles with high TimesPresented but low AcceptanceRate are your rewrite candidates." -ForegroundColor Cyan
