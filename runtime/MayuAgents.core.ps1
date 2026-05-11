param(
  [ValidateSet("daily_pulse", "biblia_proyecto", "traspaso_control_operacion", "bodega_materiales", "test")]
  [string]$Mode = "daily_pulse",
  [string]$Date = "",
  [bool]$SendEmail = $true
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"

function Get-RunbookVariable {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [bool]$Required = $true,
    [string]$Default = ""
  )

  $value = $null
  $cmd = Get-Command -Name Get-AutomationVariable -ErrorAction SilentlyContinue
  if ($cmd) {
    try { $value = Get-AutomationVariable -Name $Name } catch { $value = $null }
  }
  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    $value = [Environment]::GetEnvironmentVariable($Name)
  }
  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    $value = $Default
  }
  if ($Required -and [string]::IsNullOrWhiteSpace([string]$value)) {
    throw "Falta variable requerida: $Name"
  }
  [string]$value
}

function ConvertTo-DrivePath {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
  (($Path.Trim("/") -split "/") | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }) -join "/"
}

function Invoke-GraphGet {
  param([string]$Token, [string]$Uri)
  Invoke-RestMethod -Method Get -Uri $Uri -Headers @{ Authorization = "Bearer $Token" }
}

function Invoke-GraphJson {
  param([string]$Token, [string]$Method, [string]$Uri, [object]$Body)
  $json = $Body | ConvertTo-Json -Depth 40
  Invoke-RestMethod -Method $Method -Uri $Uri -Headers @{ Authorization = "Bearer $Token" } -ContentType "application/json" -Body $json
}

function Invoke-GraphPutBytes {
  param([string]$Token, [string]$Uri, [byte[]]$Bytes, [string]$ContentType)
  Invoke-RestMethod -Method Put -Uri $Uri -Headers @{ Authorization = "Bearer $Token" } -ContentType $ContentType -Body $Bytes
}

function Get-GraphToken {
  $tenantId = Get-RunbookVariable "MayuTenantId"
  $clientId = Get-RunbookVariable "MayuClientId"
  $clientSecret = Get-RunbookVariable "MayuClientSecret"
  $token = Invoke-RestMethod `
    -Method Post `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
      client_id = $clientId
      client_secret = $clientSecret
      scope = "https://graph.microsoft.com/.default"
      grant_type = "client_credentials"
    }
  $token.access_token
}

function Get-SiteId {
  param([string]$Token, [string]$HostName)
  (Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$HostName").id
}

function Ensure-GraphFolder {
  param([string]$Token, [string]$SiteId, [string]$FolderPath)
  if ([string]::IsNullOrWhiteSpace($FolderPath)) { return }
  $parts = @($FolderPath.Trim("/") -split "/" | Where-Object { $_ })
  $current = ""
  foreach ($part in $parts) {
    $candidate = if ($current) { "$current/$part" } else { $part }
    $encoded = ConvertTo-DrivePath $candidate
    $exists = $true
    try { Invoke-GraphGet -Token $Token -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" | Out-Null } catch { $exists = $false }
    if (-not $exists) {
      $parentEncoded = ConvertTo-DrivePath $current
      $uri = if ($parentEncoded) {
        "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$parentEncoded" + ":/children"
      } else {
        "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root/children"
      }
      Invoke-GraphJson -Token $Token -Method Post -Uri $uri -Body @{
        name = $part
        folder = @{}
        "@microsoft.graph.conflictBehavior" = "fail"
      } | Out-Null
    }
    $current = $candidate
  }
}

function Write-TextFileToGraph {
  param([string]$Token, [string]$SiteId, [string]$FilePath, [string]$Text, [string]$ContentType)
  $encoded = ConvertTo-DrivePath $FilePath
  $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive/root:/$encoded" + ":/content"
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
  Invoke-GraphPutBytes -Token $Token -Uri $uri -Bytes $bytes -ContentType $ContentType | Out-Null
}

function Send-GraphMail {
  param(
    [string]$Token,
    [string]$Sender,
    [string[]]$To,
    [string[]]$Cc,
    [string]$Subject,
    [string]$HtmlBody
  )
  $message = @{
    subject = $Subject
    body = @{ contentType = "HTML"; content = $HtmlBody }
    toRecipients = @($To | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    ccRecipients = @($Cc | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
  }
  Invoke-GraphJson -Token $Token -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" -Body @{
    message = $message
    saveToSentItems = $true
  } | Out-Null
}

function Get-Mayutime {
  param([string]$TimeZoneName)
  $candidateIds = @($TimeZoneName, "Pacific SA Standard Time", "Chile Standard Time")
  foreach ($id in $candidateIds) {
    if ([string]::IsNullOrWhiteSpace($id)) { continue }
    try {
      $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
      return [System.TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $tz).DateTime
    } catch { continue }
  }
  [DateTime]::UtcNow.AddHours(-4)
}

function HtmlEscape {
  param([object]$Value)
  [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-Number {
  param([object]$Value, [double]$Default = 0)
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $Default }
  try { return [double]$Value } catch { return $Default }
}

function Format-Clp {
  param([double]$Value)
  "$([Math]::Round($Value, 0)) CLP"
}

function Get-ConfigApiKey {
  param([object]$Config)
  $primary = [string]$Config.firebase.api_key_env
  $fallback = [string]$Config.firebase.fallback_api_key_env
  $apiKey = ""
  if ($primary) { $apiKey = [Environment]::GetEnvironmentVariable($primary) }
  if ([string]::IsNullOrWhiteSpace($apiKey) -and $fallback) { $apiKey = [Environment]::GetEnvironmentVariable($fallback) }
  if ([string]::IsNullOrWhiteSpace($apiKey)) {
    throw "Falta variable Firebase API key ($primary). Configurala como GitHub Variable, no como secreto de codigo."
  }
  $apiKey
}

function Get-FirebaseIdToken {
  param([string]$ApiKey)
  $body = @{ returnSecureToken = $true } | ConvertTo-Json
  $resp = Invoke-RestMethod -Method Post -Uri "https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=$ApiKey" -ContentType "application/json" -Body $body -TimeoutSec 30
  [string]$resp.idToken
}

function ConvertFrom-FirestoreValue {
  param([object]$Value)
  if ($null -eq $Value) { return $null }
  $props = $Value.PSObject.Properties.Name
  if ($props -contains "stringValue") { return [string]$Value.stringValue }
  if ($props -contains "integerValue") { return [int64]$Value.integerValue }
  if ($props -contains "doubleValue") { return [double]$Value.doubleValue }
  if ($props -contains "booleanValue") { return [bool]$Value.booleanValue }
  if ($props -contains "timestampValue") { return [string]$Value.timestampValue }
  if ($props -contains "nullValue") { return $null }
  if ($props -contains "arrayValue") {
    $values = @()
    if ($Value.arrayValue.values) {
      foreach ($v in @($Value.arrayValue.values)) { $values += ConvertFrom-FirestoreValue $v }
    }
    return $values
  }
  if ($props -contains "mapValue") {
    $hash = [ordered]@{}
    if ($Value.mapValue.fields) {
      foreach ($p in $Value.mapValue.fields.PSObject.Properties) {
        $hash[$p.Name] = ConvertFrom-FirestoreValue $p.Value
      }
    }
    return [pscustomobject]$hash
  }
  $null
}

function ConvertFrom-FirestoreDocument {
  param([object]$Doc)
  $hash = [ordered]@{}
  if ($Doc.fields) {
    foreach ($p in $Doc.fields.PSObject.Properties) {
      $hash[$p.Name] = ConvertFrom-FirestoreValue $p.Value
    }
  }
  $id = ([string]$Doc.name -split "/")[-1]
  $hash["id"] = $id
  [pscustomobject]$hash
}

function Get-FirestoreCollection {
  param(
    [object]$Config,
    [string]$Token,
    [string]$CollectionName,
    [int]$PageSize = 300
  )
  $projectId = [string]$Config.firebase.project_id
  $database = [string]$Config.firebase.database
  if ([string]::IsNullOrWhiteSpace($database)) { $database = "(default)" }
  $encodedDb = [Uri]::EscapeDataString($database)
  $base = "https://firestore.googleapis.com/v1/projects/$projectId/databases/$encodedDb/documents/$CollectionName"
  $items = @()
  $pageToken = ""
  do {
    $uri = "$base`?pageSize=$PageSize"
    if ($pageToken) { $uri += "&pageToken=$([Uri]::EscapeDataString($pageToken))" }
    $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -TimeoutSec 45
    foreach ($doc in @($resp.documents)) {
      if ($null -ne $doc) { $items += ConvertFrom-FirestoreDocument $doc }
    }
    $pageToken = [string]$resp.nextPageToken
  } while (-not [string]::IsNullOrWhiteSpace($pageToken))
  $items
}

function Get-FirestoreData {
  param([object]$Config)
  $apiKey = Get-ConfigApiKey -Config $Config
  $idToken = Get-FirebaseIdToken -ApiKey $apiKey
  $data = [ordered]@{}
  foreach ($p in $Config.collections.PSObject.Properties) {
    $logical = $p.Name
    $collectionName = [string]$p.Value
    try {
      $data[$logical] = @(Get-FirestoreCollection -Config $Config -Token $idToken -CollectionName $collectionName)
      Write-Host "Firestore: $collectionName -> $(@($data[$logical]).Count) docs."
    } catch {
      Write-Warning "No se pudo leer Firestore '$collectionName'. $($_.Exception.Message)"
      $data[$logical] = @()
    }
  }
  [pscustomobject]$data
}

function Get-DocList {
  param([object]$Project, [string]$Area)
  $areaObj = $null
  if ($Project.areas -and $Project.areas.PSObject.Properties[$Area]) { $areaObj = $Project.areas.$Area }
  if ($areaObj -and $areaObj.docs) { return @($areaObj.docs) }
  @()
}

function Find-ControlDoc {
  param([object]$Project, [object]$Requirement)
  $docs = Get-DocList -Project $Project -Area ([string]$Requirement.area)
  $ids = @($Requirement.doc_ids)
  @($docs | Where-Object { $ids -contains $_.id })
}

function Test-DocApproved {
  param([object]$Doc)
  if ($null -eq $Doc) { return $false }
  ([string]$Doc.status) -in @("Aprobado", "Aprobado con observaciones")
}

function Get-RequirementStatus {
  param([object]$Project, [string]$Key, [object]$Requirement)
  $docs = @(Find-ControlDoc -Project $Project -Requirement $Requirement)
  if (@($docs).Count -eq 0) {
    return [pscustomobject]@{ key = $Key; estado = "rojo"; detalle = "Documento no existe en Control"; owner = [string]$Requirement.owner; version = ""; file = "" }
  }
  $missing = @($docs | Where-Object { -not $_.fileUrl })
  $notApproved = @($docs | Where-Object { -not (Test-DocApproved $_) })
  $versions = @($docs | ForEach-Object { [string]$_.version } | Where-Object { $_ -and $_ -ne "-" })
  $files = @($docs | ForEach-Object { [string]$_.originalFileName } | Where-Object { $_ })
  if (@($missing).Count -gt 0) {
    return [pscustomobject]@{ key = $Key; estado = "rojo"; detalle = "$(@($missing).Count) doc(s) sin archivo"; owner = [string]$Requirement.owner; version = ($versions -join ", "); file = ($files -join ", ") }
  }
  if (@($notApproved).Count -gt 0) {
    return [pscustomobject]@{ key = $Key; estado = "amarillo"; detalle = "$(@($notApproved).Count) doc(s) sin aprobacion final"; owner = [string]$Requirement.owner; version = ($versions -join ", "); file = ($files -join ", ") }
  }
  [pscustomobject]@{ key = $Key; estado = "verde"; detalle = "Aprobado"; owner = [string]$Requirement.owner; version = ($versions -join ", "); file = ($files -join ", ") }
}

function New-Issue {
  param(
    [ValidateSet("rojo", "amarillo", "info")]
    [string]$Severity,
    [string]$Area,
    [string]$Title,
    [string]$Detail,
    [string]$Owner = "",
    [string]$Action = "",
    [string]$Ref = ""
  )
  [pscustomobject]@{
    severity = $Severity
    area = $Area
    title = $Title
    detail = $Detail
    owner = $Owner
    action = $Action
    ref = $Ref
  }
}

function Add-Issue {
  param([System.Collections.ArrayList]$List, [object]$Issue)
  [void]$List.Add($Issue)
}

function Get-IssueKey {
  param([object]$Issue)
  "$($Issue.severity)|$($Issue.area)|$($Issue.title)|$($Issue.ref)"
}

function Select-UniqueIssues {
  param([object[]]$Issues)
  $seen = New-Object System.Collections.Generic.HashSet[string]
  $out = @()
  foreach ($issue in @($Issues)) {
    if ($null -eq $issue) { continue }
    $key = Get-IssueKey -Issue $issue
    if ($seen.Add($key)) { $out += $issue }
  }
  $out
}

function Test-OperationalMatProject {
  param([object]$MatProject, [object[]]$Units, [object[]]$Packs)
  if ($null -eq $MatProject) { return $false }
  if (@($Units | Where-Object { $_.matProjectId -eq $MatProject.id }).Count -gt 0) { return $true }
  if (@($Packs | Where-Object { $_.matProjectId -eq $MatProject.id }).Count -gt 0) { return $true }
  if ((Get-Number $MatProject.totalPods) -gt 0) { return $true }
  if ($MatProject.podsPorTip) {
    foreach ($prop in $MatProject.podsPorTip.PSObject.Properties) {
      if ((Get-Number $prop.Value) -gt 0) { return $true }
    }
  }
  $false
}

function Get-BibliaProjectRows {
  param([object]$Config, [object[]]$ChkProjects)
  $rows = @()
  foreach ($p in @($ChkProjects)) {
    $reqRows = @()
    foreach ($req in $Config.doc_requirements.PSObject.Properties) {
      $reqRows += Get-RequirementStatus -Project $p -Key $req.Name -Requirement $req.Value
    }
    $rojos = @($reqRows | Where-Object { $_.estado -eq "rojo" }).Count
    $amarillos = @($reqRows | Where-Object { $_.estado -eq "amarillo" }).Count
    $estado = if ($rojos -gt 0) { "rojo" } elseif ($amarillos -gt 0) { "amarillo" } else { "verde" }
    $rows += [pscustomobject]@{
      id = $p.id
      name = $p.name
      client = $p.client
      crmId = $p.crmId
      status = $p.status
      estado = $estado
      rojos = $rojos
      amarillos = $amarillos
      requirements = $reqRows
    }
  }
  $rows
}

function Get-WorstStatus {
  param([object[]]$Rows)
  if (@($Rows | Where-Object { $_.estado -eq "rojo" }).Count -gt 0) { return "rojo" }
  if (@($Rows | Where-Object { $_.estado -eq "amarillo" }).Count -gt 0) { return "amarillo" }
  "verde"
}

function Get-GateStatus {
  param([object[]]$Requirements, [string[]]$Keys)
  $rows = @($Requirements | Where-Object { $Keys -contains $_.key })
  [pscustomobject]@{
    estado = Get-WorstStatus -Rows $rows
    bloqueos = @($rows | Where-Object { $_.estado -ne "verde" } | ForEach-Object {
      [pscustomobject]@{
        requisito = $_.key
        estado = $_.estado
        detalle = $_.detalle
        responsable = $_.owner
      }
    })
  }
}

function Build-BibliaReport {
  param([object]$Config, [object[]]$ChkProjects, [datetime]$Now)
  $rows = @(Get-BibliaProjectRows -Config $Config -ChkProjects $ChkProjects)
  $projects = @()
  foreach ($row in $rows) {
    $reqs = @($row.requirements)
    $aprobados = @($reqs | Where-Object { $_.estado -eq "verde" } | ForEach-Object {
      [pscustomobject]@{ requisito = $_.key; version = $_.version; archivo = $_.file }
    })
    $pendientes = @($reqs | Where-Object { $_.estado -ne "verde" } | ForEach-Object {
      [pscustomobject]@{ requisito = $_.key; estado = $_.estado; detalle = $_.detalle; responsable = $_.owner }
    })
    $projects += [pscustomobject]@{
      id = $row.id
      name = $row.name
      client = $row.client
      crmId = $row.crmId
      status = $row.status
      estadoGeneral = $row.estado
      habilitadoParaComprar = Get-GateStatus -Requirements $reqs -Keys @("bom", "gantt_compras", "finanzas")
      habilitadoParaFabricar = Get-GateStatus -Requirements $reqs -Keys @("bom", "planos_fabricacion", "gantt_fabricacion")
      habilitadoParaDespachar = Get-GateStatus -Requirements $reqs -Keys @("plan_despachos", "protocolo_transporte", "calidad")
      habilitadoParaRF = Get-GateStatus -Requirements $reqs -Keys @("calidad", "planos_fabricacion")
      documentosPendientes = $pendientes
      documentosAprobados = $aprobados
      requirements = $reqs
    }
  }
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      proyectos = @($projects).Count
      verdes = @($projects | Where-Object { $_.estadoGeneral -eq "verde" }).Count
      amarillos = @($projects | Where-Object { $_.estadoGeneral -eq "amarillo" }).Count
      rojos = @($projects | Where-Object { $_.estadoGeneral -eq "rojo" }).Count
    }
    projects = $projects
  }
}

function Get-LatestApprovedBomVersion {
  param([object]$ChkProject)
  $bom = @(Find-ControlDoc -Project $ChkProject -Requirement ([pscustomobject]@{ area = "ingenieria"; doc_ids = @("i1") })) | Select-Object -First 1
  if ($bom -and (Test-DocApproved $bom)) { return [string]$bom.version }
  ""
}

function Get-OperationalIssues {
  param([object]$Config, [object]$Data, [object[]]$BibliaRows)
  $issues = [System.Collections.ArrayList]::new()
  $matProjects = @($Data.mat_projects)
  $chkProjects = @($Data.chk_projects)
  $crmProjects = @($Data.crm_projects)
  $recetas = @($Data.fab_proyecto_recetas)
  $units = @($Data.fab_units)
  $packs = @($Data.fab_packs)

  foreach ($mp in $matProjects) {
    $isOperational = Test-OperationalMatProject -MatProject $mp -Units $units -Packs $packs
    if (-not $isOperational) {
      continue
    }
    if (-not $mp.crmId) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Traspaso Control-Operacion" -Title "Proyecto operativo sin vinculo comercial" -Detail "$($mp.name) esta en operacion, pero no esta vinculado al negocio comercial. Finanzas no puede cruzar bien los costos por proyecto." -Owner "Carlos / Valentina" -Action "Vincular el proyecto operativo con su negocio comercial." -Ref $mp.id)
    }
    if (-not $mp.chkId) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Traspaso Control-Operacion" -Title "Proyecto operativo sin biblia vinculada" -Detail "$($mp.name) esta en operacion, pero no esta conectado con Control Documental. No puede sincronizar la biblia del proyecto." -Owner "Carlos / Martin" -Action "Vincularlo al proyecto correcto en Control Documental." -Ref $mp.id)
      continue
    }
    $chk = @($chkProjects | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
    if (-not $chk) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Traspaso Control-Operacion" -Title "Vinculo documental roto" -Detail "$($mp.name) esta conectado a un proyecto documental que ya no existe o no se encontro." -Owner "Carlos / Martin" -Action "Corregir el vinculo con Control Documental." -Ref $mp.id)
      continue
    }
    $bomVigente = Get-LatestApprovedBomVersion -ChkProject $chk
    if ($bomVigente -and $bomVigente -ne [string]$mp.bomVersionUsada) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Control/Materiales" -Title "BOM aprobado pendiente de traspaso" -Detail "$($mp.name): Control Documental tiene BOM $bomVigente y Materiales todavia usa '$($mp.bomVersionUsada)'." -Owner "Carlos" -Action "Sincronizar la BOM aprobada desde Control Documental." -Ref $mp.id)
    }
    $rec = @($recetas | Where-Object { $_.matProjectId -eq $mp.id -or $_.id -eq $mp.id } | Select-Object -First 1)
    if ($mp.bomVersionUsada -and $rec -and ([string]$rec.bomVersionSincronizada) -ne ([string]$mp.bomVersionUsada)) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Materiales/Fabricacion" -Title "Fabricacion no alineada al BOM vigente" -Detail "$($mp.name): Materiales usa BOM $($mp.bomVersionUsada), pero Fabricacion quedo con '$($rec.bomVersionSincronizada)'." -Owner "Carlos / Martin" -Action "Alinear Fabricacion con la BOM vigente y regenerar packs si corresponde." -Ref $mp.id)
    }
    $unitsMp = @($units | Where-Object { $_.matProjectId -eq $mp.id })
    if ([string]$chk.status -eq "Aprobado para ejecucion" -or [string]$chk.status -eq "Aprobado para ejecución") {
      if (@($unitsMp).Count -eq 0) {
        Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Control/Fabricacion" -Title "Proyecto aprobado sin unidades creadas en fabrica" -Detail "$($mp.name) esta aprobado en Control, pero aun no tiene unidades creadas en Fabricacion." -Owner "Martin / Felipe" -Action "Crear las unidades de Fabricacion si el proyecto ya debe entrar a planta." -Ref $mp.id)
      }
    }
    $activeUnits = @($unitsMp | Where-Object { @("M00","M01","M02","M03","M04","M05","RF_FABRICA") -contains [string]$_.status })
    if (@($activeUnits).Count -gt 0) {
      $bib = @($BibliaRows | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
      if ($bib -and $bib.estado -ne "verde") {
        Add-Issue $issues (New-Issue -Severity "rojo" -Area "Control/Fabricacion" -Title "Produccion activa con documentacion incompleta" -Detail "$($mp.name) tiene $(@($activeUnits).Count) unidad(es) activas y la biblia del proyecto sigue en $($bib.estado)." -Owner "Martin / Carlos" -Action "Cerrar documentos criticos o definir formalmente como sigue la produccion." -Ref $mp.id)
      }
    }
    $packsMp = @($packs | Where-Object { $_.matProjectId -eq $mp.id })
    if (@($unitsMp).Count -gt 0 -and @($packsMp).Count -eq 0) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Fabricacion/Packs" -Title "Produccion sin packs generados" -Detail "$($mp.name) tiene unidades creadas, pero aun no tiene packs de materiales." -Owner "Felipe / Mauricio" -Action "Generar packs por modulo desde Fabricacion." -Ref $mp.id)
    }
  }
  @($issues)
}

function Build-TraspasoReport {
  param([object]$Config, [object]$Data, [object[]]$BibliaRows, [datetime]$Now)
  $issues = @(Get-OperationalIssues -Config $Config -Data $Data -BibliaRows $BibliaRows)
  $projects = @()
  foreach ($mp in @($Data.mat_projects)) {
    $isOperational = Test-OperationalMatProject -MatProject $mp -Units @($Data.fab_units) -Packs @($Data.fab_packs)
    if (-not $isOperational) { continue }
    $chk = @($Data.chk_projects | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
    $rec = @($Data.fab_proyecto_recetas | Where-Object { $_.matProjectId -eq $mp.id -or $_.id -eq $mp.id } | Select-Object -First 1)
    $projectIssues = @($issues | Where-Object { $_.ref -eq $mp.id })
    $projects += [pscustomobject]@{
      matProjectId = $mp.id
      name = $mp.name
      crmId = $mp.crmId
      chkId = $mp.chkId
      controlProjectName = if ($chk) { $chk.name } else { "" }
      bomVersionControl = if ($chk) { Get-LatestApprovedBomVersion -ChkProject $chk } else { "" }
      bomVersionMateriales = [string]$mp.bomVersionUsada
      bomVersionFabricacion = if ($rec) { [string]$rec.bomVersionSincronizada } else { "" }
      unidades = @($Data.fab_units | Where-Object { $_.matProjectId -eq $mp.id }).Count
      packs = @($Data.fab_packs | Where-Object { $_.matProjectId -eq $mp.id }).Count
      estado = Get-WorstStatus -Rows @($projectIssues | ForEach-Object {
        [pscustomobject]@{ estado = $_.severity }
      })
      issues = $projectIssues
    }
  }
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      proyectosOperativos = @($projects).Count
      proyectosConAlertas = @($projects | Where-Object { @($_.issues).Count -gt 0 }).Count
      rojos = @($issues | Where-Object { $_.severity -eq "rojo" }).Count
      amarillos = @($issues | Where-Object { $_.severity -eq "amarillo" }).Count
    }
    projects = $projects
    issues = $issues
  }
}

function Get-PackIssues {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = [System.Collections.ArrayList]::new()
  $catalog = @{}
  foreach ($sku in @($Data.inv_catalogo)) { if ($sku.code) { $catalog[[string]$sku.code] = $sku } elseif ($sku.id) { $catalog[[string]$sku.id] = $sku } }
  $activePacks = @($Data.fab_packs | Where-Object { @("planificado","armado","entregado","recibido") -contains [string]$_.estado })
  $packsSinCodigo = [ordered]@{}
  foreach ($pack in $activePacks) {
    $label = "$($pack.projectName) / $($pack.modulo) / $($pack.fecha)"
    $itemsSinSku = @($pack.items | Where-Object { -not $_.skuCode })
    if ($itemsSinSku.Count -gt 0 -or $pack.recetaIncompleta) {
      $projectKey = if ($pack.matProjectId) { [string]$pack.matProjectId } elseif ($pack.projectName) { [string]$pack.projectName } else { [string]$pack.id }
      if (-not $packsSinCodigo.Contains($projectKey)) {
        $packsSinCodigo[$projectKey] = [pscustomobject]@{
          projectName = if ($pack.projectName) { [string]$pack.projectName } else { $projectKey }
          packCount = 0
          itemCount = 0
          examples = [System.Collections.ArrayList]::new()
        }
      }
      $group = $packsSinCodigo[$projectKey]
      $group.packCount += 1
      $group.itemCount += $itemsSinSku.Count
      if ($group.examples.Count -lt 3) {
        [void]$group.examples.Add("$($pack.modulo) / $($pack.fecha)")
      }
    }
    if ([string]$pack.estado -eq "planificado" -and -not $pack.compromisoEntregaBodega) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Packs" -Title "Pack sin compromiso de Bodega" -Detail "$label esta planificado sin fecha comprometida por Bodega." -Owner "Mauricio" -Action "Asignar compromiso de entrega en Bodega." -Ref $pack.id)
    }
    if ([string]$pack.estado -eq "armado") {
      foreach ($it in @($pack.items | Where-Object { $_.skuCode })) {
        $sku = $catalog[[string]$it.skuCode]
        $qty = Get-Number $it.qtyArmada
        if ($qty -eq 0) { $qty = Get-Number $it.qtyPlanificada }
        $stock = if ($sku) { Get-Number $sku.stock } else { 0 }
        if (-not $sku -or $stock -lt $qty) {
          Add-Issue $issues (New-Issue -Severity "rojo" -Area "Packs" -Title "Pack armado con stock insuficiente" -Detail "$label tiene productos con stock insuficiente para entregar." -Owner "Mauricio" -Action "Ingresar stock o recepcionar compra antes de entregar." -Ref $pack.id)
          break
        }
      }
    }
    if ([string]$pack.estado -eq "entregado") {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Packs/Fabricacion" -Title "Pack entregado no recibido" -Detail "$label fue entregado a fabrica y falta recepcion conforme." -Owner "Felipe" -Action "Recibir pack en Fabricacion." -Ref $pack.id)
    }
    if ([string]$pack.estado -eq "recibido") {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Packs/Fabricacion" -Title "Pack recibido no cerrado" -Detail "$label esta recibido y falta cierre." -Owner "Felipe" -Action "Cerrar pack o registrar desvio/devolucion." -Ref $pack.id)
    }
  }
  foreach ($projectKey in $packsSinCodigo.Keys) {
    $group = $packsSinCodigo[$projectKey]
    $examples = @($group.examples)
    $exampleText = if ($examples.Count -gt 0) { " Ejemplos: $($examples -join '; ')." } else { "" }
    $itemText = if ($group.itemCount -gt 0) { "$($group.itemCount) producto(s) sin codigo de bodega" } else { "receta incompleta" }
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Packs" -Title "Packs no armables por productos sin definir" -Detail "$($group.projectName) tiene $($group.packCount) pack(s) no armable(s) por $itemText.$exampleText" -Owner "Carlos" -Action "Completar codificacion y cotizacion antes de pedir armado a Bodega." -Ref $projectKey)
  }
  @($issues)
}

function Get-AbastecimientoIssues {
  param([object]$Config, [object]$Data)
  $issues = [System.Collections.ArrayList]::new()
  $skuDemand = @{}
  $skuDemandPacks = @{}
  foreach ($pack in @($Data.fab_packs | Where-Object { @("planificado","armado") -contains [string]$_.estado })) {
    foreach ($it in @($pack.items | Where-Object { $_.skuCode })) {
      $code = [string]$it.skuCode
      $qty = Get-Number $it.qtyArmada
      if ($qty -eq 0) { $qty = Get-Number $it.qtyPlanificada }
      if (-not $skuDemand.ContainsKey($code)) {
        $skuDemand[$code] = 0.0
        $skuDemandPacks[$code] = [System.Collections.ArrayList]::new()
      }
      $skuDemand[$code] += $qty
      [void]$skuDemandPacks[$code].Add($pack)
    }
  }
  foreach ($mp in @($Data.mat_projects)) {
    if (-not (Test-OperationalMatProject -MatProject $mp -Units @($Data.fab_units) -Packs @($Data.fab_packs))) {
      continue
    }
    $sinSku = @($mp.items | Where-Object { -not $_.skuCode -and ([string]$_.status) -ne "inactivo" })
    if ($sinSku.Count -gt 0) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Abastecimiento" -Title "BOM con productos sin codificar" -Detail "$($mp.name) tiene $($sinSku.Count) producto(s) de la BOM sin codigo de bodega." -Owner "Carlos" -Action "Completar codificacion y cotizacion para poder comprar y armar packs." -Ref $mp.id)
    }
  }
  $stockGaps = @()
  foreach ($sku in @($Data.inv_catalogo | Where-Object { $_.activo -ne $false -and -not $_.aliasDe })) {
    $stock = Get-Number $sku.stock
    $cost = Get-Number $sku.costoPromedio
    $code = if ($sku.code) { $sku.code } else { $sku.id }
    if ($stock -le [double]$Config.thresholds.stock_critical_qty -and $skuDemand.ContainsKey($code)) {
      $stockGaps += [pscustomobject]@{
        code = $code
        demand = $skuDemand[$code]
        packs = @($skuDemandPacks[$code])
      }
    }
  }
  if ($stockGaps.Count -gt 0) {
    $packRefs = New-Object System.Collections.Generic.HashSet[string]
    $projectRefs = New-Object System.Collections.Generic.HashSet[string]
    foreach ($gap in $stockGaps) {
      foreach ($pack in @($gap.packs)) {
        if ($pack.id) { [void]$packRefs.Add([string]$pack.id) }
        if ($pack.projectName) { [void]$projectRefs.Add([string]$pack.projectName) }
      }
    }
    $projectList = @($projectRefs.GetEnumerator() | Select-Object -First 3)
    $projectText = if ($projectList.Count -gt 0) { " Proyectos impactados: $($projectList -join '; ')." } else { "" }
    Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Bodega" -Title "Productos requeridos sin stock suficiente" -Detail "Hay $($stockGaps.Count) producto(s) requeridos por packs proximos con stock bajo o cero. Impacta $($packRefs.Count) pack(s).$projectText" -Owner "Mauricio / Carlos" -Action "Definir si se recepciona compra, se reemplaza material o se reprograma entrega." -Ref "stock_packs")
  }
  $sinValor = @($Data.inv_catalogo | Where-Object {
    $_.activo -ne $false -and
    -not $_.aliasDe -and
    (Get-Number $_.stock) -gt 0 -and
    (Get-Number $_.costoPromedio) -le 0
  })
  if ($sinValor.Count -gt 0) {
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Bodega/Finanzas" -Title "Inventario con costos incompletos" -Detail "$($sinValor.Count) producto(s) con stock tienen costo registrado en cero. Esto distorsiona costos por proyecto y reportes de directorio." -Owner "Mauricio / Valentina" -Action "Valorizar inventario desde Kardex." -Ref "inv_catalogo")
  }
  foreach ($oc in @($Data.mat_ordenes)) {
    if (-not $oc.precioUnit -or [double]$oc.precioUnit -le 0) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Compras" -Title "Orden de compra sin precio" -Detail "$($oc.id) / $($oc.itemDesc) no tiene precio unitario valido." -Owner "Carlos / Valentina" -Action "Completar precio para medir costo comprometido." -Ref $oc.id)
    }
  }
  @($issues)
}

function Get-FirstText {
  param([object]$Object, [string[]]$Names)
  if ($null -eq $Object) { return "" }
  foreach ($name in $Names) {
    $prop = $Object.PSObject.Properties[$name]
    if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
      return [string]$prop.Value
    }
  }
  ""
}

function Normalize-MayuText {
  param([object]$Value)
  $text = ([string]$Value).ToLowerInvariant()
  $text = $text -replace "[^a-z0-9áéíóúñü ]", " "
  ($text -replace "\s+", " ").Trim()
}

function Get-TextSimilarity {
  param([object]$A, [object]$B)
  $aText = Normalize-MayuText $A
  $bText = Normalize-MayuText $B
  if ([string]::IsNullOrWhiteSpace($aText) -or [string]::IsNullOrWhiteSpace($bText)) { return 0.0 }
  if ($aText -eq $bText) { return 1.0 }
  $aWords = @($aText -split " " | Where-Object { $_.Length -gt 2 } | Select-Object -Unique)
  $bWords = @($bText -split " " | Where-Object { $_.Length -gt 2 } | Select-Object -Unique)
  if ($aWords.Count -eq 0 -or $bWords.Count -eq 0) { return 0.0 }
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($w in $aWords) { [void]$set.Add($w) }
  $intersection = 0
  foreach ($w in $bWords) { if ($set.Contains($w)) { $intersection++ } }
  $union = @($aWords + $bWords | Select-Object -Unique).Count
  if ($union -eq 0) { return 0.0 }
  [double]$intersection / [double]$union
}

function New-BodegaIssue {
  param(
    [string]$CheckId,
    [ValidateSet("CRITICO", "ALTO", "MEDIO", "BAJO")]
    [string]$Level,
    [string]$Block,
    [string]$Title,
    [string]$Detail,
    [string]$Owner,
    [string]$Action,
    [string]$Ref
  )
  $severity = if ($Level -eq "CRITICO") { "rojo" } elseif ($Level -in @("ALTO", "MEDIO")) { "amarillo" } else { "info" }
  [pscustomobject]@{
    severity = $severity
    level = $Level
    checkId = $CheckId
    block = $Block
    area = "Bodega + Materiales"
    title = "$CheckId - $Title"
    detail = $Detail
    owner = $Owner
    action = $Action
    ref = $Ref
  }
}

function Add-BodegaIssue {
  param([System.Collections.ArrayList]$List, [object]$Issue)
  [void]$List.Add($Issue)
}

function Get-BodegaBomItems {
  param([object[]]$MatProjects)
  $items = @()
  foreach ($mp in @($MatProjects)) {
    foreach ($it in @($mp.items)) {
      $code = Get-FirstText $it @("bomItemCode", "itemCode", "code", "matchCode", "id")
      $desc = Get-FirstText $it @("descripcion", "description", "desc", "itemDesc", "nombre", "name")
      $skuCode = Get-FirstText $it @("skuCode", "catalogCode", "codigoBodega")
      if ([string]::IsNullOrWhiteSpace($code) -and [string]::IsNullOrWhiteSpace($desc)) { continue }
      $items += [pscustomobject]@{
        matProjectId = $mp.id
        projectName = $mp.name
        code = $code
        skuCode = $skuCode
        desc = $desc
        raw = $it
      }
    }
  }
  $items
}

function Get-BodegaCatalogItems {
  param([object[]]$Catalog)
  $items = @()
  foreach ($sku in @($Catalog)) {
    $code = Get-FirstText $sku @("code", "skuCode", "id")
    $legacy = Get-FirstText $sku @("codigoLegacy", "legacyCode", "bomItemCode")
    $desc = Get-FirstText $sku @("descripcion", "description", "desc", "nombre", "name")
    $items += [pscustomobject]@{ code = $code; legacy = $legacy; desc = $desc; raw = $sku }
  }
  $items
}

function New-BodegaBomSkuIndex {
  param([object[]]$BomItems)
  $byProjectItem = @{}
  $byItem = @{}
  foreach ($b in @($BomItems)) {
    if ([string]::IsNullOrWhiteSpace([string]$b.code)) { continue }
    if ($b.matProjectId) { $byProjectItem["$($b.matProjectId)|$($b.code)"] = $b }
    if (-not $byItem.ContainsKey([string]$b.code)) { $byItem[[string]$b.code] = $b }
  }
  [pscustomobject]@{ byProjectItem = $byProjectItem; byItem = $byItem }
}

function Resolve-BodegaBomSku {
  param([object]$Index, [string]$ProjectId, [string]$ItemCode)
  if ([string]::IsNullOrWhiteSpace($ItemCode)) { return "" }
  $bomItem = $null
  if (-not [string]::IsNullOrWhiteSpace($ProjectId)) {
    $key = "$ProjectId|$ItemCode"
    if ($Index.byProjectItem.ContainsKey($key)) { $bomItem = $Index.byProjectItem[$key] }
  }
  if ($null -eq $bomItem -and $Index.byItem.ContainsKey($ItemCode)) { $bomItem = $Index.byItem[$ItemCode] }
  if ($null -eq $bomItem) { return "" }
  Get-FirstText $bomItem @("skuCode", "catalogCode", "codigoBodega")
}

function Build-BodegaMaterialesReport {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = [System.Collections.ArrayList]::new()
  $bomItems = @(Get-BodegaBomItems -MatProjects @($Data.mat_projects))
  $bomSkuIndex = New-BodegaBomSkuIndex -BomItems $bomItems
  $catalogItems = @(Get-BodegaCatalogItems -Catalog @($Data.inv_catalogo))
  $bomCodes = New-Object System.Collections.Generic.HashSet[string]
  foreach ($b in $bomItems) { if ($b.code) { [void]$bomCodes.Add([string]$b.code) } }
  $catalogCodes = New-Object System.Collections.Generic.HashSet[string]
  $legacyCodes = New-Object System.Collections.Generic.HashSet[string]
  foreach ($c in $catalogItems) {
    if ($c.code) { [void]$catalogCodes.Add([string]$c.code) }
    if ($c.legacy) { [void]$legacyCodes.Add([string]$c.legacy) }
  }
  $similarThreshold = Get-Number $Config.thresholds.bodega_materiales_similarity_threshold 0.7
  $dupThreshold = Get-Number $Config.thresholds.bodega_materiales_duplicate_threshold 0.85
  $deliveryDays = [int](Get-Number $Config.thresholds.bodega_materiales_delivery_days 7)
  $staleOcDays = [int](Get-Number $Config.thresholds.bodega_materiales_oc_stale_days 30)
  $priceDiffPct = Get-Number $Config.thresholds.bodega_materiales_price_diff_pct 0.05
  $today = $Now.Date

  foreach ($oc in @($Data.mat_ordenes)) {
    $ocId = Get-FirstText $oc @("id", "folio", "ocId")
    $projectId = Get-FirstText $oc @("matProjectId", "projectId", "proyectoId")
    $itemCode = Get-FirstText $oc @("bomItemCode", "itemCode", "matchCode", "code")
    $proveedor = Get-FirstText $oc @("proveedor", "supplierName", "vendor", "razonSocial")
    $ocSkuCode = Get-FirstText $oc @("skuCode", "catalogCode", "codigoBodega")
    $bomSkuCode = Resolve-BodegaBomSku -Index $bomSkuIndex -ProjectId $projectId -ItemCode $itemCode
    $skuCode = if ($ocSkuCode) { $ocSkuCode } else { $bomSkuCode }
    if ([string]::IsNullOrWhiteSpace($projectId) -or [string]::IsNullOrWhiteSpace($itemCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0001" "CRITICO" "Trazabilidad" "OC sin identidad BOM completa" "OC $ocId no tiene matProjectId y bomItemCode trazables." "Carlos" "Regularizar el vinculo OC -> proyecto -> item BOM antes de nuevas recepciones." $ocId)
    } elseif (-not $bomCodes.Contains($itemCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0002" "CRITICO" "Trazabilidad" "OC con item fuera del BOM activo" "OC $ocId usa item $itemCode, pero ese codigo no aparece en los BOM activos leidos." "Carlos" "Corregir itemCode o registrar excepcion autorizada por Carlos." $ocId)
    }
    if ([string]::IsNullOrWhiteSpace($skuCode)) {
      $itemDesc = Get-FirstText $oc @("itemDesc", "descripcion", "description", "nombre", "name")
      Add-BodegaIssue $issues (New-BodegaIssue "B-0003" "CRITICO" "Trazabilidad" "OC sin SKU candidato de bodega" "OC $ocId no tiene SKU de bodega asociado antes de recepcionar. Proyecto=$projectId ItemBOM=$itemCode Desc='$itemDesc' Proveedor='$proveedor'." "Carlos / Mauricio" "Linkear a SKU candidato antes de recepcion o dejar excepcion autorizada." $ocId)
    } elseif (-not $catalogCodes.Contains($skuCode)) {
      $itemDesc = Get-FirstText $oc @("itemDesc", "descripcion", "description", "nombre", "name")
      $source = if ($ocSkuCode) { "OC" } else { "BOM" }
      Add-BodegaIssue $issues (New-BodegaIssue "B-0003" "CRITICO" "Trazabilidad" "OC con SKU que no existe en bodega" "OC $ocId resuelve SKU $skuCode desde $source, pero no existe en inv_catalogo. Proyecto=$projectId ItemBOM=$itemCode Desc='$itemDesc'." "Carlos / Mauricio" "Corregir link a SKU existente o crear el codigo real antes de recepcionar." $ocId)
    }
    if (-not (Get-FirstText $oc @("cotizId", "cotizacionId", "quoteId"))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-D04" "ALTO" "Ordenes de compra" "OC sin cotizacion asociada" "OC $ocId no tiene cotizacion vinculada." "Carlos / Valentina" "Vincular cotizacion formal o justificar compra directa." $ocId)
    }
    $proveedor = Get-FirstText $oc @("proveedor", "supplierName", "vendor", "razonSocial")
    if ($proveedor -match "(?i)cotizaci[oó]n|mayu tx|proveedor|pendiente") {
      Add-BodegaIssue $issues (New-BodegaIssue "B-D06" "MEDIO" "Ordenes de compra" "OC con proveedor generico" "OC $ocId tiene proveedor '$proveedor'." "Carlos / Valentina" "Identificar proveedor real para trazabilidad y reclamos." $ocId)
    }
    $fechaOc = Get-FirstText $oc @("fecha", "createdAt", "fechaOC")
    $status = [string](Get-FirstText $oc @("status", "estado"))
    if ($fechaOc) {
      try {
        $age = ($today - ([datetime]::Parse($fechaOc)).Date).Days
        if ($age -gt $staleOcDays -and $status -in @("adjuntada", "emitida", "abierta")) {
          Add-BodegaIssue $issues (New-BodegaIssue "B-D05" "MEDIO" "Ordenes de compra" "OC antigua sin recepcion" "OC $ocId lleva $age dias en estado $status." "Carlos" "Confirmar entrega proveedor o cerrar/reprogramar OC." $ocId)
        }
      } catch {}
    }
  }

  foreach ($rec in @($Data.mat_recepciones)) {
    $recId = Get-FirstText $rec @("id", "recepcionId")
    $ocId = Get-FirstText $rec @("ocId", "ordenId", "ordenCompraId")
    $skuCode = Get-FirstText $rec @("skuCode", "catalogCode", "code", "codigoBodega")
    $itemCode = Get-FirstText $rec @("bomItemCode", "itemCode", "matchCode")
    if ([string]::IsNullOrWhiteSpace($skuCode)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0003" "CRITICO" "Trazabilidad" "Recepcion sin SKU real de bodega" "Recepcion $recId no tiene SKU real; no entra trazablemente al inventario." "Mauricio" "Linkear recepcion a SKU existente o crear SKU con justificacion y codigoLegacy." $recId)
    }
    if ([string]::IsNullOrWhiteSpace($itemCode) -and [string]::IsNullOrWhiteSpace($ocId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Recepcion sin OC ni item BOM" "Recepcion $recId no conserva OC ni item BOM." "Mauricio / Carlos" "Completar ocId y bomItemCode para que el costo llegue al proyecto." $recId)
    }
    $precioFactura = Get-Number (Get-FirstText $rec @("precioUnitFactura", "costoUnit", "precioFactura", "precioUnit"))
    $precioOc = Get-Number (Get-FirstText $rec @("precioUnitOC", "precioUnitOc", "ocPrecioUnit"))
    if ($precioFactura -gt 0 -and $precioOc -gt 0) {
      $diff = [Math]::Abs($precioFactura - $precioOc) / [Math]::Max($precioOc, 1)
      if ($diff -gt $priceDiffPct -and -not $rec.alertaConciliacion) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-E08" "ALTO" "Recepciones" "Precio factura distinto a OC sin conciliacion" "Recepcion $recId tiene precio factura $precioFactura vs OC $precioOc." "Felix / Valentina / Carlos / Mauricio" "Abrir conciliacion: reclamo, nota de credito, ajuste o aceptar diferencia." $recId)
      }
    }
  }

  $receivedByOc = @{}
  foreach ($rec in @($Data.mat_recepciones)) {
    $ocId = Get-FirstText $rec @("ocId", "ordenId", "ordenCompraId")
    if (-not $ocId) { continue }
    if (-not $receivedByOc.ContainsKey($ocId)) { $receivedByOc[$ocId] = 0.0 }
    $receivedByOc[$ocId] += Get-Number (Get-FirstText $rec @("qtyRecibida", "qty", "cantidad"))
  }
  foreach ($oc in @($Data.mat_ordenes)) {
    $ocId = Get-FirstText $oc @("id", "folio", "ocId")
    $qty = Get-Number (Get-FirstText $oc @("qty", "cantidad", "cantidadPedida"))
    $status = [string](Get-FirstText $oc @("status", "estado"))
    $received = if ($receivedByOc.ContainsKey($ocId)) { $receivedByOc[$ocId] } else { 0.0 }
    if ($qty -gt 0 -and $status -in @("recibida_completa", "completa", "total", "recibida_total") -and $received -lt ($qty * 0.99)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-D01" "CRITICO" "Ordenes de compra" "OC completa con recepcion parcial" "OC $ocId esta cerrada como completa, pero recepciones suman $received de $qty." "Carlos / Mauricio" "Cambiar a recepcion parcial y reabrir saldo pendiente." $ocId)
    }
  }

  foreach ($mov in @($Data.inv_movimientos)) {
    $movId = Get-FirstText $mov @("id", "movementId")
    $tipo = Get-FirstText $mov @("tipo", "type")
    $skuCode = Get-FirstText $mov @("skuCode", "code", "itemCode")
    $projectId = Get-FirstText $mov @("matProjectId", "projectId", "proyectoRef", "proyectoId")
    $ocId = Get-FirstText $mov @("ocId", "ordenId")
    $recId = Get-FirstText $mov @("recepcionId", "receiptId")
    $itemCode = Get-FirstText $mov @("bomItemCode", "itemCode", "matchCode")
    if ($mov.refDoc) {
      if (-not $ocId) { $ocId = Get-FirstText $mov.refDoc @("ocId", "ordenId") }
      if (-not $recId) { $recId = Get-FirstText $mov.refDoc @("recepcionId", "receiptId") }
      if (-not $itemCode) { $itemCode = Get-FirstText $mov.refDoc @("bomItemCode", "itemCode") }
    }
    if ($tipo -in @("recepcion_oc", "ingreso_directo") -and ([string]::IsNullOrWhiteSpace($projectId) -or [string]::IsNullOrWhiteSpace($skuCode))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Movimiento sin identidad completa" "Movimiento $movId tipo $tipo no conserva proyecto/SKU suficientes." "Mauricio" "Completar trazabilidad en movimiento o corregir recepcion origen." $movId)
    }
    if ($tipo -eq "recepcion_oc" -and ([string]::IsNullOrWhiteSpace($ocId) -or [string]::IsNullOrWhiteSpace($recId) -or [string]::IsNullOrWhiteSpace($itemCode))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0005" "CRITICO" "Trazabilidad" "Kardex sin OC/recepcion/BOM" "Movimiento $movId no hereda ocId, recepcionId y bomItemCode." "Mauricio / Felix" "Corregir app para heredar campos y regularizar movimiento." $movId)
    }
    if ($tipo -eq "ingreso_directo") {
      $desc = Get-FirstText $mov @("descSku", "description", "descripcion", "itemDesc")
      $similarBom = @($bomItems | Where-Object { (Get-TextSimilarity $desc $_.desc) -ge $similarThreshold } | Select-Object -First 1)
      if ($similarBom.Count -gt 0) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-E02" "CRITICO" "Recepciones" "Ingreso directo parece item BOM activo" "Movimiento $movId ingreso_directo '$desc' se parece a BOM $($similarBom[0].code) / $($similarBom[0].projectName)." "Mauricio / Carlos" "Revisar si debio recepcionarse contra OC/BOM en vez de crear identidad paralela." $movId)
      }
    }
  }

  foreach ($c in $catalogItems) {
    if ($c.legacy -and -not $bomCodes.Contains($c.legacy)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-A01" "CRITICO" "Catalogo" "SKU apunta a BOM inexistente" "$($c.code) tiene codigoLegacy $($c.legacy), pero ese codigo no existe en BOM activo." "Mauricio / Carlos" "Corregir codigoLegacy o archivar SKU si es legado." $c.code)
    }
    if (-not $c.legacy -and $c.desc) {
      $similarBom = @($bomItems | Where-Object { (Get-TextSimilarity $c.desc $_.desc) -ge $similarThreshold } | Select-Object -First 1)
      if ($similarBom.Count -gt 0) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-0004" "CRITICO" "Trazabilidad" "SKU nuevo similar a BOM sin link" "$($c.code) '$($c.desc)' se parece a BOM $($similarBom[0].code) / $($similarBom[0].projectName), pero no tiene codigoLegacy." "Mauricio / Carlos" "Linkear a BOM existente o justificar SKU nuevo." $c.code)
      }
    }
    $stock = Get-Number $c.raw.stock
    if ($stock -gt 0 -and -not $c.legacy -and -not $c.raw.refDoc -and -not $c.raw.origen) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0006" "ALTO" "Trazabilidad" "SKU con stock sin origen trazable" "$($c.code) tiene stock $stock sin codigoLegacy/origen claro." "Mauricio / Valentina" "Regularizar origen del stock para valorizar y atribuir proyecto." $c.code)
    }
    $std = Get-Number $c.raw.costoEstandar
    $avg = Get-Number $c.raw.costoPromedio
    if ($std -gt 0 -and $avg -gt 0) {
      $diff = [Math]::Abs($std - $avg) / [Math]::Max($avg, 1)
      if ($diff -gt 0.5) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-A03" "ALTO" "Catalogo" "Costo estandar muy distinto al promedio" "$($c.code): costoEstandar $std vs costoPromedio $avg." "Mauricio / Valentina" "Revisar valorizacion desde factura/guia y kardex." $c.code)
      }
    }
    if ($avg -le 1 -and $stock -gt 0) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-A05" "MEDIO" "Catalogo" "SKU con costo placeholder y stock" "$($c.code) tiene stock $stock y costoPromedio $avg." "Mauricio / Valentina" "Valorizar desde documento fisico recepcionado." $c.code)
    }
  }

  for ($i = 0; $i -lt $catalogItems.Count; $i++) {
    for ($j = $i + 1; $j -lt $catalogItems.Count; $j++) {
      if ([string]::IsNullOrWhiteSpace($catalogItems[$i].desc) -or [string]::IsNullOrWhiteSpace($catalogItems[$j].desc)) { continue }
      if ((Get-TextSimilarity $catalogItems[$i].desc $catalogItems[$j].desc) -ge $dupThreshold) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-A04" "ALTO" "Catalogo" "SKUs posiblemente duplicados" "$($catalogItems[$i].code) y $($catalogItems[$j].code) tienen descripciones casi identicas." "Mauricio / Carlos" "Unificar identidad o documentar diferencia fisica." "$($catalogItems[$i].code)|$($catalogItems[$j].code)")
        if (@($issues | Where-Object { $_.checkId -eq "B-A04" }).Count -ge 10) { break }
      }
    }
  }

  foreach ($b in $bomItems) {
    $hasStrict = ($b.code -and ($legacyCodes.Contains($b.code) -or $catalogCodes.Contains($b.code))) -or ($b.skuCode -and $catalogCodes.Contains($b.skuCode))
    if (-not $hasStrict) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-B01" "CRITICO" "BOM-Catalogo" "Item BOM sin match estricto en catalogo" "$($b.projectName): $($b.code) '$($b.desc)' no tiene codigoLegacy/code/SKU estricto en catalogo." "Carlos / Mauricio" "Crear o linkear SKU de bodega antes de comprar/recepcionar." "$($b.matProjectId)|$($b.code)")
    }
  }

  foreach ($cot in @($Data.mat_cotizaciones)) {
    $cotId = Get-FirstText $cot @("id", "cotizId")
    $file = Get-FirstText $cot @("fileName", "filename", "pdfUrl", "archivo", "documentUrl")
    if ($file -and $file -match "(?i)\.xlsx?$") {
      Add-BodegaIssue $issues (New-BodegaIssue "B-C01" "CRITICO" "Cotizaciones" "Cotizacion solo en Excel interno" "Cotizacion $cotId parece Excel, no PDF formal del proveedor." "Carlos / Valentina" "Adjuntar PDF formal con proveedor, RUT, fecha y validez." $cotId)
    }
    foreach ($it in @($cot.items)) {
      $matchCode = Get-FirstText $it @("matchCode", "bomItemCode", "itemCode")
      $matchScore = Get-Number (Get-FirstText $it @("matchScore", "score"))
      if ([string]::IsNullOrWhiteSpace($matchCode)) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-C05" "MEDIO" "Cotizaciones" "Item cotizado sin match BOM" "Cotizacion $cotId tiene item sin matchCode." "Carlos" "Asignar item BOM o descartar linea de cotizacion." $cotId)
      } elseif ($matchScore -gt 0 -and $matchScore -lt 90) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-C04" "MEDIO" "Cotizaciones" "Match de cotizacion dudoso" "Cotizacion $cotId item $matchCode tiene matchScore $matchScore." "Carlos" "Validar manualmente match de item cotizado contra BOM." "$cotId|$matchCode")
      }
    }
  }

  foreach ($mp in @($Data.mat_projects)) {
    $skuCancelados = @()
    foreach ($x in @($mp.skuCancelados)) { if ($x) { $skuCancelados += [string]$x } }
    if ($mp.recetaFabInternaPorTipologia -and $mp.recetaFabInternaPorTipologia.skuCancelados) {
      foreach ($x in @($mp.recetaFabInternaPorTipologia.skuCancelados)) { if ($x) { $skuCancelados += [string]$x } }
    }
    foreach ($sku in @($skuCancelados | Select-Object -Unique)) {
      $hasOc = @($Data.mat_ordenes | Where-Object { (Get-FirstText $_ @("bomItemCode", "itemCode", "matchCode", "code")) -eq $sku }).Count -gt 0
      if ($hasOc) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-G03" "ALTO" "Fabricacion interna" "Item fab interna con OC externa" "$($mp.name): $sku esta en skuCancelados pero tiene OC." "Martin / Carlos" "Decidir si se fabrica o se compra; no ambos." "$($mp.id)|$sku")
      }
    }
  }

  foreach ($ent in @($Data.mat_entregas)) {
    $entId = Get-FirstText $ent @("id", "entregaId")
    $skuCode = Get-FirstText $ent @("skuCode", "itemCode", "code")
    $projectId = Get-FirstText $ent @("matProjectId", "projectId", "proyectoId")
    $podNum = Get-FirstText $ent @("podNum", "pod", "unitCode")
    $tipologia = Get-FirstText $ent @("tipologia", "tipo")
    if ([string]::IsNullOrWhiteSpace($skuCode) -or [string]::IsNullOrWhiteSpace($projectId)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-0007" "CRITICO" "Entregas a fabrica" "Entrega sin trazabilidad a SKU/proyecto" "Entrega $entId no conserva SKU o proyecto." "Carlos / Mauricio" "Regularizar entrega para atribuir consumo a proyecto/POD." $entId)
    }
    if ([string]::IsNullOrWhiteSpace($podNum) -or [string]::IsNullOrWhiteSpace($tipologia)) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-F03" "MEDIO" "Entregas a fabrica" "Entrega sin POD o tipologia" "Entrega $entId no tiene POD/tipologia completos." "Mauricio" "Completar POD y tipologia para consumo real por unidad." $entId)
    }
    if ($ent.esMerma -eq $true -and -not (Get-FirstText $ent @("motivoMerma", "nota", "observacion"))) {
      Add-BodegaIssue $issues (New-BodegaIssue "B-F02" "ALTO" "Entregas a fabrica" "Merma sin motivo" "Entrega $entId marcada como merma no tiene motivo documentado." "Mauricio / Carlos" "Documentar causa de merma y responsable." $entId)
    }
  }

  $deliveryKeys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($ent in @($Data.mat_entregas)) {
    $sku = Get-FirstText $ent @("skuCode", "itemCode", "code")
    $proj = Get-FirstText $ent @("matProjectId", "projectId", "proyectoId")
    if ($sku -and $proj) { [void]$deliveryKeys.Add("$proj|$sku") }
  }
  foreach ($rec in @($Data.mat_recepciones)) {
    $recId = Get-FirstText $rec @("id", "recepcionId")
    $sku = Get-FirstText $rec @("skuCode", "catalogCode", "code", "codigoBodega")
    $proj = Get-FirstText $rec @("matProjectId", "projectId", "proyectoId")
    $dateText = Get-FirstText $rec @("fecha", "createdAt", "fechaRecepcion")
    if (-not $sku -or -not $proj -or -not $dateText) { continue }
    try {
      $age = ($today - ([datetime]::Parse($dateText)).Date).Days
      if ($age -gt $deliveryDays -and -not $deliveryKeys.Contains("$proj|$sku")) {
        Add-BodegaIssue $issues (New-BodegaIssue "B-F01" "CRITICO" "Entregas a fabrica" "Recepcion sin entrega formal" "Recepcion $recId / SKU $sku lleva $age dias sin entrega formal a fabrica." "Carlos / Mauricio" "Registrar entrega formal o justificar stock retenido." $recId)
      }
    } catch {}
  }

  $selected = @(Select-UniqueIssues -Issues @($issues))
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      criticas = @($selected | Where-Object { $_.level -eq "CRITICO" }).Count
      altas = @($selected | Where-Object { $_.level -eq "ALTO" }).Count
      medias = @($selected | Where-Object { $_.level -eq "MEDIO" }).Count
      bajas = @($selected | Where-Object { $_.level -eq "BAJO" }).Count
      total = @($selected).Count
    }
    blocks = [pscustomobject]@{
      trazabilidad = @($selected | Where-Object { $_.block -eq "Trazabilidad" })
      catalogo = @($selected | Where-Object { $_.block -eq "Catalogo" })
      bomCatalogo = @($selected | Where-Object { $_.block -eq "BOM-Catalogo" })
      cotizaciones = @($selected | Where-Object { $_.block -eq "Cotizaciones" })
      ordenesCompra = @($selected | Where-Object { $_.block -eq "Ordenes de compra" })
      recepciones = @($selected | Where-Object { $_.block -eq "Recepciones" })
      entregasFabrica = @($selected | Where-Object { $_.block -eq "Entregas a fabrica" })
      fabricacionInterna = @($selected | Where-Object { $_.block -eq "Fabricacion interna" })
    }
    issues = $selected
  }
}

function Write-BodegaPprPvcCandidates {
  param([object]$Data)
  $catalogItems = @(Get-BodegaCatalogItems -Catalog @($Data.inv_catalogo))
  $bomItems = @(Get-BodegaBomItems -MatProjects @($Data.mat_projects))
  $bomSkuIndex = New-BodegaBomSkuIndex -BomItems $bomItems
  $targets = @()
  foreach ($oc in @($Data.mat_ordenes)) {
    $projectId = Get-FirstText $oc @("matProjectId", "projectId", "proyectoId")
    $itemCode = Get-FirstText $oc @("bomItemCode", "itemCode", "matchCode", "code")
    $skuCode = Get-FirstText $oc @("skuCode", "catalogCode", "codigoBodega")
    if (-not $skuCode) { $skuCode = Resolve-BodegaBomSku -Index $bomSkuIndex -ProjectId $projectId -ItemCode $itemCode }
    if ($skuCode) { continue }
    $desc = Get-FirstText $oc @("itemDesc", "descripcion", "description", "nombre", "name")
    if ($desc -notmatch "(?i)\b(PPR|PVC|ABRAZADERA|TUBER|CODO|COPLA|TEE|TERMINAL|LLAVE DE PASO)\b") { continue }
    $targets += [pscustomobject]@{
      oc = Get-FirstText $oc @("id", "folio", "ocId")
      item = $itemCode
      desc = $desc
      proveedor = Get-FirstText $oc @("proveedor", "supplierName", "vendor", "razonSocial")
      project = $projectId
    }
  }
  Write-Output "Bodega+Materiales PPR/PVC: targets=$($targets.Count)"
  foreach ($target in @($targets | Select-Object -First 80)) {
    $candidates = @(
      $catalogItems | ForEach-Object {
        $score = Get-TextSimilarity $target.desc $_.desc
        $bonus = 0.0
        if ($_.legacy -and $target.item -and $_.legacy -eq $target.item) { $bonus += 1.0 }
        if ($_.code -and $target.item -and $_.code -eq $target.item) { $bonus += 1.0 }
        if ($target.desc -match "(?i)PPR" -and $_.desc -match "(?i)PPR") { $bonus += 0.2 }
        if ($target.desc -match "(?i)PVC" -and $_.desc -match "(?i)PVC") { $bonus += 0.2 }
        [pscustomobject]@{
          score = $score + $bonus
          code = $_.code
          legacy = $_.legacy
          desc = $_.desc
          stock = Get-Number $_.raw.stock
          costo = Get-Number $_.raw.costoPromedio
        }
      } | Sort-Object score -Descending | Select-Object -First 5
    )
    $candidateText = (@($candidates) | ForEach-Object {
      "$($_.code)|legacy=$($_.legacy)|stock=$($_.stock)|score=$([Math]::Round($_.score,2))|$($_.desc)"
    }) -join " || "
    Write-Output "Bodega+Materiales PPR/PVC candidate: oc=$($target.oc) item=$($target.item) desc=$($target.desc) proveedor=$($target.proveedor) :: $candidateText"
  }
}

function Render-BodegaMaterialesHtml {
  param([object]$Report)
  $s = $Report.summary
  $rows = Render-IssueList -Items @($Report.issues | Select-Object -First 80)
  @"
<html>
<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;font-size:14px;">
  <h2 style="margin-bottom:4px;">Agente Bodega + Materiales - $($Report.date)</h2>
  <p style="margin-top:0;color:#555;">Criticas: <strong style="color:#991b1b;">$($s.criticas)</strong> · Altas: <strong style="color:#9a3412;">$($s.altas)</strong> · Medias: <strong>$($s.medias)</strong> · Bajas: <strong>$($s.bajas)</strong></p>
  <h3>Alertas</h3>
  $rows
  <p style="font-size:12px;color:#777;">Destinatarios productivos esperados: Felix, Valentina, Carlos y Mauricio. Este HTML es evidencia del agente; el envio queda controlado por SendEmail.</p>
</body>
</html>
"@
}

function Invoke-BodegaMateriales {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now, [bool]$DoSendEmail)
  Write-Output "Bodega+Materiales: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  $report = Build-BodegaMaterialesReport -Config $Config -Data $data -Now $Now
  Write-BodegaPprPvcCandidates -Data $data
  Write-Output "Bodega+Materiales: resumen criticas=$($report.summary.criticas) altas=$($report.summary.altas) medias=$($report.summary.medias) bajas=$($report.summary.bajas) total=$($report.summary.total)."
  foreach ($group in @($report.issues | Group-Object checkId | Sort-Object Count -Descending | Select-Object -First 20)) {
    Write-Output "Bodega+Materiales breakdown: $($group.Name)=$($group.Count)"
  }
  foreach ($group in @($report.issues | Group-Object checkId | Sort-Object Count -Descending | Select-Object -First 12)) {
    foreach ($issue in @($group.Group | Select-Object -First 5)) {
      Write-Output "Bodega+Materiales ejemplo $($group.Name): [$($issue.level)] ref=$($issue.ref) detalle=$($issue.detail) accion=$($issue.action)"
    }
  }
  foreach ($issue in @($report.issues | Where-Object { $_.checkId -eq "B-0003" } | Select-Object -First 200)) {
    Write-Output "Bodega+Materiales B-0003 full: ref=$($issue.ref) detalle=$($issue.detail)"
  }
  foreach ($issue in @($report.issues | Select-Object -First 25)) {
    Write-Output "Bodega+Materiales alerta: [$($issue.level)] $($issue.checkId) $($issue.title) | $($issue.ref) | $($issue.action)"
  }
  $html = Render-BodegaMaterialesHtml -Report $report
  $dateKey = $report.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.bodega_materiales_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.bodega_materiales_folder)/$dateKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.bodega_materiales_folder)/$dateKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  if ($DoSendEmail) {
    $to = @($Config.mail.felix, $Config.mail.valentina, $Config.mail.carlos, $Config.mail.mauricio)
    $subject = "[Bodega-Materiales] Alertas $dateKey - $($report.summary.criticas) CRITICAS, $($report.summary.altas) ALTAS"
    Send-GraphMail -Token $GraphToken -Sender $Config.mail.sender -To $to -Cc @() -Subject $subject -HtmlBody $html
    Write-Output "Bodega+Materiales: correo enviado."
  } else {
    Write-Output "Bodega+Materiales: SendEmail=false, no se envia correo."
  }
  Write-Output "Bodega+Materiales: outputs guardados en SharePoint."
}

function Get-FinanceIssues {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $issues = [System.Collections.ArrayList]::new()
  $today = $Now.ToString("yyyy-MM-dd")
  $apSinProyecto = @()
  $apSinClasificar = @()
  $apVencidas = @()
  $arSinProyecto = @()
  $arVencidas = @()
  foreach ($f in @($Data.fin_facturas_ap)) {
    $estado = [string]$f.estado
    if (@("PAGADA","ANULADA","RECHAZADA") -contains $estado) { continue }
    if ((-not $f.proyectoId) -and (-not $f.asignaciones) -and ([string]$f.lineaNegocio) -ne "OPEX_CORP" -and ([string]$f.categoriaContable) -notin @("OPEX_FIJO","OPEX_VARIABLE")) {
      $apSinProyecto += $f
    }
    if ([string]$f.lineaNegocio -eq "SIN_CLASIFICAR") {
      $apSinClasificar += $f
    }
    if ($f.fechaVencimiento -and [string]$f.fechaVencimiento -lt $today) {
      $apVencidas += $f
    }
  }
  foreach ($f in @($Data.fin_facturas_ar)) {
    $estado = [string]$f.estado
    if (@("COBRADA","ANULADA") -contains $estado) { continue }
    if ((-not $f.crmProjectId) -and (-not $f.proyectoId) -and (-not $f.asignaciones)) {
      $arSinProyecto += $f
    }
    if ($f.fechaVencimiento -and [string]$f.fechaVencimiento -lt $today) {
      $arVencidas += $f
    }
  }
  $maxOverdue = [int]$Config.thresholds.max_finance_overdue_items
  if ($maxOverdue -le 0) { $maxOverdue = 8 }
  foreach ($f in @($apVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -First $maxOverdue)) {
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Caja" -Title "CxP vencida" -Detail "$($f.razonSocialContraparte) folio $($f.folio) vencio $($f.fechaVencimiento), monto $(Format-Clp (Get-Number $f.montoTotal))." -Owner "Valentina / Felix" -Action "Decidir pago o renegociacion." -Ref $f.id)
  }
  if ($apVencidas.Count -gt $maxOverdue) {
    $rest = @($apVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -Skip $maxOverdue)
    $restMonto = ($rest | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Caja" -Title "CxP vencida adicional agrupada" -Detail "$($rest.Count) factura(s) AP vencidas adicionales por aprox. $(Format-Clp $restMonto)." -Owner "Valentina / Felix" -Action "Revisar lista completa en Finanzas." -Ref "fin_facturas_ap")
  }
  foreach ($f in @($arVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -First $maxOverdue)) {
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Caja" -Title "CxC vencida" -Detail "$($f.razonSocialContraparte) folio $($f.folio) vencio $($f.fechaVencimiento), monto $(Format-Clp (Get-Number $f.montoTotal))." -Owner "Valentina / Comercial" -Action "Gestionar cobranza." -Ref $f.id)
  }
  if ($arVencidas.Count -gt $maxOverdue) {
    $rest = @($arVencidas | Sort-Object { -1 * (Get-Number $_.montoTotal) } | Select-Object -Skip $maxOverdue)
    $restMonto = ($rest | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-Issue -Severity "rojo" -Area "Caja" -Title "CxC vencida adicional agrupada" -Detail "$($rest.Count) factura(s) AR vencidas adicionales por aprox. $(Format-Clp $restMonto)." -Owner "Valentina / Comercial" -Action "Revisar lista completa en Finanzas." -Ref "fin_facturas_ar")
  }
  if ($apSinProyecto.Count -gt 0) {
    $monto = ($apSinProyecto | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Finanzas" -Title "Facturas AP sin proyecto agrupadas" -Detail "$($apSinProyecto.Count) factura(s) AP sin proyecto por aprox. $(Format-Clp $monto). Pueden superponerse con sin clasificar." -Owner "Valentina" -Action "Clasificar por lote en Finanzas para destrabar costos por proyecto." -Ref "fin_facturas_ap")
  }
  if ($apSinClasificar.Count -gt 0) {
    $monto = ($apSinClasificar | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Finanzas" -Title "Facturas AP sin clasificar agrupadas" -Detail "$($apSinClasificar.Count) factura(s) AP en SIN_CLASIFICAR por aprox. $(Format-Clp $monto)." -Owner "Valentina" -Action "Clasificar linea/cuenta/proyecto por lote." -Ref "fin_facturas_ap")
  }
  if ($arSinProyecto.Count -gt 0) {
    $monto = ($arSinProyecto | ForEach-Object { Get-Number $_.montoTotal } | Measure-Object -Sum).Sum
    Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Finanzas" -Title "Facturas AR sin proyecto agrupadas" -Detail "$($arSinProyecto.Count) factura(s) AR sin proyecto por aprox. $(Format-Clp $monto)." -Owner "Valentina" -Action "Vincular a CRM/proyecto para lectura comercial y directorio." -Ref "fin_facturas_ar")
  }
  $ocPressure = 0.0
  foreach ($oc in @($Data.mat_ordenes | Where-Object { ([string]$_.status) -notin @("total","recibida_total","cerrada","anulada") })) {
    $ocPressure += ((Get-Number $oc.qty) * (Get-Number $oc.precioUnit))
  }
  if ($ocPressure -gt 0) {
    Add-Issue $issues (New-Issue -Severity "info" -Area "Caja" -Title "Presion de OCs sobre caja" -Detail "OCs abiertas comprometen aprox. $([Math]::Round($ocPressure,0)) CLP." -Owner "Felix / Valentina" -Action "Contrastar contra flujo 13 semanas." -Ref "mat_ordenes")
  }
  @($issues)
}

function Get-CalidadIssues {
  param([object]$Config, [object]$Data, [object[]]$BibliaRows)
  $issues = [System.Collections.ArrayList]::new()
  foreach ($u in @($Data.fab_units)) {
    if ([string]$u.status -eq "M05" -and -not $u.rfFabricaOk) {
      Add-Issue $issues (New-Issue -Severity "amarillo" -Area "Calidad/RF" -Title "POD listo para RF" -Detail "$($u.unitCode) / $($u.matProjectName) esta en M05 sin RF Fabrica OK." -Owner "Felipe / Calidad" -Action "Coordinar RF." -Ref $u.id)
    }
  }
  foreach ($mp in @($Data.mat_projects)) {
    if (-not $mp.chkId) { continue }
    $bib = @($BibliaRows | Where-Object { $_.id -eq $mp.chkId } | Select-Object -First 1)
    if (-not $bib) { continue }
    $cal = @($bib.requirements | Where-Object { $_.key -eq "calidad" } | Select-Object -First 1)
    $desp = @($bib.requirements | Where-Object { $_.key -eq "plan_despachos" } | Select-Object -First 1)
    $trans = @($bib.requirements | Where-Object { $_.key -eq "protocolo_transporte" } | Select-Object -First 1)
    $unitsDespacho = @($Data.fab_units | Where-Object { $_.matProjectId -eq $mp.id -and ([string]$_.status) -in @("RF_FABRICA","QC","TERMINADO") })
    if ($unitsDespacho.Count -gt 0 -and (($cal.estado -ne "verde") -or ($desp.estado -ne "verde") -or ($trans.estado -ne "verde"))) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Calidad/Despacho" -Title "Riesgo de despacho sin documentacion completa" -Detail "$($mp.name) tiene $($unitsDespacho.Count) unidad(es) post RF/QC/PT y docs calidad/despacho/transporte pendientes." -Owner "Carlos / Gabriel / Calidad" -Action "Cerrar calidad, plan despacho y protocolo transporte antes de despacho." -Ref $mp.id)
    }
  }
  @($issues)
}

function Get-CommercialIssues {
  param([object]$Data)
  $issues = [System.Collections.ArrayList]::new()
  $closed = @($Data.crm_projects | Where-Object { $_.estado_comercial -eq "Negocio cerrado" })
  $matCrmIds = New-Object System.Collections.Generic.HashSet[string]
  foreach ($mp in @($Data.mat_projects)) { if ($mp.crmId) { [void]$matCrmIds.Add([string]$mp.crmId) } }
  foreach ($p in $closed) {
    if (-not $matCrmIds.Contains([string]$p.id)) {
      Add-Issue $issues (New-Issue -Severity "rojo" -Area "Comercial" -Title "Cerrado sin operacion" -Detail "$($p.nombre) esta cerrado en CRM y no aparece en Materiales." -Owner "Felix Escudero Vargas" -Action "Traspasar a Control/Materiales." -Ref $p.id)
    }
  }
  $plan = @($Data.crm_plan | Select-Object -First 1)
  if ($plan) {
    $meta = Get-Number $plan.meta_ventas
    $cerrado = 0.0
    foreach ($p in $closed) { $cerrado += Get-Number $p.ingreso_proyectado }
    if ($meta -gt 0) {
      $gap = $meta - $cerrado
      Add-Issue $issues (New-Issue -Severity "info" -Area "Comercial" -Title "Brecha contra meta anual" -Detail "Cerrado: $([Math]::Round($cerrado,0)) CLP / Meta: $([Math]::Round($meta,0)) CLP / Brecha: $([Math]::Round($gap,0)) CLP." -Owner "Felix Escudero Vargas" -Action "Revisar pipeline y cierres prioritarios." -Ref "plan/default")
    }
  }
  @($issues)
}

function Limit-Issues {
  param([object[]]$Issues, [int]$Max)
  @($Issues | Select-Object -First $Max)
}

function Get-DecisionGroupKey {
  param([object]$Issue)
  $ref = [string]$Issue.ref
  if ([string]::IsNullOrWhiteSpace($ref)) {
    return "$($Issue.area)|$($Issue.title)"
  }
  if ($ref -match '^pack-(PRY-[A-Z0-9]+)') { return $matches[1] }
  $ref
}

function Get-EntityDisplayName {
  param([object]$Entity, [string]$Fallback)
  if ($null -eq $Entity) { return $Fallback }
  foreach ($prop in @("name", "nombre", "projectName", "razonSocialContraparte", "title", "titulo", "client", "cliente")) {
    $property = $Entity.PSObject.Properties[$prop]
    if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
      return [string]$property.Value
    }
  }
  $Fallback
}

function Get-InvoiceDisplayName {
  param([object]$Invoice, [string]$Fallback)
  $name = Get-EntityDisplayName -Entity $Invoice -Fallback ""
  $folio = ""
  if ($Invoice -and $Invoice.PSObject.Properties["folio"]) { $folio = [string]$Invoice.folio }
  if (-not [string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($folio)) {
    return "$name folio $folio"
  }
  if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }
  if (-not [string]::IsNullOrWhiteSpace($folio)) { return "Folio $folio" }
  $Fallback
}

function Get-DecisionLabels {
  param([object]$Data)
  $labels = @{}
  if ($Data) {
    foreach ($mp in @($Data.mat_projects)) {
      if ($mp.id) { $labels[[string]$mp.id] = Get-EntityDisplayName -Entity $mp -Fallback ([string]$mp.id) }
    }
    foreach ($chk in @($Data.chk_projects)) {
      if ($chk.id) { $labels[[string]$chk.id] = Get-EntityDisplayName -Entity $chk -Fallback ([string]$chk.id) }
    }
    foreach ($crm in @($Data.crm_projects)) {
      if ($crm.id) { $labels[[string]$crm.id] = Get-EntityDisplayName -Entity $crm -Fallback ([string]$crm.id) }
    }
    foreach ($ap in @($Data.fin_facturas_ap)) {
      if ($ap.id) { $labels[[string]$ap.id] = Get-InvoiceDisplayName -Invoice $ap -Fallback ([string]$ap.id) }
    }
    foreach ($ar in @($Data.fin_facturas_ar)) {
      if ($ar.id) { $labels[[string]$ar.id] = Get-InvoiceDisplayName -Invoice $ar -Fallback ([string]$ar.id) }
    }
  }
  $labels["inv_catalogo"] = "Inventario"
  $labels["stock_packs"] = "Stock para packs proximos"
  $labels["fin_facturas_ap"] = "Facturas por pagar"
  $labels["fin_facturas_ar"] = "Facturas por cobrar"
  $labels
}

function Select-DecisionItems {
  param([object[]]$Issues, [object]$Data = $null, [int]$Max = 8)
  $labels = Get-DecisionLabels -Data $Data
  $groups = [ordered]@{}
  foreach ($issue in @($Issues)) {
    if ($null -eq $issue) { continue }
    $key = Get-DecisionGroupKey -Issue $issue
    if (-not $groups.Contains($key)) { $groups[$key] = @() }
    $groups[$key] += $issue
  }
  $items = @()
  foreach ($key in $groups.Keys) {
    $list = @($groups[$key])
    $first = $list[0]
    $titles = @($list | ForEach-Object { $_.title } | Select-Object -Unique)
    $label = if ($labels.ContainsKey($key)) { $labels[$key] } else { $key }
    $text = if ($list.Count -gt 1) {
      "${label}: $($list.Count) temas abiertos ($($titles -join '; ')). $($first.action)"
    } else {
      "${label}: $($first.title). $($first.action)"
    }
    $items += [pscustomobject]@{ text = $text; owner = $first.owner; ref = $first.ref }
    if ($items.Count -ge $Max) { break }
  }
  $items
}

function Build-Pulse {
  param([object]$Config, [object]$Data, [datetime]$Now)
  $bibliaRows = @(Get-BibliaProjectRows -Config $Config -ChkProjects @($Data.chk_projects))
  $traspaso = @(Get-OperationalIssues -Config $Config -Data $Data -BibliaRows $bibliaRows)
  $packIssues = @(Get-PackIssues -Config $Config -Data $Data -Now $Now)
  $abastIssues = @(Get-AbastecimientoIssues -Config $Config -Data $Data)
  $finIssues = @(Get-FinanceIssues -Config $Config -Data $Data -Now $Now)
  $comIssues = @(Get-CommercialIssues -Data $Data)
  $calIssues = @(Get-CalidadIssues -Config $Config -Data $Data -BibliaRows $bibliaRows)

  $max = [int]$Config.thresholds.max_items_per_section
  $all = @(Select-UniqueIssues -Issues @($traspaso + $packIssues + $abastIssues + $finIssues + $comIssues + $calIssues))
  $rojos = @($all | Where-Object { $_.severity -eq "rojo" })
  $amarillos = @($all | Where-Object { $_.severity -eq "amarillo" })
  $infos = @($all | Where-Object { $_.severity -eq "info" })
  $decisiones = @(Select-DecisionItems -Issues $rojos -Data $Data -Max 8)
  [pscustomobject]@{
    generatedAt = $Now.ToString("o")
    date = $Now.ToString("yyyy-MM-dd")
    summary = [pscustomobject]@{
      rojos = @($rojos).Count
      amarillos = @($amarillos).Count
      infos = @($infos).Count
      proyectosBiblia = @($bibliaRows).Count
      bibliaRoja = @($bibliaRows | Where-Object { $_.estado -eq "rojo" }).Count
      bibliaAmarilla = @($bibliaRows | Where-Object { $_.estado -eq "amarillo" }).Count
    }
    decisionesFelixHoy = @($decisiones)
    sections = [pscustomobject]@{
      bibliaProyecto = Limit-Issues -Issues @($bibliaRows | Where-Object { $_.estado -ne "verde" } | ForEach-Object {
        New-Issue -Severity $_.estado -Area "Biblia del Proyecto" -Title "$($_.name) - Documentacion $($_.estado)" -Detail "$($_.rojos) requisito(s) criticos / $($_.amarillos) requisito(s) pendientes." -Owner "Martin / Carlos / Valentina" -Action "Cerrar requisitos pendientes en Control Documental." -Ref $_.id
      }) -Max $max
      traspasoControlOperacion = Limit-Issues -Issues $traspaso -Max $max
      packs = Limit-Issues -Issues $packIssues -Max $max
      abastecimiento = Limit-Issues -Issues $abastIssues -Max $max
      finanzas = Limit-Issues -Issues $finIssues -Max $max
      comercial = Limit-Issues -Issues $comIssues -Max $max
      calidadRfDespacho = Limit-Issues -Issues $calIssues -Max $max
      compromisosReuniones = @()
    }
    raw = [pscustomobject]@{
      biblia = $bibliaRows
      allIssues = $all
    }
  }
}

function Render-IssueList {
  param([object[]]$Items)
  if (@($Items).Count -eq 0) { return "<p style='color:#166534;margin:6px 0;'>Sin alertas relevantes.</p>" }
  $html = "<table style='border-collapse:collapse;width:100%;font-size:13px;'><tr><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Sev</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Tema</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Accion</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Resp.</th></tr>"
  foreach ($it in @($Items)) {
    $color = if ($it.severity -eq "rojo") { "#991b1b" } elseif ($it.severity -eq "amarillo") { "#9a3412" } else { "#2563eb" }
    $html += "<tr>"
    $html += "<td style='border:1px solid #ddd;padding:6px;color:$color;font-weight:bold;'>$(HtmlEscape $it.severity)</td>"
    $html += "<td style='border:1px solid #ddd;padding:6px;'><strong>$(HtmlEscape $it.title)</strong><br><span style='color:#555;'>$(HtmlEscape $it.detail)</span><br><span style='font-size:12px;color:#777;'>$(HtmlEscape $it.ref)</span></td>"
    $html += "<td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $it.action)</td>"
    $html += "<td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $it.owner)</td>"
    $html += "</tr>"
  }
  $html += "</table>"
  $html
}

function Render-PulseHtml {
  param([object]$Config, [object]$Pulse)
  $s = $Pulse.summary
  $links = $Config.sharepoint.links
  $decisionHtml = if (@($Pulse.decisionesFelixHoy).Count -eq 0) {
    "<li>No hay decisiones rojas detectadas con la data disponible.</li>"
  } else {
    (@($Pulse.decisionesFelixHoy) | ForEach-Object { "<li><strong>$(HtmlEscape $_.owner):</strong> $(HtmlEscape $_.text)</li>" }) -join "`n"
  }
  @"
<html>
<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;font-size:14px;">
  <h2 style="margin-bottom:4px;">Pulso gerencial MAYU - $($Pulse.date)</h2>
  <p style="margin-top:0;color:#555;">Rojos: <strong style="color:#991b1b;">$($s.rojos)</strong> · Amarillos: <strong style="color:#9a3412;">$($s.amarillos)</strong> · Biblia roja/amarilla: <strong>$($s.bibliaRoja)/$($s.bibliaAmarilla)</strong></p>

  <h3>Decisiones Felix hoy</h3>
  <ol>
    $decisionHtml
  </ol>

  <h3>Riesgos rojos</h3>
  $(Render-IssueList -Items @($Pulse.raw.allIssues | Where-Object { $_.severity -eq "rojo" } | Select-Object -First 12))

  <h3>Amarillos</h3>
  $(Render-IssueList -Items @($Pulse.raw.allIssues | Where-Object { $_.severity -eq "amarillo" } | Select-Object -First 12))

  <h3>Biblia del Proyecto</h3>
  $(Render-IssueList -Items @($Pulse.sections.bibliaProyecto))

  <h3>Traspaso Control -> Operacion</h3>
  $(Render-IssueList -Items @($Pulse.sections.traspasoControlOperacion))

  <h3>Packs</h3>
  $(Render-IssueList -Items @($Pulse.sections.packs))

  <h3>Abastecimiento / Compras</h3>
  $(Render-IssueList -Items @($Pulse.sections.abastecimiento))

  <h3>Finanzas</h3>
  $(Render-IssueList -Items @($Pulse.sections.finanzas))

  <h3>Comercial</h3>
  $(Render-IssueList -Items @($Pulse.sections.comercial))

  <h3>Calidad / RF / Despacho</h3>
  $(Render-IssueList -Items @($Pulse.sections.calidadRfDespacho))

  <p style="margin-top:24px;font-size:12px;color:#777;">
    Apps: <a href="$($links.control)">Control</a> · <a href="$($links.materiales)">Materiales</a> · <a href="$($links.bodega)">Bodega</a> · <a href="$($links.fabricacion)">Fabricacion</a> · <a href="$($links.finanzas)">Finanzas</a> · <a href="$($links.crm)">CRM</a>
  </p>
  <p style="font-size:12px;color:#777;">Generado por MAYU Agents. Secciones parciales no inventan datos; solo reportan lo que existe en Firestore.</p>
</body>
</html>
"@
}

function Render-AgentSummaryHtml {
  param([string]$Title, [object]$Report)
  $summaryRows = ""
  foreach ($p in $Report.summary.PSObject.Properties) {
    $summaryRows += "<tr><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $p.Name)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $p.Value)</td></tr>"
  }
  $projectRows = ""
  foreach ($p in @($Report.projects | Select-Object -First 30)) {
    $estado = if ($p.PSObject.Properties["estadoGeneral"]) { $p.estadoGeneral } else { $p.estado }
    $alertas = if ($p.PSObject.Properties["issues"]) { @($p.issues).Count } elseif ($p.PSObject.Properties["documentosPendientes"]) { @($p.documentosPendientes).Count } else { 0 }
    $projectRows += "<tr><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $p.name)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $estado)</td><td style='border:1px solid #ddd;padding:6px;'>$(HtmlEscape $alertas)</td></tr>"
  }
  @"
<html>
<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;font-size:14px;">
  <h2 style="margin-bottom:4px;">$(HtmlEscape $Title) - $($Report.date)</h2>
  <h3>Resumen</h3>
  <table style='border-collapse:collapse;width:100%;font-size:13px;'>$summaryRows</table>
  <h3>Proyectos</h3>
  <table style='border-collapse:collapse;width:100%;font-size:13px;'>
    <tr><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Proyecto</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Estado</th><th style='text-align:left;border:1px solid #ddd;padding:6px;'>Alertas/Pendientes</th></tr>
    $projectRows
  </table>
  <p style="font-size:12px;color:#777;">Detalle completo en JSON en SharePoint.</p>
</body>
</html>
"@
}

function Invoke-BibliaProyecto {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now)
  Write-Output "Biblia: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  $report = Build-BibliaReport -Config $Config -ChkProjects @($data.chk_projects) -Now $Now
  $dateKey = $report.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.biblia_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.biblia_folder)/$dateKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.biblia_folder)/$dateKey.html" -Text (Render-AgentSummaryHtml -Title "Agente Biblia del Proyecto" -Report $report) -ContentType "text/html; charset=utf-8"
  Write-Output "Biblia: outputs guardados en SharePoint."
}

function Invoke-TraspasoControlOperacion {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now)
  Write-Output "Traspaso: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  $bibliaRows = @(Get-BibliaProjectRows -Config $Config -ChkProjects @($data.chk_projects))
  $report = Build-TraspasoReport -Config $Config -Data $data -BibliaRows $bibliaRows -Now $Now
  $dateKey = $report.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.traspaso_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.traspaso_folder)/$dateKey.json" -Text ($report | ConvertTo-Json -Depth 80) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.traspaso_folder)/$dateKey.html" -Text (Render-AgentSummaryHtml -Title "Agente Traspaso Control-Operacion" -Report $report) -ContentType "text/html; charset=utf-8"
  Write-Output "Traspaso: outputs guardados en SharePoint."
}

function Invoke-DailyPulse {
  param([object]$Config, [string]$GraphToken, [string]$SiteId, [datetime]$Now, [bool]$DoSendEmail)
  Write-Output "Pulso: leyendo Firestore."
  $data = Get-FirestoreData -Config $Config
  Write-Output "Pulso: construyendo analisis."
  $pulse = Build-Pulse -Config $Config -Data $data -Now $Now
  $html = Render-PulseHtml -Config $Config -Pulse $pulse
  $json = $pulse | ConvertTo-Json -Depth 80
  $dateKey = $pulse.date
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.pulso_folder
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.biblia_folder
  Ensure-GraphFolder -Token $GraphToken -SiteId $SiteId -FolderPath $Config.sharepoint.traspaso_folder
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.pulso_folder)/$dateKey.html" -Text $html -ContentType "text/html; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.pulso_folder)/$dateKey.json" -Text $json -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.biblia_folder)/$dateKey.json" -Text ($pulse.raw.biblia | ConvertTo-Json -Depth 60) -ContentType "application/json; charset=utf-8"
  Write-TextFileToGraph -Token $GraphToken -SiteId $SiteId -FilePath "$($Config.sharepoint.traspaso_folder)/$dateKey.json" -Text ($pulse.sections.traspasoControlOperacion | ConvertTo-Json -Depth 30) -ContentType "application/json; charset=utf-8"
  if ($DoSendEmail) {
    $subject = "Pulso gerencial MAYU - $dateKey - R$($pulse.summary.rojos) A$($pulse.summary.amarillos)"
    Send-GraphMail -Token $GraphToken -Sender $Config.mail.sender -To @($Config.mail.felix) -Cc @() -Subject $subject -HtmlBody $html
    Write-Output "Pulso: correo enviado a $($Config.mail.felix)."
  } else {
    Write-Output "Pulso: SendEmail=false, no se envia correo."
  }
  Write-Output "Pulso: outputs guardados en SharePoint."
}

$configJson = Get-RunbookVariable "MayuAgentsConfigJson"
$config = $configJson | ConvertFrom-Json
$now = if ($Date) { [datetime]::Parse($Date) } else { Get-Mayutime -TimeZoneName $config.timezone }
$graphToken = Get-GraphToken
$siteId = Get-SiteId -Token $graphToken -HostName $config.sharepoint.host

Write-Output "MAYU Agents iniciado. Modo=$Mode Fecha=$($now.ToString("yyyy-MM-dd")) SendEmail=$SendEmail"
Ensure-GraphFolder -Token $graphToken -SiteId $siteId -FolderPath $config.sharepoint.base_folder

if ($Mode -eq "test") {
  $body = "<html><body><p>MAYU Agents operativo.</p><p>Hora MAYU: $($now.ToString("yyyy-MM-dd HH:mm"))</p></body></html>"
  if ($SendEmail) {
    Send-GraphMail -Token $graphToken -Sender $config.mail.sender -To @($config.mail.felix) -Cc @() -Subject "Prueba MAYU Agents" -HtmlBody $body
    Write-Output "Prueba enviada a $($config.mail.felix)."
  } else {
    Write-Output "Prueba OK sin correo."
  }
} elseif ($Mode -eq "daily_pulse") {
  Invoke-DailyPulse -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
} elseif ($Mode -eq "biblia_proyecto") {
  Invoke-BibliaProyecto -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now
} elseif ($Mode -eq "traspaso_control_operacion") {
  Invoke-TraspasoControlOperacion -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now
} elseif ($Mode -eq "bodega_materiales") {
  Invoke-BodegaMateriales -Config $config -GraphToken $graphToken -SiteId $siteId -Now $now -DoSendEmail $SendEmail
}
