param
(
	$login = "admin@domain.com",
    $password = "LeMotDePasse",
    $region = "EMEA", # NorthAmerica, EMEA, APAC, SouthAmerica, Oceania, JPN, CAN, IND, and NorthAmerica2
    $authType = "Office365", # AD, IDF, oAuth, Office365,
    $csvSeparator = ";" # For CSV export
)

Clear-Host;
$ErrorActionPreference = "Stop";
$DebugPreference = "Continue";

# Add lines to prevent progress bar display overlog
Write-Host "";
Write-Host "";
Write-Host "";
Write-Host "";

# Handling XrmTooling package
Write-Host "Checking " -NoNewline -ForegroundColor Gray;
Write-Host "XrmTooling" -NoNewline -ForegroundColor Yellow;
Write-Host " prerequisite..." -NoNewline -ForegroundColor Gray;
$packagePath = "$PSScriptRoot\XrmTooling";
if((Test-Path -Path $packagePath) -eq $false)
{ 
    Write-Host "[OK]" -ForegroundColor Green;
        
    Write-Host "`t> Provisionning XrmTooling folder..." -NoNewline -ForegroundColor Gray;
    New-Item -ItemType Directory -Force -Path $packagePath | Out-Null;
    $tempPath = "$packagePath\temp";
    New-Item -ItemType Directory -Force -Path  $tempPath | Out-Null;
    Write-Host "[OK]" -ForegroundColor Green;

    $nuggetPackage = "Microsoft.CrmSdk.XrmTooling.CrmConnector.PowerShell";
    $nugetUrl = "https://www.nuget.org/api/v2/package";
    $zipPackagePath = "$packagePath\$nuggetPackage.zip";
    
    Write-Host "`t> Downloading XrmTooling package from $nugetUrl..." -NoNewline -ForegroundColor Gray;
    Invoke-WebRequest "$nugetUrl/$nuggetPackage" -OutFile $zipPackagePath;
    Expand-Archive -Path $zipPackagePath -DestinationPath  $tempPath -Force;
    Remove-Item -Path $zipPackagePath -Force;
    Write-Host "[OK]" -ForegroundColor Green;
    
    Write-Host "`t> Cleaning repository..." -NoNewline -ForegroundColor Gray;
    Copy-Item "$tempPath\tools" $packagePath -Recurse -Force;
    Remove-Item -Path $tempPath -Recurse -Force;
    Write-Host "[OK]" -ForegroundColor Green;
}
else
{
    Write-Host "[OK]" -ForegroundColor Green;
}

# Loading XrmTooling
Write-Host "Loading " -NoNewline -ForegroundColor Gray;
Write-Host "XrmTooling" -NoNewline -ForegroundColor Yellow;
Write-Host " module..." -NoNewline -ForegroundColor Gray;
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
Import-Module "$packagePath\tools\Microsoft.Xrm.Tooling.CrmConnector.PowerShell\Microsoft.Xrm.Tooling.CrmConnector.Powershell.dll";
Write-Host "[OK]" -ForegroundColor Green;

# Loading organizations from discovery
Write-Host "Discovering organizations..." -NoNewline -ForegroundColor Gray;
$securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force;
$credentials = New-Object System.Management.Automation.PSCredential($login, $securePassword);
$crmOrganizations = Get-CrmOrganizations -OnLineType Office365 -DeploymentRegion $region -Credential $credentials;
Write-Host "[OK]" -ForegroundColor Green;

# Prepare query for solutions
$querySolutions = New-Object "Microsoft.Xrm.Sdk.Query.QueryExpression" -ArgumentList "solution";
$querySolutions.ColumnSet.AddColumn("uniquename");
$querySolutions.ColumnSet.AddColumn("friendlyname");
$querySolutions.ColumnSet.AddColumn("ismanaged");
$querySolutions.ColumnSet.AddColumn("version");
$querySolutions.ColumnSet.AddColumn("createdon");
$querySolutions.ColumnSet.AddColumn("installedon");
$querySolutions.ColumnSet.AddColumn("modifiedon");
$orderType = [Microsoft.Xrm.Sdk.Query.OrderType]::Ascending;
$order = New-Object "Microsoft.Xrm.Sdk.Query.OrderExpression" -ArgumentList "uniquename", $orderType;
$querySolutions.Orders.Add($order);

$allSolutions = @();
$current = 0;
$total = $crmOrganizations.Count;
foreach($crmOrganization in $crmOrganizations)
{
    $current++;
    $percent = ($current/$total)*100;
	Write-Progress -Activity "Pulling solution data" -Status "[$current/$total] Processing instance '$($crmOrganization.FriendlyName)'" -PercentComplete $percent;

    # Connecting to CRM instance with connection string
    Write-Host "Connecting to " -NoNewline -ForegroundColor Gray;
    Write-Host $crmOrganization.FriendlyName -NoNewline -ForegroundColor Yellow;
    Write-Host " instance..." -NoNewline -ForegroundColor Gray;
	$crmConnectionString = "AuthType=$authType;Username=$login;Password=$password;Url=$($crmOrganization.WebApplicationUrl);RequireNewInstance=true;";
    try
    {
        $crmClient = Get-CrmConnection -ConnectionString $crmConnectionString;
        Write-Host "[OK]" -ForegroundColor Green;
    }
    catch
    {
        Write-Host "[KO] => Reason: $($_.Exception.Message))" -ForegroundColor Red;
        continue;
    }
    
    # Retrieving CRM solutions
    Write-Host "`t> Loading solutions..." -NoNewline -ForegroundColor Gray;
    $solutions = $crmClient.RetrieveMultiple($querySolutions);
    Write-Host "[OK]" -ForegroundColor Green;
    
    # Retrieving CRM build version (thanks R�mi Boigey)
    Write-Host "`t> Retrieving version..." -NoNewline -ForegroundColor Gray;
    $versionResponse = Invoke-WebRequest "$($crmOrganization.WebApplicationUrl)/nga/version.txt" -Method Get;
    $buildVersion = $versionResponse.Content;
    Write-Host "[OK]" -ForegroundColor Green;
    
    # Loading data into custom object collection for final rendering
    Write-Host "`t> Processing data..." -NoNewline -ForegroundColor Gray;
    foreach($solution in $solutions.Entities)
    {
         $solutionObject = New-Object -TypeName psobject;
         $solutionObject | Add-Member -MemberType NoteProperty -Name Instance -Value $crmOrganization.FriendlyName;
         $solutionObject | Add-Member -MemberType NoteProperty -Name InstanceBuild -Value $buildVersion;
         $solutionObject | Add-Member -MemberType NoteProperty -Name InstanceDBVersion -Value $crmOrganization.Version;
         foreach($attribute in $solution.Attributes)
         {
            $solutionObject | Add-Member -MemberType NoteProperty -Name $attribute.Key -Value $attribute.Value;
         }
         $allSolutions  += $solutionObject;
    }
    Write-Host "[OK]" -ForegroundColor Green;
}
write-progress one one -completed;

# Handle selection for CSV output into clipboard
$selectedSolutions = $allSolutions | Out-GridView -OutputMode Multiple;
if($selectedSolutions.Count -eq 0)
{
    Exit;
}
$outputCsv = New-Object "System.Text.StringBuilder";
$headers = $selectedSolutions[0].psobject.properties.name;
foreach($header in $headers)
{
    $outputCsv.Append($header + $csvSeparator) | Out-Null;
}
$outputCsv.AppendLine("") | Out-Null;
foreach($solution in $selectedSolutions)
{
    foreach($header in $headers)
    {
        $value = "";
        try
        {
            $value = $solution.$header;
        }
        catch
        {
        
        }
        $outputCsv.Append([string]::Concat($value, $csvSeparator)) | Out-Null;
    }
    $outputCsv.AppendLine("") | Out-Null;
}
Set-Clipboard -Value $outputCsv.ToString();