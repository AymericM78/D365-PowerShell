Clear-Host; 
# Parameters
$nugetPackage = "D365.DevOps.Powershell";
$nugetUrl = "https://www.nuget.org/api/v2/package";
$packagePath = "$($env:temp)\D365.DevOps.Powershell";
$zipPackagePath = "$packagePath\$nugetPackage.zip";


# Get package registration details
$url = "https://api.nuget.org/v3/registration3/$($nugetPackage.ToLower())/index.json";
$response = Invoke-WebRequest $url -UseBasicParsing;
$packageMetadata =  $response.Content | ConvertFrom-Json;
$latestVersion = $packageMetadata.items.upper;
Write-Host "D365.DevOps.Powershell latest version = $latestVersion" -ForegroundColor Yellow;

$forceVersion = "2020.4.211";
Write-Host "D365.DevOps.Powershell force version = $forceVersion" -ForegroundColor Yellow;
$latestVersion = $forceVersion;

$packagePath = "$packagePath\$latestVersion";

# Handling D365.DevOps.Powershell package
Write-Host "Check if D365.DevOps.Powershell exists..." -NoNewline -ForegroundColor Gray;
if(Test-Path -Path $packagePath)
{ 
    Write-Host "[OK : exist]" -ForegroundColor Green;
}
else
{
    Write-Host "[OK : not exist]" -ForegroundColor Green;

    # Create D365.DevOps.Powershell folder
    Write-Host "Provisionning D365.DevOps.Powershell folder..." -NoNewline -ForegroundColor Gray;
    New-Item -ItemType Directory -Force -Path $packagePath | Out-Null;
    Write-Host "[OK]" -ForegroundColor Green;

    # Download package to D365.DevOps.Powershell folder
    Write-Host "Downloading D365.DevOps.Powershell package from $nugetUrl..." -NoNewline -ForegroundColor Gray;
    Invoke-WebRequest "$nugetUrl/$nugetPackage/$latestVersion" -OutFile $zipPackagePath -UseBasicParsing;
    Expand-Archive -Path $zipPackagePath -DestinationPath $packagePath -Force;
    Remove-Item -Path $zipPackagePath -Force;
    Write-Host "[OK]" -ForegroundColor Green;
}

[Environment]::SetEnvironmentVariable("D365.DevOps.Powershell.Path", $packagePath);
Write-Host "##vso[task.setvariable variable=D365.DevOps.Powershell.Path;]$packagePath";
Write-Host "DevOps variable 'D365.DevOps.Powershell.Path' defined with '$packagePath'";
