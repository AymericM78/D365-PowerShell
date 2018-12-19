Clear-Host;

function Provision-Folder
{
    PARAM
    (
        [parameter(Mandatory=$true)]$folderPath
    )

	if(Test-Path $folderPath)
	{
		Remove-Item -Recurse -Force $folderPath;
	}
    New-Item -ItemType Directory -Force -Path $folderPath | Out-Null;
}

# Folder hierarchy
$crmSdkPath = "$PSScriptRoot\CrmSdk";
Provision-Folder -folderPath $crmSdkPath;

$packagesPath = "$crmSdkPath\packages";
Provision-Folder -folderPath $packagesPath;
$toolsPath = "$crmSdkPath\tools";
Provision-Folder -folderPath $toolsPath;
$assembliesPath = "$crmSdkPath\assemblies";
Provision-Folder -folderPath $assembliesPath;
$samplesPath = "$crmSdkPath\samples";
Provision-Folder -folderPath $samplesPath;

# Parameters
$nugetUrl = "https://www.nuget.org/api/v2/package";
$defaultFrameworkFolder = "net452";

$packages = @();
# https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/org-service/subscribe-sdk-assembly-updates-using-nuget#BKMK_GetNuGetPackages
$packages += "Tool;Microsoft.CrmSdk.CoreTools";
$packages += "Assembly;Microsoft.CrmSdk.Deployment";
$packages += "Assembly;Microsoft.CrmSdk.Outlook";
$packages += "Assembly;Microsoft.CrmSdk.Workflow";
$packages += "Assembly;Microsoft.CrmSdk.CoreAssemblies";
$packages += "Tool;Microsoft.CrmSdk.XrmTooling.CrmConnector.PowerShell";
$packages += "Tool;Microsoft.CrmSdk.XrmTooling.PackageDeployment.PowerShell";
$packages += "Tool;Microsoft.CrmSdk.XrmTooling.PluginRegistrationTool";
$packages += "Tool;Microsoft.CrmSdk.XrmTooling.ConfigurationMigration.Wpf";
$packages += "Assembly;Microsoft.CrmSdk.Outlook";
$packages += "Sample;Microsoft.CrmSdk.Samples.HelperCode-CS";
$packages += "Sample;Microsoft.CrmSdk.WebApi.Samples.HelperCode";
$packages += "Tool;Microsoft.CrmSdk.XrmTooling.PackageDeployment.Wpf";
$packages += "Assembly;Microsoft.CrmSdk.XrmTooling.PackageDeployment";
$packages += "Assembly;Microsoft.CrmSdk.XrmTooling.WpfControls";


$output = "";
foreach($packageInfos in $packages)
{
    $packageType = $packageInfos.Split(";")[0];
    $packageName = $packageInfos.Split(";")[1];

    $zipPackagePath = "$packagesPath\$packageName.zip";
    $unzipPackagePath = "$packagesPath\$packageName";

    Write-Host "Package '$packageName' : " -ForegroundColor Yellow;

    Write-Host " > Retrieving package..." -NoNewline -ForegroundColor Gray;
    Invoke-WebRequest "$nugetUrl/$packageName" -OutFile $zipPackagePath;
    Write-Host "[OK]" -ForegroundColor Green;

    Write-Host " > Unziping package..." -NoNewline -ForegroundColor Gray;
    Expand-Archive -Path $zipPackagePath -DestinationPath $unzipPackagePath -Force;
    Write-Host "[OK]" -ForegroundColor Green;

    Write-Host " > Removing zip package..." -NoNewline -ForegroundColor Gray;
    Remove-Item -Path $zipPackagePath -Force;
    Write-Host "[OK]" -ForegroundColor Green;
    
    $sourcePaths = @();
    $targetPath = "";
    if($packageType -eq "Tool")
    {
        $targetPath = "$toolsPath\$packageName";
        $sourcePaths += "$unzipPackagePath\tools";
        $sourcePaths += "$unzipPackagePath\content\bin";
    }
    elseif($packageType -eq "Assembly")
    {
        $targetPath = "$assembliesPath\$packageName";
        $sourcePaths += "$unzipPackagePath\lib\$defaultFrameworkFolder";
    }
    elseif($packageType -eq "Sample")
    {
        $targetPath = "$samplesPath\$packageName";
        $sourcePaths += "$unzipPackagePath\content";
    }

    Write-Host " > Locating package content..." -NoNewline -ForegroundColor Gray;
    $validSourcePath = "";
    foreach($sourcePath in $sourcePaths)
    {
        if(Test-Path -Path $sourcePath)
        {
            $validSourcePath = $sourcePath;
            Write-Host "[OK]" -ForegroundColor Green;
        }
    }
    if($validSourcePath -eq "")
    {
        Write-Host "[KO]" -ForegroundColor Red;
    }
    
    Write-Host " > Clearing target package folder..." -NoNewline -ForegroundColor Gray;
    if(Test-Path -Path $targetPath)
    {
        Remove-Item $targetPath -Recurse -Force;
    }
    Write-Host "[OK]" -ForegroundColor Green;

    Write-Host " > Moving package to target $packageType folder..." -NoNewline -ForegroundColor Gray;
    Copy-Item $validSourcePath $targetPath -Recurse -Force;
    Write-Host "[OK]" -ForegroundColor Green;

    Write-Host " > Clearing source package folder..." -NoNewline -ForegroundColor Gray;
    if(Test-Path -Path $validSourcePath)
    {
        Remove-Item $validSourcePath -Recurse -Force;
    }
    Write-Host "[OK]" -ForegroundColor Green;
}