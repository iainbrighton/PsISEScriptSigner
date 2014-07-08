$currentDir = Split-Path -parent $MyInvocation.MyCommand.Path;
$userModuleDir = "$([System.Environment]::GetFolderPath("MyDocuments"))\WindowsPowerShell\Modules\PsISEScriptSigner";
$iseProfileFile = "$([System.Environment]::GetFolderPath("MyDocuments"))\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1";

$copyToModules = Read-Host 'Install PsISEScriptSigner to your Modules directory [y/n]?';
if ($copyToModules -ieq 'y') {
	if (Test-Path $userModuleDir) {
		Write-Host "Removing directory '$userModuleDir'..." -NoNewline;
		Remove-Item -Path $userModuleDir -Force	-Recurse;
        Write-Host "OK";
	}
	
	Write-Host "Unblocking PSISEScriptSigner files..." -NoNewLine;
	Get-ChildItem (Join-Path $currentDir "PsISEScriptSigner") | Unblock-File;
	Write-Host "OK";

	Write-Host "Copying PSISEScriptSigner files to '$userModuleDir'..." -NoNewline;
	Copy-Item -Path (Join-Path $currentDir "PsISEScriptSigner") -Destination $userModuleDir -Recurse -Force;
    Write-Host "OK";
}

Write-Host "";

$installToProfile = Read-Host 'Install PsISEScriptSigner to ISE Profile (will start when ISE starts) [y/n]?';

if ($installToProfile -ieq 'y') {
	if (!(Test-Path $iseProfileFile)) {
		Write-Host "Creating file '$iseProfileFile'..." -NoNewline;
		New-Item -Path $iseProfileFile -ItemType file | Out-Null;
        Write-Host "OK";
		$contents = "";
	} else {
		Write-Host "Reading file '$iseProfileFile'..." -NoNewLine;
		$contents = Get-Content -Path $iseProfileFile | Out-String;
        Write-Host "OK";
	}

	$importModule = "Import-Module PsISEScriptSigner";

	if ($contents -inotmatch $importModule) {
		Write-Host "Adding '$importModule'..." -NoNewLine;
		Add-Content -Path $iseProfileFile -Value $importModule | Out-Null;
        Write-Host "OK";
	} else {
		Write-Host "Import command for PsISEScriptSigner already exists in Powershell ISE profile.";
	}
}
