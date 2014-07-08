<#
.Synopsis
    Registers the custom object for overriding the .ToString() method for
    ComboBoxes to display the code signing certificates.
#>
function Register-PsISEScriptSigner() {

    ## Register the custom object time for use with the ComboBox
    try {
        Add-Type @"
        /// <summary>
        /// Custom object for storing Certificates and overriding the display method
        /// </summary>
        public class CertificateComboBoxItem {
            public string Text { get; set; }
            public System.Security.Cryptography.X509Certificates.X509Certificate2 Certificate { get; set; }

            public override string ToString() {
                string certificateIssuedName = Certificate.GetNameInfo(System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, false);
                string certificateIssuerName = Certificate.GetNameInfo(System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, true);
                return System.String.Format("{0} : {1}", certificateIssuedName, certificateIssuerName);
            }
        }
"@ | Out-Null;
    }
    catch {
        Write-Warning ($_.Exception.Message);
    }

    Register-PsISEScriptSignerMenus;
}

<#
.Synopsis
    Registers the PsISEScriptSigner custom menus
#>
function Register-PsISEScriptSignerMenus() {

	$PsISEScriptSignerRootMenu = Register-PsISEScriptSignerMenuRoot;
    if (!$PsISEScriptSignerRootMenu) { return; }

	Register-PsISEScriptSignerMenu -MenuName 'Check Script Signature' -ScriptBlock { Get-PsISEScriptSignature; } -HotKey "CTRL+SHIFT+C";
	Register-PsISEScriptSignerMenu -MenuName 'Save and Sign Script' -ScriptBlock { Save-PsISECurrentScript; Show-PsISEScriptSigner; } -HotKey "CTRL+SHIFT+S";
	Register-PsISEScriptSignerMenu -MenuName 'Sign Script' -ScriptBlock { Show-PsISEScriptSigner; } -HotKey "ALT+S";
    Register-PsISEScriptSignerMenu -MenuName 'Sign Script (Force Window)' -ScriptBlock { Show-PsISEScriptSigner -ForceWindow; } -HotKey "CTRL+ALT+S";
}

<#
.Synopsis
    Registers the PsISEScriptSigner root menu
#>
function Register-PsISEScriptSignerMenuRoot() {

    $PsISEScriptSignerMenuRoot = Find-PsISEScriptSignerMenuRoot;
    if ($PsISEScriptSignerMenuRoot) {
        $psISE.CurrentPowershellTab.AddOnsMenu.SubMenus.Remove($PsISEScriptSignerMenuRoot);
    }
	return $psISE.CurrentPowershellTab.AddOnsMenu.SubMenus.Add("PsISEScriptSigner", $null, $null);
}

<#
.Synopsis
    Registers a custom script block and hot key for a PsISEScriptSigner menu
#>
function Register-PsISEScriptSignerMenu($MenuName, $ScriptBlock, $HotKey) {

    $PsISEScriptSignerMenuRoot = Find-PsISEScriptSignerMenuRoot;
	$PsISEScriptSignerMenuRoot.SubMenus.Add($MenuName, $ScriptBlock, $HotKey);
}

<#
.Synopsis
    Locates the PsISEScriptSigner root menu for adding/removing custom menu options
#>
function Find-PsISEScriptSignerMenuRoot() {

    $PsISEScriptSignerMenuName = 'PsISEScriptSigner';
    $PsISEScriptSignerSubMenus = $psISE.CurrentPowershellTab.AddOnsMenu.SubMenus;
    return $PsISEScriptSignerSubMenus | where { $_.DisplayName -eq $PsISEScriptSignerMenuName; }
}

<#
.Synopsis
    Enumerates and adds personal and machine code signing certificates
    to the PsISEScriptSigner ComboBox.
#>
function Get-CodeSigningCertificates () {

    ## Add the user's code signing certificates to the ComboBox
    $UserCodeSigningCertificates = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert;
    if ($UserCodeSigningCertificates) {
        foreach ($UserCodeSigningCertificate in $UserCodeSigningCertificates) {
            $ComboBoxItem = New-Object CertificateComboBoxItem;
            $ComboBoxItem.Certificate = $UserCodeSigningCertificate;
            $CodeSigningCertificates.Items.Add($ComboBoxItem);
            $CodeSigningCertificates.SelectedItem = $ComboBoxItem;
        }
    }

    ## Add the computer's code signing certificates to the ComboBox
    $ComputerCodeSigningCertificates = Get-ChildItem -Path Cert:\LocalMachine\My -CodeSigningCert;
    if ($ComputerCodeSigningCertificates) {
        foreach ($ComputerCodeSigningCertificate in $ComputerCodeSigningCertificates) {
            $ComboBoxItem = New-Object CertificateComboBoxItem;
            $ComboBoxItem.Certificate = $ComputerCodeSigningCertificate;
            $CodeSigningCertificates.Items.Add($ComboBoxItem);
            $CodeSigningCertificates.SelectedItem = $ComboBoxItem;
        }
    }
}

<#
.Synopsis
    Saves the currently selected PowerShell ISE script.
#>
function Save-PsISECurrentScript () {

    ## How is the file encoded as Set-AuthenticodeSignature can't handle "Big-Endian"
    if ($psISE.CurrentFile.Encoding.EncodingName -match "Big-Endian") {
        ## Save the file in Unicode format
        $psise.CurrentFile.Save([Text.Encoding]::Unicode) | Out-Null;
    }
    else { $psISE.CurrentFile.Save(); }
}

<#
.Synopsis
    Signs a PowerShell (or any type of) file.
#>
function Add-PsISEScriptSignature($ScriptFilePath, $CodeSigningCertificate, $TimeStampServer) {

    try {
        ## Attempt to sign the script
        $SignResult = Set-AuthenticodeSignature -FilePath $ScriptFilePath -Certificate $CodeSigningCertificate -TimestampServer $TimeStampServer -errorAction Stop;

        ## Close and reopen the script
        Close-PsISEScriptFile;
        Open-PsISEScriptFile $ScriptFilePath;
    }
    catch {
        Write-Warning ("PsISEScriptSigner: Script signing failed. {0}" -f $_.Exception.Message);
        return $null;
    }

    ## Check the signature status
    if ($SignResult.Status -ne "Valid") { Write-Warning ("PsISEScriptSigner: Script signing failed. {0}" -f $SignResult.Status); }
    else { Write-Host "PsISEScriptSigner: '$ScriptFilePath' $($SignResult.StatusMessage.ToLower())"; }
}

<#
.Synopsis
    Gets the signature status of the currently selected PowerShell ISE script.
#>
function Get-PsISEScriptSignature() {

    try {
        $SignResult = Get-AuthenticodeSignature -FilePath $psISE.CurrentFile.FullPath -ErrorAction Stop;

        ## Check the signature status
        if ($SignResult.Status -ne "Valid") {
            Write-Warning ("PsISEScriptSigner: Script signature is invalid ({0})." -f $SignResult.Status);
        }
        else { 
            Write-Output "PsISEScriptSigner: '$($psISE.CurrentFile.FullPath)' $($SignResult.StatusMessage.ToLower())";
        }
    }
    catch {
        Write-Warning ("PsISEScriptSigner: Getting script signature failed. {0}" -f $_.Exception.Message);
    }
}

<#
.Synopsis
    Closes the currently selected PowerShell ISE script.
#>
function Close-PsISEScriptFile() {
    ## Close the script file
    $psISE.CurrentPowerShellTab.Files.Remove($psISE.CurrentFile) | Out-Null;
}

<#
.Synopsis
    Opens the requested script in the PowerShell ISE.
#>
function Open-PsISEScriptFile ($ScriptFilePath) {
    ## Add the script file
    $psISE.CurrentPowerShellTab.Files.Add($ScriptFilePath);
}

<#
.Synopsis
    Saves (if required) and signs the current PowerShell ISE script file.
#>
function OnCodeSignOkButtonClick() {

    ## Is the file untitled?
    if ($psISE.CurrentFile.IsUntitled) {
        [System.Windows.MessageBox]::Show("The script file has not been saved. Please save the file before attempting to sign it.", "Unsaved script", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Exclamation);
        if ($CodeSignForm -ne $null) { $CodeSignForm.Close(); }
        return;
    }
    ## Does the file need saving first?
    if (!$psISE.CurrentFile.IsSaved) {
        $SaveWarningDialogTitle = "Save: $($psISE.CurrentFile.DisplayName.TrimEnd("*"))";
        $SaveWarningDialog = [System.Windows.MessageBox]::ShowDialog("The script file must be saved before signing. Do you wish to save the file now?", $SaveWarningDialogTitle, [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Question);
        if ($SaveWarningDialog -eq "Cancel") {
            if ($CodeSignForm -ne $null) { $CodeSignForm.Close(); }
            return;
        }
        else { Save-PsISECurrentScript; }
    }

    # Save the filepath for the current script so it can be re-opened later
    $ScriptFilePath = $psISE.CurrentFile.FullPath;

    $SignResult = Add-PsISEScriptSignature $ScriptFilePath $CodeSigningCertificates.SelectedItem.Certificate "http://timestamp.verisign.com/scripts/timestamp.dll";

    if ($CodeSignForm -ne $null) { $CodeSignForm.Close(); }
}

<#
.Synopsis
    Displays the PsISEScriptSigner pop-up Window/ComboBox.
#>
function Show-PsISEScriptSigner([Switch]$ForceWindow) {

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing");
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");

    $CodeSignForm = New-Object System.Windows.Forms.Form;
    $CodeSignForm.Name = "PsISEScriptSigner";
    $CodeSignForm.Text = "Sign: '$($PsIse.CurrentFile.DisplayName.TrimEnd("*"))'";
    $CodeSignForm.Size = New-Object System.Drawing.Size(390, 150);
    $CodeSignForm.StartPosition = "CenterScreen";
    $CodeSignForm.MaximizeBox = $false;
    $CodeSignForm.MinimizeBox = $false;
    $CodeSignForm.SizeGripStyle = "Hide"
    $CodeSignForm.TopMost = $true;
    $CodeSignForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\System32\WindowsPowershell\v1.0\Powershell.exe");

    $CodeSignLabel = New-Object System.Windows.Forms.Label;
    $CodeSignLabel.Location = New-Object System.Drawing.Size(15, 15);
    $CodeSignLabel.Size = New-Object System.Drawing.Size(262.5, 20);
    $CodeSignLabel.Text = "Please select code-signing certificate:";
    $CodeSignForm.Controls.Add($CodeSignLabel);

    $CodeSigningCertificates = New-Object System.Windows.Forms.ComboBox;
    $CodeSigningCertificates.Location = New-Object System.Drawing.Size(15, 40);
    $CodeSigningCertificates.Size = New-Object System.Drawing.Size(345, 50);
    $CodeSigningCertificates.Sorted = $true;
    $CodeSignForm.Controls.Add($CodeSigningCertificates);

    $CodeSignOkButton = New-Object System.Windows.Forms.Button;
    $CodeSignOkButton.Location = New-Object System.Drawing.Size(205, 70);
    $CodeSignOkButton.Size = New-Object System.Drawing.Size(75, 25);
    $CodeSignOkButton.Text = "&OK";
    $CodeSignOkButton.Add_Click({ OnCodeSignOkButtonClick; });
    $CodeSignForm.Controls.Add($CodeSignOkButton);
    $CodeSignForm.AcceptButton = $CodeSignOkButton;

    $CodeSignCancelButton = New-Object System.Windows.Forms.Button;
    $CodeSignCancelButton.Location = New-Object System.Drawing.Size(285, 70);
    $CodeSignCancelButton.Size = New-Object System.Drawing.Size(75, 25);
    $CodeSignCancelButton.Text = "&Cancel";
    $CodeSignCancelButton.Add_Click({ $CodeSignForm.Close(); });
    $CodeSignForm.Controls.Add($CodeSignCancelButton);
    $CodeSignForm.CancelButton = $CodeSignCancelButton;

    ## Load the ComboBox
    Get-CodeSigningCertificates | Out-Null;

    ## Check whether we have any certificates
    if ($CodeSigningCertificates.Items.Count -eq 0) { [void] [System.Windows.Forms.MessageBox]::Show("No code signing certificates found.", "Oops", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); }

    ## If we only have 1 certificate, no need to select it/click OK, so invoke the OnClick method
    elseif (($CodeSigningCertificates.Items.Count -eq 1) -and (!$ForceWindow)) { OnCodeSignOkButtonClick; }

    ## Multiple certificates (or been forced) so show the selection form
    else {
        $CodeSignForm.Add_Shown({$CodeSignForm.Activate()})
        [void]$CodeSignForm.ShowDialog();
    }

}

## Register PsISEScriptSigner!
Register-PsISEScriptSigner;
