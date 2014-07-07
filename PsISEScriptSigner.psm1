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

# SIG # Begin signature block
# MIIaogYJKoZIhvcNAQcCoIIakzCCGo8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUc+Jqn7phtofCNH7U6OHESQu6
# iDGgghXYMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggaUMIIFfKADAgECAhAG8BXYFUYj6XmzRgEaZJSVMA0GCSqGSIb3DQEBBQUAMG8x
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xLjAsBgNVBAMTJURpZ2lDZXJ0IEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBLTEwHhcNMTMwNDE3MDAwMDAwWhcNMTUwNzE2MTIwMDAwWjBg
# MQswCQYDVQQGEwJHQjEPMA0GA1UEBxMGT3hmb3JkMR8wHQYDVQQKExZWaXJ0dWFs
# IEVuZ2luZSBMaW1pdGVkMR8wHQYDVQQDExZWaXJ0dWFsIEVuZ2luZSBMaW1pdGVk
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1dxm3r1cUKp7rYZBDAeo
# Lm0iLIgYuzeg7tC2mt7kEJfvGiSVx4/d3pYw2/GpDB08JjsoAYIfhWOuGtUf0RRy
# 5QcyrfWDCmLfUApf83/GJZrATWs1OPzdYEsLzVrx7ZtvcCVvlEIyG4RJmhSG2mZS
# 6P0D68a2/U4QmcNEGpnbTyszHds8BnVL1D3oQP+rcXN2jDP83/rECmGgYGexvRkV
# K/+HHrporgkT4KRMbrWXMRPrLQazIFeg1mnm1UtjxTXN7IPaY97qwxhxPqwpL3DH
# PdF/6+rC1ZQZ27akf5qporAlsftUe3URHFmmJ8NrLivANrwco9BY3If4iAvz9ipl
# mQIDAQABo4IDOTCCAzUwHwYDVR0jBBgwFoAUe2jOKarAF75JeuHlP9an90WPNTIw
# HQYDVR0OBBYEFNQ3nxxDKFobighYZExYqzXq8SQTMA4GA1UdDwEB/wQEAwIHgDAT
# BgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHR8EbDBqMDOgMaAvhi1odHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vYXNzdXJlZC1jcy0yMDExYS5jcmwwM6AxoC+GLWh0dHA6
# Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9hc3N1cmVkLWNzLTIwMTFhLmNybDCCAcQGA1Ud
# IASCAbswggG3MIIBswYJYIZIAYb9bAMBMIIBpDA6BggrBgEFBQcCARYuaHR0cDov
# L3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsG
# AQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABD
# AGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABh
# AGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQBy
# AHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBn
# ACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABs
# AGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABp
# AG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBl
# AGYAZQByAGUAbgBjAGUALjCBggYIKwYBBQUHAQEEdjB0MCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTAYIKwYBBQUHMAKGQGh0dHA6Ly9jYWNl
# cnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENvZGVTaWduaW5nQ0Et
# MS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQUFAAOCAQEAPsyUuxkYkEGE
# 1bl4g3Muy5QxQq8frp34BPf+Sm6E9J915eBizW72ofbm08O9NkQvszbT4GTZaO/o
# SExSDbLIxHI98zi7AavVPuRpmVnfoF55yVomh/BYAU8vu0M7FvUeIhSAUfz0Q8PK
# wT5U+SdNoE7+xgxd4zHjBA3kUo3TZ+R/+MDd2Hzv6vrgxUfGeQfBCwafdEjD4pHr
# 0kvXcPq6VnQpsv92P3wvgsCrsTKIgtaNIfkGe5eCcTQ7pYTBauZl+XmyFvyiADKo
# 6Dng4jyuxYRP3EdCGVlZK7sEmiz1Y2f3zh0xoF58B3xXDnRJxo8ArlEAG8KzXn6w
# ryaA1vbgITCCBqMwggWLoAMCAQICEA+oSQYV1wCgviF2/cXsbb0wDQYJKoZIhvcN
# AQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJl
# ZCBJRCBSb290IENBMB4XDTExMDIxMTEyMDAwMFoXDTI2MDIxMDEyMDAwMFowbzEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNzdXJlZCBJRCBDb2Rl
# IFNpZ25pbmcgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJx8
# +aCPCsqJS1OaPOwZIn8My/dIRNA/Im6aT/rO38bTJJH/qFKT53L48UaGlMWrF/R4
# f8t6vpAmHHxTL+WD57tqBSjMoBcRSxgg87e98tzLuIZARR9P+TmY0zvrb2mkXAEu
# sWbpprjcBt6ujWL+RCeCqQPD/uYmC5NJceU4bU7+gFxnd7XVb2ZklGu7iElo2NH0
# fiHB5sUeyeCWuAmV+UuerswxvWpaQqfEBUd9YCvZoV29+1aT7xv8cvnfPjL93Sos
# MkbaXmO80LjLTBA1/FBfrENEfP6ERFC0jCo9dAz0eotyS+BWtRO2Y+k/Tkkj5wYW
# 8CWrAfgoQebH1GQ7XasCAwEAAaOCA0MwggM/MA4GA1UdDwEB/wQEAwIBhjATBgNV
# HSUEDDAKBggrBgEFBQcDAzCCAcMGA1UdIASCAbowggG2MIIBsgYIYIZIAYb9bAMw
# ggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3Bz
# LXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUA
# cwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMA
# bwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYA
# IAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQA
# IAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUA
# bQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkA
# dAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAA
# aABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMBIGA1UdEwEB
# /wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8v
# b2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6
# MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYEFHtozimqwBe+SXrh
# 5T/Wp/dFjzUyMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqG
# SIb3DQEBBQUAA4IBAQB7ch1k/4jIOsG36eepxIe725SS15BZM/orh96oW4AlPxOP
# m4MbfEPE5ozfOT7DFeyw2jshJXskwXJduEeRgRNG+pw/alE43rQly/Cr38UoAVR5
# EEYk0TgPJqFhkE26vSjmP/HEqpv22jVTT8nyPdNs3CPtqqBNZwnzOoA9PPs2TJDn
# dqTd8jq/VjUvokxl6ODU2tHHyJFqLSNPNzsZlBjU1ZwQPNWxHBn/j8hrm574rpyZ
# lnjRzZxRFVtCJnJajQpKI5JA6IbeIsKTOtSbaKbfKX8GuTwOvZ/EhpyCR0JxMoYJ
# mXIJeUudcWn1Qf9/OXdk8YSNvosesn1oo6WQsQz/MYIENDCCBDACAQEwgYMwbzEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNzdXJlZCBJRCBDb2Rl
# IFNpZ25pbmcgQ0EtMQIQBvAV2BVGI+l5s0YBGmSUlTAJBgUrDgMCGgUAoHgwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
# tnsYUAaWX0VCS3Nar+b4Xb4LDgEwDQYJKoZIhvcNAQEBBQAEggEAoMhGfPeDO3pg
# cd+PSAzEdso9sXpyexbfBwlb3R96h6dPrDp6B1a+6L0gX8v4UDhlwu8C+/f0pxMs
# Z0iwfwKmDfN9EV3sxWfV2d+X9of+t40I20P7blpqg+6SRW/EhK0+7WIBrWEZZ8WK
# nGbW2Rq4rdaqRSUSDgNBZfrdTDJgOQCf1cyNW4Dfpm8iRf/zEwjyA4a7Sl/22vvb
# VnzIlgroD4j+HASeQ3Qhb5A/9iT0Ku3cM1e2sPMD8ihrvD0wUX+fG3xZUd6QP40M
# 8xWJsjc0W3vhXAN2cOgTiJQKQ8YFH0eC/8PMdz+SLUdnTqBKI7r8/b40pom2IeVE
# bWihzSRqrKGCAgswggIHBgkqhkiG9w0BCQYxggH4MIIB9AIBATByMF4xCzAJBgNV
# BAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjEwMC4GA1UEAxMn
# U3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQSAtIEcyAhAOz/Q4yP6/
# NW4E2GqYGxpQMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0xNDA2MjQyMDU3MjVaMCMGCSqGSIb3DQEJBDEWBBS8
# sYa5G71xBi5k59sb7IgzEVC/7jANBgkqhkiG9w0BAQEFAASCAQAoT1DoakHmHdyF
# eA5PmlsSjKuM49T+JekxRGVXL5axANQVw4Rik6Bdaohy4B5oLc8f8680WPRG8JHm
# UyB9ho3d2oDKMJqAn4j71S00TaTGYel3GbEB3tL+NR/cxeJJeNTDfVwLAXo5vxbL
# C0hzzJMM7HR1HYR0us7jC1l5TH1tgHXkLV8vonRRdhUkp6SfI78TrQATHi1KAHVU
# Wyv0+IYKCOn1QxPbu1i9D5W8+2rgRc7LOz+TCDWbM/cSLGvQjAWMgB15Y6ELKDgs
# IXjSyNaz4rVKs0fhSwt/GPI/sEw01qNtHLDzANF7SKpAFUUdY6D8+hBWwyz4jq+I
# fmKNX7ff
# SIG # End signature block
