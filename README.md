## Powershell ISE Addon - Script Signer

<a href="https://github.com/iainbrighton/PsISEScriptSigner/releases/latest">Download</a>

##### Description

Adds script signing, certification selection and signature validation capabilities to the PowerShell ISE. 

* Check script signature (CTRL+SHIFT+C) - Validates the signature of the current file in the ISE.
* Save and sign script (CTRL+SHIFT+S) - Saves and then signs the current file in the ISE.
* Sign script (ALT+S) - Signs the current file within the ISE.
* Sign script (Force Window) (CTRL+ALT+S) - Forces the signature selection window to appear.

Requires Powershell 3.0 or above.

If you find it useful, unearth any bugs or have any suggestions for improvements, feel free to report an <a href="https://github.com/iainbrighton/PsISEScriptSigner/issues">issue</a> or place a comment on the project home page.

##### Screenshots
![ScreenShot](./PsISEScriptSignerMenus.png?raw=true)

![ScreenShot](./PsISEScriptSignerCertificateSelection.png?raw=true)

##### Installation

* Automatic:
 * Run 'Install_to_User_Modules.bat'.
* Manual:
 * Ensure all the files are unblocked (properties of the file / General).
 * Copy PSISEScriptSigner to $env:USERPROFILE\Documents\WindowsPowerShell\Modules.
 * Launch the PowerShell ISE.
 * Run 'Import-Module PsISEScriptSigner'.
 * If you want it to be loaded automatically when ISE starts, add the line above to your ISE profile (see $profile).

##### Usage

When you have a file open in the PowerShell ISE that you wish to sign, select the 'Add-ons > PsISEScriptSigner' menu and choose either the 'Save and sign script' or 'Sign script' options (or utilise the keyboard shortcuts). The file will be reloaded and script's signature status will be display in the ISE's console window.

If you wish to check the validity of a script's signature, open the file in the PowerShell ISE and select the 'Add-ons > PsISEScriptSigner > Check script signature' option. The script's signature status will be displayed in the ISE's console window.

##### Why?

Firstly - because I always forget how to sign scripts!

Secondly - if you only have a single code signing certificate you won't be prompted for a certificate selection speeding up the process. If you have multiple code-signing certificates installed (either in the user or machine store), you will be prompted to select the required certificate to sign the file. This is incredibly handy if you need to sign scripts for different entities.

##### Implementation Details

Written in PowerShell! To modify keyboard shortcuts, edit PsISEScriptSigner.psm1 file and either remove the signature or resign the file yourself.