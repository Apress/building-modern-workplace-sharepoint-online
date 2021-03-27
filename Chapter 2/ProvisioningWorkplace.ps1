Install-Module SharePointPnPPowerShellOnline
Connect-PnPOnline -Url https://cloudhadi.sharepoint.com/sites/Workplace -UseWebLogin
Enable-PnPFeature -Identity 3bae86a2-776d-499d-9db8-fa4cdc7884f8 -Scope Site -ErrorAction Stop 
Apply-PnPProvisioningTemplate -Path .\ProvisioningWorkplace.xml