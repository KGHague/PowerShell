<#
.SYNOPSIS
Office365 MFA Report

.DESCRIPTION
This script is used to report on all users MFA status in Office365. The following properties are exported 
DisplayName, UPN, AssignedLicence, Licensed, DefaultMethod, MFA Enabled

.Module Dependencies
Install-Module -Name MSOnline
Install-Module -Name importexcel

.EXAMPLE
.\Office365_MFA_Report.ps1 -ExportPath C:\Temp\

#>
param(
    [parameter(Mandatory)]
    [String]$ExportPath 
    ) 


## Module Dependencies
#Install-Module -Name MSOnline
#Install-Module -Name importexcel


## Import user list
$users = Get-MsolUser -All

## Set result array
$results = @()
foreach ($user in $Users){

    ## Get user properties
    Write-Host "Checking $($user.UserPrincipalName) For strong authentication" -ForegroundColor Green
    if($user.StrongAuthenticationMethods){
        Write-Host "Authentication Method found for  $($user.UserPrincipalName)" -ForegroundColor Yellow

        ## Create report hash table
        $props = @{
            FirstName = $user.FirstName
            LastName = $user.LastName
            UPN = $user.UserPrincipalName
            "MFA Enabled" = "True"
            DefaultMethod = $user.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq "True"} | Select-Object MethodType -ExpandProperty MethodType
            Licensed  = $user.IsLicensed
            AssignedLicense = if($user.Licenses.AccountSkuId){$user.Licenses.AccountSkuId -join ","} else {"No License Assigned"}
        }

        ## Create result object
        $results += New-Object PSObject -Property $props
    }
    else {
        Write-Host "No Authentication Method found on $($user.UserPrincipalName)" -ForegroundColor red

        ## Create report hash table
        $props = @{
            FirstName = $user.FirstName
            LastName = $user.LastName
            UPN = $user.UserPrincipalName
            "MFA Enabled" = "False"
            DefaultMethod = "N/A"
            Licensed  = $user.IsLicensed
            AssignedLicense = if($user.Licenses.AccountSkuId){$user.Licenses.AccountSkuId -join ","} else {"No License Assigned"}
        }

        ## Create result object
        $results += New-Object PSObject -Property $props
    }
}

## Export results
#$results | Export-Csv "$ExportPath\$((Get-MsolCompanyInformation).DisplayName) - $(Get-Date -Format yyyyMMdd-HHmmss).csv" -NoTypeInformation
$Text1 = New-ConditionalText Never
$style = New-ExcelStyle -FontSize 16 -Bold -Range "A1:G1" -HorizontalAlignment Center -Merge
$CompanyName = (Get-MsolCompanyInformation).DisplayName
$results | sort LastName,FirstName | select FirstName,LastName,UPN,"MFA Enabled",DefaultMethod,Licensed,AssignedLicense | Export-Excel -Path "$ExportPath\$CompanyName - $(Get-Date -Format yyyyMMdd-HHmmss).xlsx" -AutoSize -ConditionalText $Text1 -Title "$CompanyName" -TableStyle Medium9 -Style $style