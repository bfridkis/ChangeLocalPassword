Write-Host "`n"
Import-Module ActiveDirectory

#Clear any previous data
$accountToUpdate = $adminPassword = $confirmAdminPassword = $null

#Query AD for a list of all domain computers
Get-ADObject -LDAPFilter "(objectClass=computer)" | 
Where-Object { $_.Name -notlike "PCNVS*" -and $_.Name -notlike "DEVVS*" -and $_.Name -notlike "PCNVC*" } | 
Select-Object -ExpandProperty Name | sort-object name | Set-Variable -Name Computers

#To set computers manually, comment out the previous three lines, uncomment the line below and add comma separated machine names after the '=' character.
#$computers = "SEL-PCN-BLNDSER"

#Get username and password to update
$accountToUpdate = Read-Host "Enter Username For Password Update"

Do {
    $adminPassword = Read-Host "Enter New Password" -AsSecureString
    If((New-Object PSCredential '.', $adminPassword).GetNetworkCredential().Password -eq 'q') { exit }
    $confirmAdminPassword = Read-Host "Confirm New Password" -AsSecureString
    If((New-Object PSCredential '.', $confirmAdminPassword).GetNetworkCredential().Password -eq 'q') { exit }
    If((New-Object PSCredential '.', $adminPassword).GetNetworkCredential().Password -ne `
        (New-Object PSCredential '.', $confirmAdminPassword).GetNetworkCredential().Password) {  
        Write-Host "Passwords Entered Do Not Match. Please Try Again or Enter 'q' to quit." 
    }
} While ($adminPassword -eq $null -or (New-Object PSCredential '.', $adminPassword).GetNetworkCredential().Password `
                                  -ne (New-Object PSCredential '.', $confirmAdminPassword).GetNetworkCredential().Password)

#Initialize lists for results and errors output
$results = New-Object System.Collections.Generic.List[System.Object]

Write-Host "`nRunning...Please Wait..."

ForEach($computer in $computers) {  
    Try {
        Test-Connection $computer -Count 2 -ErrorAction Stop
        $adminUser = [ADSI]("WinNT://$computer/$accountToUpdate,user")  
        #$adminUser.psbase.invoke("setpassword",$adminPassword)
        $adminUser.SetPassword((New-Object PSCredential '.', $adminPassword).GetNetworkCredential().Password)
        $results.Add([PSCustomObject]@{'Hostname'=$computer ; 'Result' = "Password Updated Successfully"})
    }
    Catch { $results.Add([PSCustomObject]@{'Hostname'=$computer ; 'Result' = $_.Exception.Message}) }
}  

$results | Export-CSV ".\LocalUser_$($accountToUpdate)_PasswordUpdateResults-$(Get-Date -Format MMddyyyy_HHmmss).csv" -NoTypeInformation
$results | Format-Table -AutoSize

$adminPassword = $confirmAdminPassword = $null

# References
# https://stackoverflow.com/questions/38901752/verify-passwords-match-in-windows-powershell
# https://serverfault.com/questions/929362/resetting-local-admin-password-for-a-remote-computer-using-powershell
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/convertfrom-securestring?view=powershell-7.2
# https://stackoverflow.com/questions/28352141/convert-a-secure-string-to-plain-text
