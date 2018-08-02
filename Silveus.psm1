### Created by Daniel Hoyt 3/19/18 dihoyt@outlook.com ###
import-module activedirectory
Write-Host "Most of these cmdlets must be run in a domain administrator shell." -ForegroundColor Yellow
Write-Host "To view a full list of cmdlets run Get-Command -Module Silveus" -ForegroundColor Yellow




####---- Windows Explorer CSV File Select ----####
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    
}



####---- Windows Explorer CSV Save Location ----####
Function Export-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "CSV (*.csv)| *.csv"
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.filename
    
}



####---- Import CSV Yes/No 
####---- Outputs to $importcsv ----####
Function Get-PromptForCSV()
{
Write-host "Would you like to import a csv file for this operation?(Default is No)" -ForegroundColor Yellow 
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes, Choose a CSV File"; $true} 
       N {Write-Host "No, I will input manually"; $false} 
       Default {Write-Host "Default, input manually."; $false} 
     }
} 
####---- End of PromptFor-CSV ----####



 
####----Output the user's Distinguished Name ----####
Function Get-DistinguishedName ($strUserName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$strUserName))" 
   $result = $searcher.FindOne() 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
####---- End of Get-DistinguishedName ----#### 




####---- Create User ----####
Function Add-User ()
{
    $importcsv = Get-PromptForCSV 
    If ($importcsv -eq $false) {
    $first = Read-Host "First Name"
    $last = Read-Host "Last Name"
    $job = Read-Host "Job Title"
    $dep = Read-Host "Department"
    $company = Read-Host "Company"
    $emaildomain = Read-Host "Sign in Domain (ex: cropins.net)"
    $password = Read-Host -AsSecureString "Password"
    $manager = Read-Host "Manager (first.last)"
    $OU = 'OU=Warsaw,OU=SIG,OU=Cedar Holdings,OU=Users,OU=MyBusiness,DC=SilveusInsurance,DC=local'
    $fullname = $first + ' ' + $last
    $samacc = $first + '.' + $last
    $upn = $samacc + '@' + $emaildomain
    $replyto = $samacc + '@' + $emaildomain


    Write-Host "A user will be created with the following details."
    Write-Host 'Name:'$First $Last
    Write-Host 'Email/Sign-On Address:' $upn
    Write-Host 'Organization Info:'$Job',' $dep',' $company 
    Write-Host 'Manager:'$manager
    Write-Host ''
    Write-Host "If this info looks innacurate press CTRL+C to quit. Otherwise,"

    New-ADUser -SamAccountName $samacc -UserPrincipalName $upn -AccountPassword $password -Name $fullname `
    -GivenName $first -Surname $last -DisplayName $fullname -Path $OU -Manager $manager -Company $company `
    -Title $job -Department $dep -Enabled $true -Email $replyto
    Write-Host 'User has been created in the following directory:'$OU
} ElseIf ($importcsv -eq $true) {
    $inputfile = Get-FileName "C:\Users\$([Environment]::UserName)\Desktop"
    $inputdata = Import-Csv $inputfile
    $password = Read-Host -AsSecureString "Password for accounts."
    foreach ($line in $inputdata) {
        $first = $inputdata.First
        $last = $inputdata.Last
        $job = $inputdata.Title
        $dep = $inputdata.Department
        $company = $inputdata.Company
        $emaildomain = $inputdata.Domain
        $manager = $inputdata.Manager
        $OU = 'OU=Warsaw,OU=SIG,OU=Cedar Holdings,OU=Users,OU=MyBusiness,DC=SilveusInsurance,DC=local'
        $fullname = $first + ' ' + $last
        $samacc = $first + '.' + $last
        $upn = $samacc + '@' + $emaildomain
        $replyto = $samacc + '@' + $emaildomain

        New-ADUser -SamAccountName $samacc -UserPrincipalName $upn -AccountPassword $password -Name $fullname `
        -GivenName $first -Surname $last -DisplayName $fullname -Path $OU -Manager $manager -Company $company `
        -Title $job -Department $dep -Enabled $true -Email $replyto
        Write-Host 'Users have been added in the following directory:'$OU
    }
    }
}
####---- End Create User ----####





####---- Remove Users ----####
Function Remove-User($user)
{
$importcsv = Get-PromptForCSV
if ($importcsv -eq $true) {
    $inputfile = Get-FileName "C:\Users\$([Environment]::UserName)\Desktop"
    $inputdata = import-csv $inputfile
    import-module activedirectory
    foreach ($line in $inputdata) {
    $identity = $inputdata.First + "." + $inputdata.Last
       Remove-ADUser -Identity $identity
    Write-Host "If you receive errors that this command can only be run on leaf objects please run 'Remove-UserWChildren' on the affected users."
}

} Else {
if ($user = $null)
{
    $user = Read-Host "Enter Username (Example: daniel.hoyt)"
    } Else {
    Remove-ADUser -Identity $user
    Write-Host "If you receive an error that this command can only be run on leaf objects please run 'Remove-UserWChildren' on the affected user." 
    $user = $null
    }             
}
}

####---- End Remove Users ----####





Function Suspend-User ($strUserName) 
{
    $importcsv = Get-PromptForCSV
    if ($importcsv -eq $false) {
        if ($strUserName -eq $null) {
            $strUserName = Read-Host 'Enter User to Disable'
        } Else {
            $strDN = Get-DistinguishedName $strUserName
            $strReplyTo = get-aduser $strUserName -pr proxyaddresses |select -ExpandProperty proxyaddresses |? {$_ -cmatch '^SMTP'}
            $OU = 'OU=Disabled Users,OU=Users,OU=MyBusiness,DC=SilveusInsurance,DC=local'

            $strManager = Get-ADUser $strUserName -pr manager | select -ExpandProperty manager
            #Set-ADUser -Identity $strManager -add @{proxyaddresses=$strReplyTo.ToLower()} ##Currently Errors out
            Set-ADUser -Identity $strUserName -Manager $null -Clear ProxyAddresses 
            Disable-ADAccount -Identity $strUserName
            Set-ADObject -Identity $strDN -Replace @{msExchHideFromAddressLists = $true}
            Move-ADObject -Identity $strDN -TargetPath $OU

            Get-ADUser -Identity $strUserName -Properties MemberOf | ForEach-Object {
              $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false
              $strUserName = $null }
          }
            Write-Host "The User has been disabled and removed from all groups."
            Write-Host "The Password has been set to 'd1sabled'."
            Write-Host "Primary reply-to address assigned to manager."
            Write-Host "Please assign reply-to address to manager."
    } Else {
        $inputfile = Get-FileName "C:\Users\$([Environment]::UserName)\Desktop"
        $inputdata = Import-Csv $inputfile
        foreach ($line in $inputdata) {
            $strUserName = $inputdata.Username
            $strDN = Get-DistinguishedName $strUserName
            $strReplyTo = get-aduser $strUserName -pr proxyaddresses |select -ExpandProperty proxyaddresses |? {$_ -cmatch '^SMTP'}
            $OU = 'OU=Disabled Users,OU=Users,OU=MyBusiness,DC=SilveusInsurance,DC=local'

            $strManager = Get-ADUser $strUserName -pr manager | select -ExpandProperty manager
            ##Set-ADUser -Identity $strManager -add @{proxyaddresses=$strReplyTo.ToLower()} ##Currently Errors out
            Set-ADUser -Identity $strUserName -Manager $null -Clear ProxyAddresses 
            Disable-ADAccount -Identity $strUserName
            Set-ADObject -Identity $strDN -Replace @{msExchHideFromAddressLists = $true}
            Move-ADObject -Identity $strDN -TargetPath $OU

            Get-ADUser -Identity $strUserName -Properties MemberOf | ForEach-Object {
                $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false
            Write-Host "User" $strUserName "has been disabled."
            }
        }
    }
}

####---- Remove-UserwChildren ----####
Function Remove-UserWChildren ($strUserName)
{
    $strUserName = Read-Host "Username"
    $strDN = Get-DistinguishedName $strUserName 
    Remove-ADObject -Identity $strDN -Recursive   
}

Function Get-PwdLastSet ($days){
if ($days -eq $null) 
    {
        $days = Read-Host "How many days ago?"
} Else {
    $outputfile = Export-FileName
    Get-ADUser -Filter "pwdLastSet -lt $((Get-Date).AddDays(-$days).ToFileTimeUTC()) -and pwdLastSet -ne 0" | Export-csv $outputfile
    $days = $null
    }
}

Function Get-LastLogon ($days)
{
    if ($days -eq $null) {
        $days = Read-Host "How many days ago?"
    } Else {
        $outputfile = Export-FileName
        Search-ADAccount -UsersOnly -AccountInactive -TimeSpan $days | Export-Csv $outputfile
        $days = $null
    }
}

