

# Location of directory with Excel sheets containing new user details - references current directory of this script
$dir_with_excel = $PSScriptRoot+'\user_accounts_excel'

# Path to OU in which you wish to search for users.
$All_ous = 'OU=London,DC=loncc,DC=local'

# Lists all departments ( OUs) in London OU of loncc.local Domain
function list_of_dep {
    $list_of_ous = Get-ADOrganizationalUnit -Filter * -SearchBase $All_ous -SearchScope OneLevel | Select-Object Name, DistinguishedName
    Write-Host "Please choose a Department ( OU ) to search in:`n"
        For ($i=0; $i -lt $list_of_ous.Count; $i++)  {
          Write-Host "$($i+1): $($list_of_ous[$i].Name)"
    }

    [int]$number = Read-Host "`nPress the number to select a Department "

    Write-Host "`nYou've selected =  " -NoNewline; Write-Host $($list_of_ous.Name[$number-1])"`n"-ForegroundColor Green
    $Global:dep_name = $($list_of_ous.Name[$number-1])
    $Global:dep = $list_of_ous.DistinguishedName[$number-1]
    get_user_to_copy
}

# Lists all users in chosed OU - prints their Name and Job Position. 
function get_user_to_copy {
    $list_of_users_in_ou = Get-ADUser -Properties Description, SamAccountName -Filter * -SearchBase $dep | Select-Object Name, Description, SamAccountName
    Write-Host "Please choose a User - his/hers properties will be copied to new user`n"
    #$Results = '' | SELECT UserName,JobPosition

        $table = For ($i=0; $i -lt $list_of_users_in_ou.Count; $i++)  {
                                 $($Results = '' | SELECT Num,UserName,JobPosition
                                $Results.Num = $i+1
                                $Results.UserName = $list_of_users_in_ou[$i].Name
                                 $Results.JobPosition = $list_of_users_in_ou[$i].Description)
                                 $Results
                                 
    }
    $table | Format-Table -AutoSize 
    [int]$number = Read-Host "Press the number to select a user: "

    Write-Host "`nYou've selected ; " -NoNewline; Write-Host $($list_of_users_in_ou.Name[$number-1])"`n" -ForegroundColor Green
    $Global:user_to_cp = $list_of_users_in_ou.SamAccountName[$number-1]
    copy_ad_user
}



#function get_user_to_copy {
   # $list_of_users_in_ou = Get-ADUser -Properties Description, SamAccountName -Filter * -SearchBase $dep | Select-Object Name, Description, SamAccountName
   # Write-Host "Please choose a User:`n"
    #    For ($i=0; $i -lt $list_of_users_in_ou.Count; $i++)  {
        #  Write-Host "$($i+1): $($list_of_users_in_ou[$i].Name) -Job Title:  $($list_of_users_in_ou[$i].Description)"
   # }

   # [int]$number = Read-Host "Press the number to select a user: "

   # Write-Host "`nYou've selected ; " -NoNewline; Write-Host $($list_of_users_in_ou.Name[$number-1])"`n" -ForegroundColor Green
   # $Global:user_to_cp = $list_of_users_in_ou.SamAccountName[$number-1]
#}

#Welcome screen
function welcome{
Write-Host "`n***************************************"
Write-Host "*****                             *****"
Write-Host "*****    Create AD user script    *****"
Write-Host "*****                             *****"
Write-Host "***************************************`n`n"
List_of_excel_files
}

# Lists all Excel files in specified directory - lists all file names.
function List_of_excel_files {
    $list_of_files = Get-ChildItem -Path $dir_with_excel -Filter "*" | where {$_.extension -eq ".xlsx"} | Select-Object Name | Sort-Object CreationTime -Descending
    Write-Host "`nPlease choose Excel file to copy data from:`n"
        For ($i=0; $i -lt $list_of_files.Count; $i++)  {
          Write-Host "$($i+1): $($list_of_files[$i].Name)"
    }

    [int]$number = Read-Host "`nPress the number to select Excel file "

    Write-Host "`nYou've selected = " -NoNewline; Write-Host $($list_of_files.Name[$number-1])"`n" -ForegroundColor Green
    $Global:excel_file_to_use = $dir_with_excel+'\'+$list_of_files.Name[$number-1]
    get_data_from_excel
}

# Reads chosen Excel file and copies into Global variables persons First Name, Lst Name, Email address 
# and Password - this is a random AD password which user will have to type upon first login. He will be asked to change it. 
# You can change belof Cell Values to suit your Excel file
function get_data_from_excel {
    Write-Host "Reading data from selected Excel file...." 
    $excel = Open-ExcelPackage -Path $excel_file_to_use
    $worksheet = $excel.Workbook.Worksheets['Sheet1']
    $Global:AD_User_pass = $worksheet.Cells['E11'].Value
    $Global:AD_User_first_name = $worksheet.Cells['E5'].Value
    $Global:AD_User_last_name = $worksheet.Cells['F5'].Value
    $Global:AD_User_email = $worksheet.Cells['E9'].Value
    Write-Host "Reading data complete.`n"
    list_of_dep
}

# Creates new user with properties of the selected user. It copies following info:
#City, Postal code, street address, phone number , description ( job titile). Also copies All Group Membership.

function copy_ad_user {

# must specify the properties to copy over
$u = Get-ADUser -Identity $user_to_cp -Properties city,postalcode,"streetaddress", officephone, description
$ou = (Get-AdUser $user_to_cp).distinguishedName.Split(',',2)[1]

$cn = "$AD_User_first_name $AD_User_last_name" # "Firstname LastName"
$sam = $AD_User_first_name.ToLower()+"."+$AD_User_last_name.ToLower() # "firstname.lastName" lowercase
#  Creating new user based on $u original user
New-ADUser -SamAccountName "$sam" -Instance $u -DisplayName "$cn" -UserPrincipalName "$sam@loncc.local" -Name "$cn" -GivenName "$AD_User_first_name" -Surname "$AD_User_last_name" -EmailAddress $AD_User_email -HomeDirectory "C:\User_folders\$sam" -Path "$ou" -Enabled $True -ChangePasswordAtLogon $true -AccountPassword (ConvertTo-SecureString "$AD_User_pass" -AsPlainText -force)
# copying memberOf from original user
$m = Get-ADPrincipalGroupMembership -Identity $user_to_cp | where{$_.name -ne "Domain users"} # filters all MemberOf excelt Domain Users
$m  | foreach { Add-ADPrincipalGroupMembership -Identity $sam -MemberOf $_ -ErrorAction SilentlyContinue }
Write-Host "New user " -NoNewline; Write-Host $sam -NoNewline -ForegroundColor Green; Write-Host " has been created with properties from " -NoNewline; Write-Host "$user_to_cp." -NoNewline -ForegroundColor Green; Write-Host " in same OU " -NoNewline; Write-Host "$dep_name." -ForegroundColor Green
Read-Host "`nPress any key to exit"
}

welcome