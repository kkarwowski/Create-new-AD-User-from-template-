---

![alt text](https://github.com/kkarwowski/Gifs/blob/master/copy%20ad.gif  "GIF showing the script")


# About

This PowerShell Script assist in making new AD users in your company. Usually you would copy existing AD user rights (from the same department , with the same job title as the starter) and fill out the name, email password etc. 

I create Excel file which has New Users name, email, job title, department, AD user name and AD password. This file is stored in secure drive with limited access by AD Admins. This file is then printed and given to the starter.

Above script will take name, email, job title, AD user name and AD password from selected Excel file and copy those details when creating AD credentials for this user. No more manual copying! Script will ask you to choose existing AD user to copy details from. 


# Before running

In user_account_excel folder you will find sample Excel file with new starter details. Script takes info from specific Cells of this file. You can edit those in the script as you wish.

You may also change directory where the Excel files are located by adjusting below part of the code:

```powershell
$dir_with_excel = $PSScriptRoot+'\user_accounts_excel'
```
This variable is a Path of OU which will be searched by script for users. This usually is called Users but may be different in your Domain.

```powershell
$All_ous = 'OU=London,DC=loncc,DC=local'
```


# Usage

Run:
```bash
powershell Copy AD USER.ps1
```
