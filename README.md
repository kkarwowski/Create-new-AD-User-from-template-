---

![alt text](https://github.com/kkarwowski/Gifs/blob/master/copy%20ad.gif  "GIF showing the script")


# About

This PowerShell Script assist in making new AD users in your company. Ususally you would copy existing AD user rights ( from the same department , with the same job title as the starter) and fill out the name, email password etc. 

I create Excel file which has New Users name, email, job title, deparment, AD user name and AD password. This file is stored in secure drive with limited access by AD Admins. This file is then printed and given to the starter.

Above script will take name, email, job title, AD user namme and AD password from selected Excel file and copy those details when creating AD credentials for this user. No more manual copying! Script will ask you to choose exisitng AD user to copy details from. 


# Installation

Clone this repository and install [requirements.txt](requirements.txt) dependencies.
```bash
pip install -r requirements.txt
```

# Usage

Run:
```bash
powershell Copy AD USER.ps1
```
 	
