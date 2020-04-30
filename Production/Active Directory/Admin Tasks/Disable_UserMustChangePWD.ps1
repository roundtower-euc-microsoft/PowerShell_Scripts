[4/14 5:01 PM] Corey St. Pierre
    $PLSValue = 0
#0 to enable the User must change password at next logon option
#-1 to disable the User must change password at next logon option
$ObjFilter = "(&(objectCategory=person)(objectCategory=User))"
    $objSearch = New-Object System.DirectoryServices.DirectorySearcher
    $objSearch.PageSize = 15000
    $objSearch.Filter = $ObjFilter
    $objSearch.SearchRoot = "LDAP://OU=User Accounts,DC=santhosh,DC=lab"
    $AllObj = $objSearch.FindAll()
    foreach ($Obj in $AllObj)
           {
            $objItemS = $Obj.Properties
            $UserN = $objItemS.name
            $UserDN = $objItemS.distinguishedname
            $user = [ADSI] "LDAP://$userDN"
            $user.psbase.invokeSet("pwdLastSet",$PLSValue)
            Write-host -NoNewLine "Modifying $UserN Properties...."
            $user.setinfo()
            Write-host "Done!"
            }

