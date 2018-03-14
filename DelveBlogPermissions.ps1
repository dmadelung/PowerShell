# Use this script to remove viwer permissions from all user delve blogs that have been created
# A user will still be able to view their existing blogs and create blogs but people will not be able to see them
# This would allow you to choose in the future if you want to make them live
# 
# This could be updated to run on a schedule as this will not remove any new blogs that are created

### ENTER YOU VARIABLES HERE ###

#Path to the SP CSOM files 
$csomPath = "C:\...." 
################

#Prompt for parameters
#TenantDomain is beginning of "tenantdomain.sharepoint.com.."
$TenantDomain = Read-Host -Prompt "Tenant domain"
$AdminAccount = Read-Host -Prompt "Admin account"
$AdminPass = Read-Host -Prompt "Password for $AdminAccount" –AsSecureString

#Set SharePoint admin url
$AdminURI = "https://" + $TenantDomain + "-admin.sharepoint.com"

#Get CSOM files
Add-type -Path "$csomPath\Microsoft.SharePoint.Client.dll"
Add-type -Path "$csomPath\Microsoft.SharePoint.Client.Runtime.dll"

#Begin the process
$loadInfo1 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
$loadInfo2 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
$loadInfo3 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")

#Set credentials for CSOM
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminAccount, $AdminPass)

#Add the path of the User Profile Service to the SPO admin URL, then create a new webservice proxy to access it
$proxyaddr = "$AdminURI/_vti_bin/UserProfileService.asmx?wsdl"
$UserProfileService= New-WebServiceProxy -Uri $proxyaddr -UseDefaultCredential False
$UserProfileService.Credentials = $creds

#Set variables for authentication cookies
$strAuthCookie = $creds.GetAuthenticationCookie($AdminURI)
$uri = New-Object System.Uri($AdminURI)
$container = New-Object System.Net.CookieContainer
$container.SetCookies($uri, $strAuthCookie)
$UserProfileService.CookieContainer = $container

#Sets the first User profile, at index -1
$UserProfileResult = $UserProfileService.GetUserProfileByIndex(-1)

Write-Host "Starting- This could take a while."

#Getting total number of profiles
$NumProfiles = $UserProfileService.GetUserProfileCount()
$i = 1

#Create array to track users
$users = @()

#As long as the next User profile is NOT the one we started with (at -1)...
While ($UserProfileResult.NextValue -ne -1) 
{
    Write-Host "Reviewing profile $i of $NumProfiles"

    #Look for the Point Publishing Blog url object in the User Profile and retrieve it
    #It will be empty for users which it has not been created for

    #Get personal blog publishing URL
    $Prop = $UserProfileResult.UserProfile | Where-Object { $_.Name -eq "SPS-PointPublishingUrl" } 
    $Url= $Prop.Values[0].Value

    #Get user UPN - Can be used for reporting
    #$Prop = $userProfileResult.UserProfile | Where-Object { $_.Name -eq "SPS-UserPrincipalName"}
    #$Upn= $Prop.Values[0].Value

    #If the blog site exists then add it to an array to review
    if ($Url) {
        $users += $Url
    }

    #And now we check the next profile the same way...
    $UserProfileResult = $UserProfileService.GetUserProfileByIndex($UserProfileResult.NextValue)
    $i++
}

#Loop through all identified sites to remove blog viewers
foreach($user in $users){
    #Set blog site url
    $siteurl = "https://" + $TenantDomain + ".sharepoint.com" + $user

    #Connect to blog site collection
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteurl)
    $ctx.Credentials = $Creds
 
    #Connect to web and get site groups
    $web = $ctx.Web
    $groups = $ctx.Web.SiteGroups
    $ctx.Load($web)
    $ctx.Load($groups)
    $ctx.ExecuteQuery()
    
    #Get the viewers group
    $group = $groups | where { $_.Title -eq "Viewers"}
    if($group){
        #Get the users in the viewers group
        $users = $group.Users
        $ctx.Load($users)
        $ctx.ExecuteQuery()

        #Remove all users from the viewers group
        foreach($u in $users){
            $group.Users.RemoveByLoginName($u.LoginName)
            $web.Update()
            $ctx.ExecuteQuery()
        }
    }
}