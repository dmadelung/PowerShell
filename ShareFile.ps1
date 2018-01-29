# Use this script to share a file via CSOM and PowerShell
# ShareObject https://msdn.microsoft.com/en-us/library/office/mt684216.aspx
# External sharing blog https://blogs.msdn.microsoft.com/vesku/2015/10/02/external-sharing-api-for-sharepoint-and-onedrive-for-business/

### ENTER YOU VARIABLES HERE ###

#path to the SP CSOM files 
$csomPath = "C:\" 

#Email of person running the script
$adminEmail = "user@domain.com"

#Site collection to be connected to
$siteUrl = "https://domain.sharepoint.com/sites/site"

#Library title where the file exists
$libraryTitle = "Documents" 

#Filename including file type
$fileName = "Test Document 1.docx"

#Email of who the document is being shared to
$emailSharedTo = "user2@domain.com"

#UNVALIDATED_EMAIL_ADDRESS if they are in AD or GUEST_USER if they are not
$principalType = "UNVALIDATED_EMAIL_ADDRESS"  

#role:1073741826 = View, role:1073741827 = Edit
$roleValue = "role:1073741827"

#A flag to determine if permissions should be pushed to items with unique permissions.
$propageAcl = $true

#Flag to determine if an e-mail notification should to sent, if e-mail is configured.
$sendEmail = $true  

#If an e-mail is being sent, this determines if an anonymous link should be added to the message.
$includedAnonymousLinkInEmail = $false  

#The ID of the group to be added to. Use zero if not adding to a permissions group. Not actually used by the code even when user is added to existing group. 
$groupId = 0

#Doesn't matter as it isn't sent in current email format
$emailSubject = ""

#Text for the body of the e-mail.
$emailBody = "Check out my email body"  

#Use modern sharing links instead of directly granting access
$useSimplifiedRoles = $true
################

# Get CSOM files
Add-type -Path "$csomPath\Microsoft.SharePoint.Client.dll"
Add-type -Path "$csomPath\Microsoft.SharePoint.Client.Runtime.dll"

# Connnect to site
$ss = Read-Host -Prompt "Enter admin password" -AsSecureString
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($adminEmail, $ss)
$ctx.Credentials = $creds
if(!$ctx.ServerObjectIsNull.Value) { 
    Write-Host "Connected to site:" $siteUrl -ForegroundColor Green 
} 
# Get web
$web = $ctx.Web

# Connect to library
$list = $web.Lists.GetByTitle($libraryTitle)
$ctx.Load($web)
$ctx.Load($list)
$ctx.Load($list.RootFolder)
$ctx.ExecuteQuery()

# Get doc
$query = New-Object Microsoft.SharePoint.Client.CamlQuery
$caml ="<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>" + $fileName + "</Value></Eq></Where></Query></View>"
$query.ViewXml = $caml
$item = $list.GetItems($query)
$ctx.Load($item)
$ctx.ExecuteQuery()
if (!$item) {
    Write-Host "Could not find the file:" $fileName -ForegroundColor Yellow 
} else {
    Write-Host "Sharing the the file:" $item.FieldValues.FileLeafRef -ForegroundColor Green 
}

# Get doc url
$itemUrl = $item.FieldValues.FileRef
$split = $web.Url -split '/'
$itemUrl = "https://" + $split[2] + $itemUrl

# Build user object to be shared to
$jsonPerson = "[{`"Key`":`"$emailSharedTo`",
`"Description`":`"$emailSharedTo`",
`"DisplayText`":`"$emailSharedTo`",
`"EntityType`":`"`",
`"ProviderDisplayName`":`"`",
`"ProviderName`":`"`",
`"IsResolved`":true,
`"EntityData`":{`"Email`":`"$emailSharedTo`",
    `"AccountName`":`"$emailSharedTo`",
    `"Title`":`"$emailSharedTo`",
    `"PrincipalType`":`"$principalType`"},
`"MultipleMatches`":[]}]"

# Initiate share
$result = [Microsoft.SharePoint.Client.Web]::ShareObject($web.Context,$itemUrl,$jsonPerson,$roleValue,$groupid,$propageAcl,$sendEmail,$includedAnonymousLinkInEmail,$emailSubject,$emailBody,$useSimplifiedRoles)
$web.Context.Load($result)
$web.Context.ExecuteQuery()

Write-Host "Status of the share:" $result.StatusCode -ForegroundColor Green