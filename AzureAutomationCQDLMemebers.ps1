#created by Elazar Ohayon
  
 
 param
(
    [Parameter (Mandatory = $false)]
    [object] $WebhookData
)


# Retrieve Queue Number from Webhook request body
$Body = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)


$regex = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"


$emailList = [System.Collections.Generic.List[string]]::new()


$Body.Text | Select-String -AllMatches $regex


foreach ($user in $Body.Text) {
    $user -match $regex | Out-Null
    $emailList.Add($Matches.values)
}


$emailList
# Authenticate to msonline
#Change the name it to your Azure Credential in azure automation 
$Credential = Get-AutomationPSCredential -Name 'Teams Admin'
Connect-MsolService  -Credential $credential


$SecureGroup = Get-MsolGroup -all | where-object { $_.EmailAddress -eq "$emailList"}


 $DLGroup = $SecureGroup.ObjectId


# Authenticate to Teams
#Change the name it to your Azure Credential in azure automation 
$Credential = Get-AutomationPSCredential -Name 'Teams Admin'
Connect-MicrosoftTeams -Credential $credential


$neway = "we change your CQ to : "
# Set Queue Number

#adding your Call queue identity you can find it by run : get-csCallQueue | select Identity,name 

Set-CsCallQueue -Identity "adding Call queue Identity" -DistributionLists "$DLGroup"


#adding url of your incoming webhook

$URI = 'Adding your URI for incoming webhook'


# @type - Must be set to `MessageCard`.
# @context - Must be set to [`https://schema.org/extensions`](<https://schema.org/extensions>).
# title - The title of the card, usually used to announce the card.
# text - The card's purpose and what it may be describing.
# activityTitle - The title of the section, such as "Test Section", displayed in bold.
# activitySubtitle - A descriptive subtitle underneath the title.
# activityText - A longer description that is usually used to describe more relevant data.
$
$JSON = @{
  "@type"    = "MessageCard"
  "@context" = "<http://schema.org/extensions>"
  "title"    = 'Neway Support Bot  '
  "text"     = "$neway <br /> $emailList  "
  
  
  "sections" = @(
    @{
      "activityTitle"    = 'Completed'
      "activitySubtitle" = '100%: '
      "activityText"     = 'Have a nice day'
     
    }
  )
} | ConvertTo-JSON








# You will always be sending content in via POST and using the ContentType of 'application/json'
# The URI will be the URL that you previously retrieved when creating the webhook
$Params = @{
  "URI"         = $URI
  "Method"      = 'POST'
  "Body"        = $JSON
  "ContentType" = 'application/json'
}


Invoke-RestMethod @Params


Disconnect-MicrosoftTeams
