#region help
<#  
.SYNOPSIS
    

.DESCRIPTON
    Create vcard (.vcf) file from AD attributes and include a picture using Graph API if one is present.

.EXAMPLE
    ---EXAMPLE---
    C:\PS> .\Create_VCard.ps1 -user <user>

.PARAMETER user
    
.NOTES
    Created by: Josh Crabtree (@uc_crab) 16 Jan 2021
#>
#endregion #help

#region variables

#AD OU you want to run this against.
$OU = "<YOUR OU PATH HERE>"

#AD User properties - feel free to add and/or subtract from this list
$Properties = "company","displayname","department","givenname","l","mail","mobile","postalcode","sn","streetaddress","telephonenumber","title"

#You may want to modify things in the 'where-object' here.
$Users = ($OU | foreach{
    Get-ADUser -SearchBase $OU  -Properties $Properties -Filter * | Where-Object {($_.enabled -eq $true)}
})

#Get today's date.
$Today = (Get-Date).ToString('MM-dd-yyyy')

#Path to a log file to log all output.
$logfile = "C:\<YOUR PATH HERE>-$($Today).log"

#Graph API Token.
$token = '<TOKEN TO CONNECT TO GRAPH API GOES HERE>'

#Graph API Header Info.
$headers = @{
    authorization = $token
    'content-type' = 'image/jpeg'
}
#endregion variables

#region functions

#region Log-IT Function
#'Log-It' Function used for color-coded screen output and output to the log file.
function Log-It {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline=$True
        )]
        [String]$Message,
        [ValidateSet(
            "General","Process","Success","Failure","Warning","Notification","LogOnly","ScreenOnly"
        )]
        [String]$Status = "General"
    )
    Switch($Status){
        "General"{
            $Color="Cyan"
            $Type="[INFORMA] "
        }
        "Process"{
            $Color="White"
            $Type="[PROCESS] "
        }
        "Failure"{
            $Color="Red"
            $Type="[FAILURE] "
        }
        "Success"{
            $Color="Green"
            $Type="[SUCCESS] "
        }
        "Warning"{
            $Color="Yellow"
            $Type="[WARNING] "
        }
        "Notification"{
            $Color="Gray"
            $Type="[NOTICES] "
        }
        "ScreenOnly"{
            $Color="Magenta"
            $Type="[INFORMA] "
        }
        "LogOnly"{
            $Color=$Null
            $Type="[INFORMA] "
        }

    }
    if($Color -ne $Null){Write-Host -ForegroundColor $Color $Type$Message}
    if($Color -ne "Magenta"){"$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Type$Message" | Out-File $logfile -Append}
}
#endregion Log-IT Function

#region Create-VCard Function
#Function to create the VCard file. 
function Create-VCard {
    [CmdletBinding()]
    
    #user parameter to take a user from the pipeline
    param(
        [Parameter(
            Mandatory = $true,
            position = 0,
            ValueFromPipeline = $true
        )]
        $User
    )
    
    begin{
        Import-Module ActiveDirectory 
        $FN = "Create-VCard"
        "$FN | BEGIN: Create VCard Function." | Log-It -Status Notification  
        $endpoint = 'https://graph.microsoft.com/v1.0/users'
    }
    
    process{
        $upn = $user.userprincipalname
        
        #request url should look like this once concatenated - 'https://graph.microsoft.com/v1.0/users('<upn@domain.com>'/photo/$value')
        $requesturl = $endpoint + "('$upn')/photo/`$value"
        
        $VCardPath = "C:\<YOUR FILE PATH HERE>\$($user.displayname).vcf"
        $FilePath = "C:\<YOUR FILE PATH HERE>\$($user.displayname).jpg"
      
        #Test if VCard exists
        $OutputVCard = Test-Path $VCardPath 

        #If VCard doesnt exist, create the VCard
        If (!$outputvcard){
            $outputvcard = New-Item -Path $vCardPath -ItemType File -Force
            "$FN | Created new vcard for user: $($user.displayname)."| Log-it -Status Success
        }

        #Add AD attribute data to each VCard property
        Add-Content -Path $vCardPath "BEGIN:VCARD"
        Add-Content -Path $vCardPath "VERSION:3.0"
        Add-Content -Path $vCardPath ("N;LANGUAGE=en-us:" + $user.sn + ";" + $user.givenName)
        Add-Content -Path $vCardPath ("FN:" + $user.displayName)
        Add-Content -Path $vCardPath ("ORG:" + $user.company + ";" + $user.department)
        Add-Content -Path $vCardPath ("TITLE:" + $user.title)
        Add-Content -Path $vCardPath ("TEL;WORK;VOICE:" + $user.telephoneNumber)
        Add-Content -Path $vCardPath ("TEL;CELL;VOICE:" + $user.mobile)
        Add-Content -Path $vCardPath ("ADR;WORK;PREF:" +";;" +$user.streetAddress)
        Add-Content -Path $vCardPath ("EMAIL;PREF;INTERNET:" + $user.mail)

        $photo = Invoke-WebRequest -Method Get -Uri $requesturl -Headers $headers -OutFile $FilePath  
    
        #If a photo is not found, return that a photo wasn't found to convert. If one is found, convert it to Base64 format and set the photo in the vcard.
        if (!$filepath){
            "$FN | Cannot find a photo to convert for user: $($user.displayname). VCard has been created without a photo." | Log-It -Status Failure
        }
        else{ 
            #VCard expects a base64 string in the photo property. This takes the .jpg from Graph API and converts it to a base64 string        
            $GetPhoto = [convert]::ToBase64String((get-content -Path "D:\VCardTest\$($user.displayname).jpg" -encoding byte))
        }
        Add-Content -Path $vCardPath ("PHOTO;ENCODING=b;TYPE=JPEG:" + $GetPhoto)
        Add-Content -Path $vCardPath "END:VCARD"
    }
    
    end{
         "$FN | Ending script." | Log-it -Status Process
       }
}
#endregion Create-VCard Function

#endregion functions

#region processing

#Pipe your AD users created in the variables above to the Create-VCard function to create the VCard.
$Users | Create-VCard

#endregion processing
