#region Functions
Function Get-AuthTokenMSAL {

    <#
    .SYNOPSIS
    This function is used to authenticate with the Graph API REST interface
    .DESCRIPTION
    The function authenticate with the Graph API Interface with the tenant name
    .EXAMPLE
    Get-AuthTokenMSAL
    Authenticates you with the Graph API interface using MSAL.PS module
    .NOTES
    NAME: Get-AuthTokenMSAL
    #>

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        $User
    )

    $userUpn = New-Object 'System.Net.Mail.MailAddress' -ArgumentList $User
    if ($userUpn.Host -like '*onmicrosoft.com*') {
        $tenant = Read-Host -Prompt 'Please specify your Tenant name i.e. company.com'
        Write-Host
    }
    else {
        $tenant = $userUpn.Host
    }

    Write-Host 'Checking for MSAL.PS module...'
    $MSALModule = Get-Module -Name 'MSAL.PS' -ListAvailable
    if ($null -eq $MSALModule) {
        Write-Host
        Write-Host 'MSAL.PS Powershell module not installed...' -f Red
        Write-Host "Install by running 'Install-Module MSAL.PS -Scope CurrentUser' from an elevated PowerShell prompt" -f Yellow
        Write-Host "Script can't continue..." -f Red
        Write-Host
        exit
    }
    if ($MSALModule.count -gt 1) {
        $Latest_Version = ($MSALModule | Select-Object version | Sort-Object)[-1]
        $MSALModule = $MSALModule | Where-Object { $_.version -eq $Latest_Version.version }
        # Checking if there are multiple versions of the same module found
        if ($MSALModule.count -gt 1) {
            $MSALModule = $MSALModule | Select-Object -Unique
        }
    }

    $ClientId = 'd1ddf0e4-d672-4dae-b554-9d5bdfd93547'
    $RedirectUri = 'urn:ietf:wg:oauth:2.0:oob'
    $Authority = "https://login.microsoftonline.com/$Tenant"

    try {
        Import-Module $MSALModule.Name
        if ($PSVersionTable.PSVersion.Major -ne 7) {
            $authResult = Get-MsalToken -ClientId $ClientId -Interactive -RedirectUri $RedirectUri -Authority $Authority
        }
        else {
            $authResult = Get-MsalToken -ClientId $ClientId -Interactive -RedirectUri $RedirectUri -Authority $Authority -DeviceCode
        }
        # If the accesstoken is valid then create the authentication header
        if ($authResult.AccessToken) {
            # Creating header for Authorization token
            $authHeader = @{
                'Content-Type'  = 'application/json'
                'Authorization' = 'Bearer ' + $authResult.AccessToken
                'ExpiresOn'     = $authResult.ExpiresOn
            }
            return $authHeader
        }
        else {
            Write-Host
            Write-Host 'Authorization Access Token is null, please re-run authentication...' -ForegroundColor Red
            Write-Host
            break
        }
    }
    catch {
        Write-Host $_.Exception.Message -f Red
        Write-Host $_.Exception.ItemName -f Red
        Write-Host
        break
    }
}
Function Add-GoogleApplication() {

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        $PackageID
    )

    $graphApiVersion = 'Beta'
    $App_resource = 'deviceManagement/androidManagedStoreAccountEnterpriseSettings/approveApps'

    try {

        $PackageID = 'app:' + $PackageID
        $Packages = New-Object -TypeName psobject
        $Packages | Add-Member -MemberType NoteProperty -Name 'approveAllPermissions' -Value 'true'
        $Packages | Add-Member -MemberType NoteProperty -Name 'packageIds' -Value @($PackageID)
        $JSON = $Packages | ConvertTo-Json -Depth 3

        $uri = "https://graph.microsoft.com/$graphApiVersion/$($App_resource)"
        Invoke-RestMethod -Uri $uri -Method Post -ContentType 'application/json' -Body $JSON -Headers $authToken
        Write-Host "Successfully added $PackageID from Managed Google Store" -ForegroundColor Green

    }

    catch {
        $exs = $Error.ErrorDetails
        $ex = $exs[0]
        Write-Host "Response content:`n$ex" -f Red
        Write-Host
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Message)"
        Write-Host
        break
    }

}
Function Invoke-SyncGoogleApplication() {

    [cmdletbinding()]

    $graphApiVersion = 'Beta'
    $App_resource = '/deviceManagement/androidManagedStoreAccountEnterpriseSettings/syncApps'

    try {

        $uri = "https://graph.microsoft.com/$graphApiVersion/$($App_resource)"
        Invoke-RestMethod -Uri $uri -Method Post -ContentType 'application/json' -Body $JSON -Headers $authToken
        Write-Host 'Successfully synchronised Google Apps' -ForegroundColor Green

    }

    catch {
        $exs = $Error.ErrorDetails
        $ex = $exs[0]
        Write-Host "Response content:`n$ex" -f Red
        Write-Host
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Message)"
        Write-Host
        break
    }

}

#endregion Functions

#region Authentication
# Checking if authToken exists before running authentication
if ($global:authToken) {

    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

    if ($TokenExpires -le 0) {

        Write-Host 'Authentication Token expired' $TokenExpires 'minutes ago' -ForegroundColor Yellow
        Write-Host

        # Defining User Principal Name if not present

        if ($null -eq $User -or $User -eq '') {

            $User = Read-Host -Prompt 'Please specify your user principal name for Azure Authentication'
            Write-Host

        }

        $global:authToken = Get-AuthTokenMSAL -User $User

    }
}

# Authentication doesn't exist, calling Get-AuthToken function

else {

    if ($null -eq $User -or $User -eq '') {

        $User = Read-Host -Prompt 'Please specify your user principal name for Azure Authentication'
        Write-Host

    }

    # Getting the authorization token
    $global:authToken = Get-AuthTokenMSAL -User $User
    Write-Host 'Connected to Graph API' -ForegroundColor Green
    Write-Host
}

#endregion

#region Script
$AndroidAppIds = New-Object -TypeName System.Collections.ArrayList
$AndroidAppIds.AddRange(@(
    'com.gotomeeting',
    'com.microsoft.mobile.polymer',
    'com.sec.android.gallery3d',
    'com.tma.fungi',
    'com.velux.roof_pitch',
    'com.conferma.trippay',
    'com.pb.confirm.connect',
    'com.wordwebsoftware.android.wordweb',
    'com.mm.android.DMSS',
    'com.docusign.ink',
    'com.nuance.dragonanywhere',
    'com.dropbox.android',
    'com.duosecurity.duomobile',
    'com.digitalbarriers.viewer',
    'com.egress.switchdroid',
    'com.ehi.csma',
    'com.entrust.identityGuard.mobile',
    'com.letsenvision.envisionai',
    'epson.print',
    'uk.gov.HomeOffice.ho1',
    'com.eventbrite.attendee',
    'com.eventbrite.organizer',
    'com.facebook.katana',
    'com.cube.rca',
    'net.flexiroute.driver2',
    'air.uk.co.nhbc.depthcalcplus',
    'com.samsung.accessory.neobeanmgr',
    'com.samsung.android.waterplugin',
    'com.samsung.android.app.watchmanager2',
    'com.samsung.android.app.watchmanager',
    'com.genie.companion',
    'com.geniecpms.GeniePointMobile',
    'de.gfa.GfAPlus',
    'com.google.android.googlequicksearchbox',
    'com.google.android.apps.authenticator2',
    'com.android.chrome',
    'com.google.android.apps.classroom',
    'com.google.android.apps.fitness',
    'com.google.android.apps.maps',
    'com.google.android.apps.translate',
    'com.gopro.smarty',
    'uk.co.gowash.washapp',
    'com.qingniu.HealthKeep',
    'com.frs.hwprodtest',
    'com.hootsuite.droid.full',
    'com.chasingthestigma.hubofhope',
    'com.humanware.hwbuddy',
    'com.acs.nomad.app',
    'com.instagram.android',
    'com.microsoft.windowsintune.companyportal',
    'org.ipaf.epal',
    'uk.gov.cheshirewestandchester.itravelsmart',
    'com.alivecor.aliveecg',
    'com.callpod.android_apps.keeper',
    'com.nexstreaming.app.kinemasterfree',
    'uk.co.knifesavers',
    'de.komoot.android',
    'com.lazarillo',
    'cn.huidu.huiduapp',
    'com.linkedin.android',
    'com.google.audio.hearing.visualization.accessibility.scribe',
    'com.LYD',
    'com.microsoft.launcher.enterprise',
    'com.facebook.orca',
    'com.facebook.pages.app',
    'com.microsoft.office.officehubrow',
    'com.azure.authenticator',
    'com.microsoft.emmx',
    'com.microsoft.office.excel',
    'com.microsoft.launcher',
    'com.microsoft.skydrive',
    'com.microsoft.office.onenote',
    'com.microsoft.office.outlook',
    'com.microsoft.planner',
    'com.microsoft.sharepoint',
    'com.microsoft.stream',
    'com.microsoft.teams',
    'com.microsoft.todos',
    'com.microsoft.office.word',
    'mobile.appC60wOPPP50',
    'uk.co.modernmindset.xapp',
    'com.monday.monday',
    'com.tranzmate',
    'com.mrfill.wastemanager',
    'com.app.p3681GD',
    'net.iplato.mygp',
    'com.WORKSuite.MyTime_WORKSuite',
    'uk.co.jaama.myvehicle',
    'com.neosistec.NaviLens',
    'tv.netweather.netweatherradar',
    'com.wilysis.cellinfolite',
    'uk.ac.shef.oak.pheactiveten',
    'com.nhs.online.nhsonline',
    'com.phe.couchto5K',
    'com.phe.daysoff',
    'com.doh.smokefree',
    'com.antbits.nhsSafeguardingGuide',
    'com.nhs.weightloss',
    'app.envitech.max.Northlincs',
    'net.sourceforge.opencamera',
    'com.staircase3.opensignal',
    'org.PSSLive.pssLive',
    'uk.co.ordnancesurvey.oslocate.android',
    'uk.co.ordnancesurvey.osmaps',
    'uk.co.patient.patientaccess',
    'com.paybyphone',
    'biz.peopleplanner.PPMobileV3',
    'com.pixlr.express',
    'tascomi.planningapp',
    'com.fws.plantsnap2',
    'com.microsoft.msapps',
    'com.led595.powerledlts',
    'com.pssltd.pssliveplus2',
    'ch.opengis.qfield',
    'com.r2conline.ddc',
    'uk.org.ramblers.walkreg',
    'com.goassemble.ramblers',
    'com.bt.relayuk',
    'com.logmein.rescuemobile',
    'com.ringapp',
    'com.rotacloud',
    'com.s12solutions.s12',
    'com.safetyculture.iauditor',
    'com.sec.android.app.popupcalculator',
    'com.sec.android.app.shealth',
    'com.samsung.android.app.notes',
    'com.sec.android.app.voicenote',
    'com.microsoft.seeingai',
    'org.thoughtcrime.securesms',
    'com.visa.spendmanagement',
    'com.spotify.music',
    'com.sproutsocial.android',
    'com.sja.firstaid',
    'com.msearcher.taptapsee.android',
    'com.tascomi.buildapp',
    'com.teamviewer.quicksupport.market',
    'com.tesco.grocery.view',
    'uk.ac.jisc.govroam',
    'com.instagram.barcelona',
    'com.zhiliaoapp.musically',
    'com.thetileapp.tile',
    'com.thetrainline',
    'com.trello',
    'uk.org.alcoholconcern.dryjanuary',
    'uk.nhs.covid19.production',
    'app.envitech.UKAir',
    'com.nhs.verify.care.id',
    'com.frontrow.vlog',
    'com.app.p8238DE',
    'com.waze',
    'com.what3words.android',
    'com.whatsapp',
    'com.twitter.android',
    'com.google.android.youtube',
    'com.zapmap.zapmap',
    'us.zoom.videomeetings' 
    )
)

foreach ($AndroidAppId in $AndroidAppIds) {
    Add-GoogleApplication -PackageID $AndroidAppId
}
Invoke-SyncGoogleApplication
#endregion Script