$mailboxGuid = $datasource.selectedMailbox.Guid

# Used to connect to Exchange using user credentials (MFA not supported).
$ConnectionUri = $ExchangeConnectionUri
$Username = $ExchangeUsername
$Password = $ExchangePassword
$AuthenticationMethod = $ExchangeAuthenticationMethod

# PowerShell commands to import
$commands = @(
    "Get-Mailbox"
)

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ''
        }

        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
            # $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message # Does not show the correct error message for the Raet IAM API calls
            $httpErrorObj.ErrorMessage = $ErrorObject.Exception.Message

        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [HelloID.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }

        Write-Output $httpErrorObj
    }
}

function Get-ErrorMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $errorMessage = [PSCustomObject]@{
            VerboseErrorMessage = $null
            AuditErrorMessage   = $null
        }

        if ( $($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException')) {
            $httpErrorObject = Resolve-HTTPError -Error $ErrorObject

            $errorMessage.VerboseErrorMessage = $httpErrorObject.ErrorMessage

            $errorMessage.AuditErrorMessage = $httpErrorObject.ErrorMessage
        }

        # If error message empty, fall back on $ex.Exception.Message
        if ([String]::IsNullOrEmpty($errorMessage.VerboseErrorMessage)) {
            $errorMessage.VerboseErrorMessage = $ErrorObject.Exception.Message
        }
        if ([String]::IsNullOrEmpty($errorMessage.AuditErrorMessage)) {
            $errorMessage.AuditErrorMessage = $ErrorObject.Exception.Message
        }

        Write-Output $errorMessage
    }
}
#endregion functions

try {
    #region Connect to Exchange
    try {
        Write-Verbose "Connecting to Exchange: $connectionUri"
    
        # Connect to Exchange in an unattended scripting scenario using user credentials (MFA not supported).
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $credential = [System.Management.Automation.PSCredential]::new($username, $securePassword)
        $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -IdleTimeout (New-TimeSpan -Minutes 5).TotalMilliseconds # The session does not time out while the session is active. Please enter this value to time out the Exchangesession when the session is removed
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $credential -Authentication $authenticationMethod -AllowRedirection -SessionOption $sessionOption -EnableNetworkAccess:$false -ErrorAction Stop
        $null = Import-PSSession $exchangeSession -CommandName $commands -AllowClobber

        Write-Information "Successfully connected to Exchange: $connectionUri"
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

        throw "Error connecting to Exchange: $connectionUri. Error Message: $($errorMessage.AuditErrorMessage)"
    }
    #endregion Connect to Exchange

    try {
        $exchangeQuerySplatParams = @{
            Identity   = $mailboxGuid
            ResultSize = "Unlimited"
        }

        Write-Information "Querying SendOnBehalf permissions to mailbox [$($exchangeQuerySplatParams.Identity)]"
        $mailboxPermissions = (Get-Mailbox @exchangeQuerySplatParams -ErrorAction Stop).GrantSendOnBehalfTo

        $resultCount = ($mailboxPermissions | Measure-Object).Count
        Write-Information "Result count: $resultCount"
    
        if ($resultCount -gt 0) {
            foreach ($mailboxPermission in $mailboxPermissions) {
                Write-Output $mailboxPermission
            }
        }
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

        throw "Error querying SendOnBehalf permissions to mailbox [$($exchangeQuerySplatParams.Identity)]. Error Message: $($errorMessage.AuditErrorMessage)"
    }
}
catch {
    $ex = $PSItem
    $errorMessage = Get-ErrorMessage -ErrorObject $ex

    Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($($errorMessage.VerboseErrorMessage))"

    Write-Error "Error querying SendOnBehalf permissions to mailbox [$($mailboxGuid)]. Error Message: $($errorMessage.AuditErrorMessage)"
}