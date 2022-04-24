Param (
    # username to authenticate to target server
    [Parameter(Mandatory = $true)]
    [string]$DomainUsername,

    # password to authenticate to target server
    [Parameter(Mandatory = $true)]
    [string]$DomainPword,

    [Parameter(Mandatory = $true)]
    [string]$EnvName,

    [Parameter(Mandatory = $true)]
    [string]$Operation

)

$DomainAccountPassword = ConvertTo-SecureString -String $DomainPword -AsPlainText -Force
$DomainCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $DomainUsername, $DomainAccountPassword
$so = New-PSSessionOption -SkipCACheck -SkipCNCheck


$configFilename = ".\json\iis_instance_prod.json"
$instanceConfig = $(Get-Content -Raw $configFilename | ConvertFrom-Json)
$env = $instanceConfig.environments | Where-Object { $_.name -eq $EnvName }
$poolInstances = $env.web_instances

	
foreach ($instance in $poolInstances) {
    Write-Output $instance
    Invoke-Command -ComputerName $instance -Credential $DomainCredential -UseSSL -SessionOption $so -ScriptBlock {
        & { iisreset /$using:Operation }   
    }
}
