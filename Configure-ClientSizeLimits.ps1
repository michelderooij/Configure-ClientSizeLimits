<#
    .SYNOPSIS
    Configure client-specific message size limits for Exchange Web Services, Outlook WebApp
    or ActiveSync workloads. Can run locally or remotely, from the Exchange Management Shell.
       
    Michel de Rooij
    michel@eightwone.com
    http://eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.4, July 19th, 2019

    .DESCRIPTION
    Configure client-specific message size limits. Specified limits are in 1KB units.
    See https://technet.microsoft.com/en-us/library/hh529949%28v=exchg.150%29.aspx
    
    .PARAMETER Server
    Specifies server to configure. When omitted, will configure local server.

    .PARAMETER AllServers
    Process all Exchange 2013/2016/2019 servers

    .PARAMETER OWA
    Specifies size limit of OWA requests in bytes.

    .PARAMETER EWS
    Specifies size limit of EWS requests in bytes.

    .PARAMETER EAS
    Specifies size limit of ActiveSync requests in bytes. Where a KB setting is configured,
    the value is rounded upwards.

    .PARAMETER Reset
    Indicate you want to run an IISRESET after reconfiguring web.config files.

    .PARAMETER NoBackup
    Switch to tell the script not to make backup copies of the config files changed.

    .LINK
    http://eightwone.com

    Revision History
    ---------------------------------------------------------------------
    1.0	  Initial release
    1.1   Added admin check for running locally
          When using all servers, processes server hosting EMS session last'
    1.11  Fixed issue running remotely against a server
    1.2   Added EAS parameter
          Changed unit size to bytes (you can use GB/MB/KB etc).
          Introduced changes to support Exchange 2016
          Added NoBackup switch to skip creating backup files
          Some code reformatting
    1.3   Added WhatIf and Confirm support
    1.31  Fixed updating of maxAllowedContentLength & maxRequestLength
          Some code optimizations
          Some cosmetics related to Ex2019
    1.4   Fixed setting EWS & EAS settings in wrong node

    .EXAMPLE
    Configure-ClientSizeLimits.ps1 -OWA 25MB -EWS 15MB -EAS 25MB
    Configure client size limit of 25MB for OWA, 15MB for EWS and 25MB for ActiveSync.
    
#>
#Requires -Version 3.0

[cmdletbinding(SupportsShouldProcess=$true, DefaultParameterSetName= 'Server')]
param(
  [parameter( Mandatory=$false, ParameterSetName = 'Server')]
    [string]$Server= $env:ComputerName,
    [parameter( Mandatory=$false, ParameterSetName = 'All')]
    [parameter( Mandatory=$false, ParameterSetName = 'Server')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$OWA,
    [parameter( Mandatory=$false, ParameterSetName = 'All')]
    [parameter( Mandatory=$false, ParameterSetName = 'Server')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$EWS,
    [parameter( Mandatory=$false, ParameterSetName = 'All')]
    [parameter( Mandatory=$false, ParameterSetName = 'Server')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$EAS,
    [parameter( Mandatory=$true, ParameterSetName = 'All')]
        [switch]$AllServers,
    [parameter( Mandatory=$false, ParameterSetName = 'All')]
    [parameter( Mandatory=$false, ParameterSetName = 'Server')]
        [switch]$Reset,
    [parameter( Mandatory=$false, ParameterSetName = 'All')]
    [parameter( Mandatory=$false, ParameterSetName = 'Server')]
        [switch]$NoBackup
)

process {

    $ERR_NOEMS                      = 1001
    $ERR_NOT201316SERVER            = 1002
    $ERR_CANTACCESSWEBCONFIG	    = 1004
    $ERR_RUNNINGNONADMINMODE        = 1010

    function Get-WCFFileName {
        param(
          $Identity,
          $ExInstallPath,
          $FileName
        )
        $WebConfigFile= '\\{0}\{1}${2}' -f $Identity, $ExInstallPath[0], (Split-Path (Join-Path -Path $ExInstallPath -ChildPath $FileName) -NoQualifier)
        If( -not (Test-Path -Path $WebConfigFile)) {
            Write-Error ('Can''t determine or access {0}' -f $WebConfigFile)
            Exit $ERR_CANTACCESSWEBCONFIG
        }
        return $WebConfigFile
    }

    Function is-Admin {
        $currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ( [Security.Principal.WindowsIdentity]::GetCurrent() )
        return ( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
    }
   
    Function RoundTo-KBUpward {
        param(
            [uint32]$Value
        )
        Return [uint32](( $Value + 1KB) / 1KB)
    }

    Function Configure-XMLAttribute {
        param(
            [ref]$WebConfig,
            $Path,
            $Attribute,
            $Value
        )
        If(! ($WebConfig.Value).SelectSingleNode( $Path)) {
            $Elements= $Path -split '/'
            $CurrentPath= '/'
            ForEach( $Element in $Elements) {
                If( $Element) {
                    $ProjectedPath= ('{0}/{1}' -f $CurrentPath, $Element )
                    If( ($WebConfig.Value).SelectSingleNode( $ProjectedPath)) {
                        Write-Verbose ('Node {0} exists' -f $ProjectedPath)
                    }
                    Else {
                        # Node doesn't exist - create element
                        Write-Verbose ('Creating node {0}' -f $ProjectedPath)
                        $NewNode= ($WebConfig.Value).CreateElement( $Element)
                        $null= ($WebConfig.Value).SelectSingleNode( $CurrentPath).AppendChild( $NewNode)
                    }
                    $CurrentPath= $ProjectedPath
                }
                Else {
                    # Empty element (root)
                }
            }
            $OldValue= 'N/A'
        }
        Else {
            Try {
                $OldValue= ($WebConfig.Value).SelectSingleNode( $Path).GetAttribute( $Attribute)
            }
            Catch {
                $OldValue= ''
            }
        }
        Write-Verbose ('Set {0} attribute {1}: {2} (was {3})' -f $Path, $Attribute, $Value, $OldValue)
        ($WebConfig.Value).SelectSingleNode( $Path).SetAttribute( $Attribute, [string]$Value)
    }

    Function Get-XMLPath {
        param(
            $Element,
            $Name
        )
        If( $Element -ne $null) {
            If( $Element.ParentNode) {
                $Path= '{0}/{1}' -f (Get-XMLPath -Element $Element.ParentNode -Name $null), $Element.Name 
            }
            Else {
                $Path= '/'
            }
        }
        Else {
            # NOP
        }
        return $Path
    }

    # MAIN
    Try {
        $null= Get-ExchangeServer -ErrorAction SilentlyContinue
    }
    Catch {
        Write-Error 'Exchange Management Shell not loaded'
        Exit $ERR_NOEMS
    }

    If( $AllServers) {
        If(! ( is-Admin)) {
            Write-Error 'Script requires running in elevated mode'
            Exit $ERR_RUNNINGNONADMINMODE
        }
        $ServerList= Get-ExchangeServer | Where-Object { ($_.AdminDisplayVersion).Major -eq 15 } | Sort-Object -Property { $_.Fqdn -eq (Get-PSSession).ComputerName }
    }
    Else {
        If( (Get-ExchangeServer -Identity $Server).adminDisplayVersion.Major -ne 15) {
            Write-Error ('{0} appears not to be an Exchange 2013/2016/2019 server' -f $Server)
            Exit $ERR_NOT201316SERVER
        }
        $ServerList= @( $Server)
    }

    ForEach( $Identity in $ServerList) {

        $ThisServer= Get-ExchangeServer -Identity $Identity
        $Version= $ThisServer.AdminDisplayVersion.Major
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Identity)
        $ExInstallPath = $reg.OpenSubKey(('SOFTWARE\Microsoft\ExchangeServer\v{0}\Setup' -f $Version)).GetValue('MsiInstallPath')

        Write-Verbose ('Processing server {0} (v{1}), InstallPath {2} ..' -f $Identity, $Version, $ExInstallPath)

        if( $EWS) {
            If( $ThisServer.isClientAccessServer) {
                $wcfFile= Get-WCFFileName -Identity $Identity -ExInstallPath $ExInstallPath -FileName 'FrontEnd\HttpProxy\ews\web.config'
                $wcfXML= [xml](Get-Content -Path $wcfFile)
                Write-Output ('Processing {0}' -f $wcfFile)
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/system.webServer/security/requestFiltering/requestLimits' -Attribute 'maxAllowedContentLength' -Value $EWS
                If(! $NoBackup) {
                    Copy-Item -Path $wcfFile -Destination ('{0}_{1}.bak' -f $wcfFile, (Get-Date).toString('yyyMMddHHmmss')) -Force
                }
                if ($pscmdlet.ShouldProcess($wcfFile, "Saving modifications")) {
                    $wcfXML.Save( $wcfFile)
                }
            }
            If( $ThisServer.isMailboxServer) {
                $wcfFile= Get-WCFFileName -Identity $Identity  -ExInstallPath $ExInstallPath -FileName 'ClientAccess\exchweb\ews\web.config'
                $wcfXML= [xml](Get-Content -Path $wcfFile)
                Write-Output ('Processing {0}' -f $wcfFile)
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/system.webServer/security/requestFiltering/requestLimits' -Attribute 'maxAllowedContentLength' -Value $EWS
                $Elem= $wcfXML.SelectNodes('//*[@maxReceivedMessageSize]') | Where-Object { $_.ParentNode.ChildNodes.Name -notcontains 'UMLegacyMessageEncoderSoap11Element' }
                $Name= 'maxReceivedMessageSize'
                $Value= [string]($EWS) 
                $Elem | ForEach-Object { 
                    $Path= Get-XMLPath -Element $_ -Name $_.Name
                    $OldValue= $_.GetAttribute( $Name)
                    Write-Verbose ('Set {0} attribute {1}: {2} (was {3})' -f $Path, $Name, $Value, $OldValue)
                    $_.maxReceivedMessageSize= $Value
                }
                If(! $NoBackup) {
                    Copy-Item -Path $wcfFile -Destination ('{0}_{1}.bak' -f $wcfFile, (Get-Date).toString('yyyMMddHHmmss')) -Force
                }
                if ($pscmdlet.ShouldProcess($wcfFile, "Saving modifications")) {
                    $wcfXML.Save( $wcfFile)
                }
            }
        }

        if( $OWA) {
            If( $ThisServer.isClientAccessServer) {
                $wcfFile= Get-WCFFileName -Identity $Identity -ExInstallPath $ExInstallPath -FileName 'FrontEnd\HttpProxy\owa\web.config'
                $wcfXML= [xml](Get-Content -Path $wcfFile)
                Write-Output ('Processing {0}' -f $wcfFile)
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/location/system.webServer/security/requestFiltering/requestLimits' -Attribute 'maxAllowedContentLength' -Value $OWA
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/location/system.web/httpRuntime' -Attribute 'maxRequestLength' -Value (RoundTo-KBUpward -Value $OWA)
                If(! $NoBackup) {
                    Copy-Item -Path $wcfFile -Destination ('{0}_{1}.bak' -f $wcfFile, (Get-Date).toString('yyyMMddHHmmss')) -Force
                }
                if ($pscmdlet.ShouldProcess($wcfFile, "Saving modifications")) {
                    $wcfXML.Save( $wcfFile)
                }
            }
            If( $ThisServer.isMailboxServer) {
                $wcfFile= Get-WCFFileName -Identity $Identity  -ExInstallPath $ExInstallPath -FileName 'ClientAccess\Owa\web.config'
                $wcfXML= [xml](Get-Content -Path $wcfFile)
                Write-Output ('Processing {0}' -f $wcfFile)
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/location/system.webServer/security/requestFiltering/requestLimits' -Attribute 'maxAllowedContentLength' -Value $OWA
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/location/system.web/httpRuntime' -Attribute 'maxRequestLength' -Value (RoundTo-KBUpward -Value $OWA)
                $Elem= $wcfXML.SelectNodes('//*[@maxReceivedMessageSize]')
                $Name= 'maxReceivedMessageSize'
                $Value= [string]$OWA
                $Elem | ForEach-Object { 
                    $Path= Get-XMLPath -Element $_ -Name $_.Name
                    $OldValue= $_.GetAttribute( $Name)
                    Write-Verbose ('Set {0} attribute {1}: {2} (was {3})' -f $Path, $Name, $Value, $OldValue)
                    $_.maxReceivedMessageSize= $Value
                }
                $Elem= $wcfXML.SelectNodes('//*[@maxStringContentLength]') | Where-Object { $_.ParentNode.Name -ne 'MsOnlineShellService_BindingConfiguration' }
                $Name= 'maxStringContentLength'
                $Value= [string]$OWA
                $Elem | ForEach-Object { 
                    $Path= Get-XMLPath -Element $_ -Name $_.Name
                    $OldValue= $_.GetAttribute( $Name)
                    Write-Verbose ('Set {0} attribute {1} to {2} (was {3})' -f $Path, $Name, $Value, $OldValue)
                    $_.maxStringContentLength= $Value
                }
                If(! $NoBackup) {
                    Copy-Item -Path $wcfFile -Destination ('{0}_{1}.bak' -f $wcfFile, (Get-Date).toString('yyyMMddHHmmss')) -Force
                }
                if ($pscmdlet.ShouldProcess($wcfFile, "Saving modifications")) {
                    $wcfXML.Save( $wcfFile)
                }
            }
        }

        if( $EAS) {
            If( $ThisServer.isClientAccessServer) {
                $wcfFile= Get-WCFFileName -Identity $Identity  -ExInstallPath $ExInstallPath -FileName 'FrontEnd\HttpProxy\Sync\web.config'
                $wcfXML= [xml](Get-Content -Path $wcfFile)
                Write-Output ('Processing {0}' -f $wcfFile)
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/system.webServer/security/requestFiltering/requestLimits' -Attribute 'maxAllowedContentLength' -Value $EAS
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/system.web/httpRuntime' -Attribute 'maxRequestLength' -Value (RoundTo-KBUpward -Value $EAS)
                If(! $NoBackup) {
                    Copy-Item -Path $wcfFile -Destination ('{0}_{1}.bak' -f $wcfFile, (Get-Date).toString('yyyMMddHHmmss')) -Force
                }
                if ($pscmdlet.ShouldProcess($wcfFile, "Saving modifications")) {
                    $wcfXML.Save( $wcfFile)
                }
            }
            If( $ThisServer.isMailboxServer) {
                $wcfFile= Get-WCFFileName -Identity $Identity  -ExInstallPath $ExInstallPath -FileName 'ClientAccess\Sync\web.config'
                $wcfXML= [xml](Get-Content -Path $wcfFile)
                Write-Output ('Processing {0}' -f $wcfFile)
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/system.webServer/security/requestFiltering/requestLimits' -Attribute 'maxAllowedContentLength' -Value $EAS
                Configure-XMLAttribute -WebConfig ([ref]$wcfXML) -Path '//configuration/system.web/httpRuntime' -Attribute 'maxRequestLength' -Value (RoundTo-KBUpward -Value $EAS)

                $Elem= $wcfXML.SelectSingleNode('//configuration/appSettings/add[@key="MaxDocumentDataSize"]')
                $Name= 'MaxDocumentDataSize'
                $Value= [string]$EAS
                $Path= Get-XMLPath -Element $Elem -Name $Elem.Name
                $OldValue= $Elem.Value
                Write-Verbose ('Set {0} value {1} to {2} (was {3})' -f $Path, $Name, $Value, $OldValue)
                $Elem.Value= $Value

                If(! $NoBackup) {
                    Copy-Item -Path $wcfFile -Destination ('{0}_{1}.bak' -f $wcfFile, (Get-Date).toString('yyyMMddHHmmss')) -Force
                }
                if ($pscmdlet.ShouldProcess($wcfFile, "Saving modifications")) {
                    $wcfXML.Save( $wcfFile)
                }
            }
        }
    }
    If( $Reset) {
        ForEach( $Identity in $ServerList) {
            Write-Output ('Restarting IIS on {0}' -f $Identity)
            IISRESET.EXE $Identity /NOFORCE /TIMEOUT:300
        }
    } 
}