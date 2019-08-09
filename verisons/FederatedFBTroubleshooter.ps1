#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################
#                                                                               #
#                      Version 0.1 beta/initial release                         #
#                                                                               #
#################################################################################

Param(
[switch]$debug #not functional yet
)

[bool]$script:Exchange2010 = $false
[bool]$script:Exchange2013 = $false
[bool]$script:Exchange2016 = $false

function GetLocalServerInfo
{
    $localServerName = ${env:computername}
    $Error.Clear()
    $localServerInfo = Get-ExchangeServer -Identity $localServerName -ErrorAction SilentlyContinue
   
    if($Error)
    {
        #something went wrong getting server name
        Write-Host -ForegroundColor Red "Failed"
        exit
    }
    
    $localVersionMajor = $localServerInfo.AdminDisplayVersion.Major
    $localVersionMinor = $localServerInfo.AdminDisplayVersion.Minor

    if($localVersionMajor -eq 8)
    {
        # "2007 detected and is not supported with this script"
        exit
    }
    elseif($localVersionMajor -eq 14)
    {
        # "2010 detected"
        $script:Exchange2010 = $true
    }
    elseif(($localVersionMajor -eq 15) -and ($localVersionMinor -eq 0))
    {
        # "2013 detected"
        $script:Exchange2013 = $true
    }
    elseif(($localVersionMajor -eq 15) -and ($localVersionMinor -eq 1))
    {
        # "2016 detected"
        $script:Exchange2016 = $true
    }
    else
    {
        # "something went wrong and couldn't detect local server version"
        Write-Host -ForegroundColor Red "Failed"
        exit
    }
    Write-Host -ForegroundColor Green "Success"
    
    return $localServerInfo.Site.Name
}

function GetClientAccessServers
{
    $error.Clear()

    if($Exchange2010 -or $Exchange2013)
    {
        $CASServers = Get-ClientAccessServer
    }
    if($Exchange2016)
    {
        $CASServers = Get-ClientAccessService
    }
    if($error)
    {
        Write-Host -ForegroundColor Red "Failed"
        exit
    } else {
        Write-Host -ForegroundColor Green "Success"
        return $CASServers
    }
}

function GetRemoteServerInfo($CASServers)
{
    $error.Clear()
    $ServerArray = @()
    foreach($CASServer in $CASServers)
    {
        $obj = $null
        
        $ServerInfo = Get-ExchangeServer -Identity $CASServer.Name -ErrorAction SilentlyContinue
        
        if($ServerInfo.AdminDisplayVersion.Major -eq 14)
        {
            $obj = New-Object System.Object
            $obj | Add-Member -type NoteProperty -name Name -value $ServerInfo.Name
            $obj | Add-Member -type NoteProperty -name Site -value $ServerInfo.Site.Name
            $obj | Add-Member -type NoteProperty -name Version -value "Exchange 2010"

        }elseif(($ServerInfo.AdminDisplayVersion.Major -eq 15) -and ($ServerInfo.AdminDisplayVersion.Minor -eq 0))
        {
            #we will grab mailbox roles later
            $2013exists = $true
        }elseif(($ServerInfo.AdminDisplayVersion.Major -eq 15) -and ($ServerInfo.AdminDisplayVersion.Minor -eq 1))
        {
            $obj = New-Object System.Object
            $obj | Add-Member -type NoteProperty -name Name -value $ServerInfo.Name
            $obj | Add-Member -type NoteProperty -name Site -value $ServerInfo.Site.Name
            $obj | Add-Member -type NoteProperty -name Version -value "Exchange 2016"

        }else
        {
            #not a supported version
        }

        if($obj)
        {

            $ServerArray += $obj

        }
    }

    if($error)
    {
        Write-Host -ForegroundColor Red "Failed"
        exit
    } else {
        Write-Host -ForegroundColor Green "Success"
    }

    if($2013exists)
    {
        #if 2013 servers exist, we have to consider the possibility of split roles.  Mailbox role is what needs to be reset.
        $error.clear()
        Write-Host "Exchange 2013 detected, so grabbing Mailbox server roles..." -NoNewline
        $MBXServers = Get-MailboxServer

        foreach($MBXServer in $MBXServers)
        {
            if(($MBXServer.AdminDisplayVersion.Major -eq 15) -and ($MBXServer.AdminDisplayVersion.Minor -eq 0))
            {
                 $MBXServerInfo = Get-ExchangeServer -Identity $MBXServer.Name -ErrorAction SilentlyContinue
                 $obj2 = New-Object System.Object
                 $obj2 | Add-Member -type NoteProperty -name Name -Value $MBXServerInfo.Name
                 $obj2 | Add-Member -type NoteProperty -Name Site -Value $MBXServerInfo.Site.Name
                 $obj2 | Add-Member -type NoteProperty -Name Version -Value "Exchange 2013"
                 $ServerArray += $obj2   
            }
        }
        
        if($error)
        {
            Write-Host -ForegroundColor Red "Failed"
            exit
        } else {

            Write-Host -ForegroundColor Green "Success"
        }

    }

    return $ServerArray
}

function ResetWSSecurityAll($CASServers)
{
    Write-Host "Resetting WSSecurity on ALL servers"

    foreach($CASServer in $CASServers)
    {
        $Error.Clear()
        write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
        if($CASServer.Version -eq "Exchange 2010")
        {
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
        }else{
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
        }
        if($Error)
        {
        Write-Host -ForegroundColor Red "Failed"
        } else {
        Write-Host -ForegroundColor Green "Success"
        }

        $Error.Clear()
        Write-Host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
        if($CASServer.Version -eq "Exchange 2010")
        {
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
        }else{
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
        }
        if($Error)
        {
            Write-Host -ForegroundColor Red "Failed"
        } else {
            Write-Host -ForegroundColor Green "Success"
        }
    }
}

function ResetWSSecurityLocalSite($localSite,$CASServers)
{
    write-host "Resetting WSSecurity on ALL servers in site $localSite"

    foreach($CASServer in $CASServers)
    {
        if($CASServer.Site -eq $localSite)
        {
            $Error.Clear()
            write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
            if($CASServer.Version -eq "Exchange 2010")
            {
                Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
            }else{
                Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
            }
            if($Error)
            {
            Write-Host -ForegroundColor Red "Failed"
            } else {
            Write-Host -ForegroundColor Green "Success"
            }

            $Error.Clear()
            Write-Host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
            if($CASServer.Version -eq "Exchange 2010")
            {
                Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
            }else{
                Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
            }
            if($Error)
            {
                Write-Host -ForegroundColor Red "Failed"
            } else {
                Write-Host -ForegroundColor Green "Success"
            }
        }
    }
}

function ResetWSSecurityVersion($CASServers)
{
    $e14 = $false
    $e15 = $false
    $e16 = $false

    foreach($CASServer in $CASServers)
    {
        if($CASServer.Version -eq "Exchange 2010")
        {
            $e14 = $true
        }
        if($CASServer.Version -eq "Exchange 2013")
        {
            $e15 = $true
        }
        if($CASServer.Version -eq "Exchange 2016")
        {
            $e16 = $true
        }
    }
    if($e14)
    {
        Write-Host "2010| Exchange 2010"
    }
    if($e15)
    {
        Write-Host "2013| Exchange 2013"
    }
    if($e16)
    {
        Write-Host "2016| Exchange 2016"
    }
    Write-Host ""
    $selection = Read-Host "Please enter the version of Exchange to reset WSSecurity (2010/2013/2016)"
    if($selection -eq "2010")
    {
        if($e14 -ne $true)
        {
            Write-Host -ForegroundColor Red "You have no Exchange 2010 servers. Please run the script again."
            quit
        }
        else
        {
            foreach($CASServer in $CASServers)
            {
                $Error.Clear()
                if($CASServer.Version -eq "Exchange 2010")
                {
                    write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2010"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2010") {
                    Write-Host -ForegroundColor Green "Success"
                }

                $error.Clear()
                if($CASServer.Version -eq "Exchange 2010")
                {
                    write-host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2010"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2010") {
                    Write-Host -ForegroundColor Green "Success"
                }
            }
        }
    }
    elseif($selection -eq "2013")
    {
        if($e15 -ne $true)
        {
            Write-Host -ForegroundColor Red "You have no Exchange 2013 servers. Please run the script again."
            quit
        }
        else
        {
            foreach($CASServer in $CASServers)
            {
                $Error.Clear()
                if($CASServer.Version -eq "Exchange 2013")
                {
                    write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2013"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2013")
                {
                    Write-Host -ForegroundColor Green "Success"
                }

                $error.Clear()
                if($CASServer.Version -eq "Exchange 2013")
                {
                    write-host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2013"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2013")
                {
                    Write-Host -ForegroundColor Green "Success"
                }
            }
        }
    }
    elseif($selection -eq "2016")
    {
        if($e16 -ne $true)
        {
            Write-Host -ForegroundColor Red "You have no Exchange 2016 servers. Please run the script again."
            quit
        }
        else
        {
            foreach($CASServer in $CASServers)
            {
                $Error.Clear()
                if($CASServer.Version -eq "Exchange 2016")
                {
                    write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2016"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2016")
                {
                    Write-Host -ForegroundColor Green "Success"
                }

                $error.Clear()
                if($CASServer.Version -eq "Exchange 2016")
                {
                    write-host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2016"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2016")
                {
                    Write-Host -ForegroundColor Green "Success"
                }
            }
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "You have entered an invalid version of Exchange. Please run the script again"
    }
}

function ResetWSSecurityVersionSite($CASServers, $localSite)
{
    $e14 = $false
    $e15 = $false
    $e16 = $false

    foreach($CASServer in $CASServers)
    {
        if($CASServer.Version -eq "Exchange 2010")
        {
            $e14 = $true
        }
        if($CASServer.Version -eq "Exchange 2013")
        {
            $e15 = $true
        }
        if($CASServer.Version -eq "Exchange 2016")
        {
            $e16 = $true
        }
    }
    if($e14)
    {
        Write-Host "2010| Exchange 2010"
    }
    if($e15)
    {
        Write-Host "2013| Exchange 2013"
    }
    if($e16)
    {
        Write-Host "2016| Exchange 2016"
    }
    Write-Host ""
    $selection = Read-Host "Please enter the version of Exchange to reset WSSecurity in site $localSite (2010/2013/2016)"
    if($selection -eq "2010")
    {
        if($e14 -ne $true)
        {
            Write-Host -ForegroundColor Red "You have no Exchange 2010 servers. Please run the script again."
            quit
        }
        else
        {
            foreach($CASServer in $CASServers)
            {
                $Error.Clear()
                if(($CASServer.Version -eq "Exchange 2010") -and ($CASServer.Site -eq $localSite))
                {
                    write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2010"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2010") {
                    Write-Host -ForegroundColor Green "Success"
                }

                $error.Clear()
                if(($CASServer.Version -eq "Exchange 2010") -and ($CASServer.Site -eq $localSite))
                {
                    write-host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2010"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2010") {
                    Write-Host -ForegroundColor Green "Success"
                }
            }
        }
    }
    elseif($selection -eq "2013")
    {
        if($e15 -ne $true)
        {
            Write-Host -ForegroundColor Red "You have no Exchange 2013 servers. Please run the script again."
            quit
        }
        else
        {
            foreach($CASServer in $CASServers)
            {
                $Error.Clear()
                if(($CASServer.Version -eq "Exchange 2013") -and ($CASServer.Site -eq $localSite))
                {
                    write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2013"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2013")
                {
                    Write-Host -ForegroundColor Green "Success"
                }

                $error.Clear()
                if(($CASServer.Version -eq "Exchange 2013") -and ($CASServer.Site -eq $localSite))
                {
                    write-host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2013"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2013")
                {
                    Write-Host -ForegroundColor Green "Success"
                }
            }
        }
    }
    elseif($selection -eq "2016")
    {
        if($e16 -ne $true)
        {
            Write-Host -ForegroundColor Red "You have no Exchange 2016 servers. Please run the script again."
            quit
        }
        else
        {
            foreach($CASServer in $CASServers)
            {
                $Error.Clear()
                if(($CASServer.Version -eq "Exchange 2016") -and ($CASServer.Site -eq $localSite))
                {
                    write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2016"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2016")
                {
                    Write-Host -ForegroundColor Green "Success"
                }

                $error.Clear()
                if(($CASServer.Version -eq "Exchange 2016") -and ($CASServer.Site -eq $localSite))
                {
                    write-host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
                    Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue
                }
                if(($Error) -and ($CASServer.Version -eq "Exchange 2016"))
                {
                    Write-Host -ForegroundColor Red "Failed"
                } elseif($CASServer.Version -eq "Exchange 2016")
                {
                    Write-Host -ForegroundColor Green "Success"
                }
            }
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "You have entered an invalid version of Exchange. Please run the script again"
    }
}

Function ResetWSSecurityServerName($CASServers)
{
    $server = Read-Host "Enter the short name (not FQDN) of the server you'd like to reset WSSecurity"
    $match = $false

    foreach($CASServer in $CASServers)
    {
        if(($CASServer.Name -eq $server) -and ($CASServer.Version -eq "Exchange 2010"))
        {
            $Error.Clear()
            write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue

            if($Error)
            {
                write-Host -ForegroundColor Red "Failed"
            } else {
                write-Host -ForegroundColor Green "Success"
            }   

            $Error.Clear()
            write-Host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Default Web Site)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue

            if($Error)
            {
                write-Host -ForegroundColor Red "Failed"
            } else {
                write-Host -ForegroundColor Green "Success"
            } 

            $match = $true

        }elseif(($CASServer.Name -eq $server) -and ($CASServer.Version -eq "Exchange 2013"))
        {
            $Error.Clear()
            write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue

            if($Error)
            {
                write-Host -ForegroundColor Red "Failed"
            } else {
                write-Host -ForegroundColor Green "Success"
            }   

            $Error.Clear()
            write-Host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue

            if($Error)
            {
                write-Host -ForegroundColor Red "Failed"
            } else {
                write-Host -ForegroundColor Green "Success"
            }    
            $match = $true

        }elseif(($CASServer.Name -eq $server) -and ($CASServer.Version -eq "Exchange 2016"))
        {

            $Error.Clear()
            write-host -NoNewline "Resetting WSSecurity for EWS on $($CASServer.Name)..."
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-WebServicesVirtualDirectory -Identity "$($CASServer.Name)\ews (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue

            if($Error)
            {
                write-Host -ForegroundColor Red "Failed"
            } else {
                write-Host -ForegroundColor Green "Success"
            }   

            $Error.Clear()
            write-Host -NoNewline "Resetting WSSecurity for Autodiscover on $($CASServer.Name)..."
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$False -ErrorAction SilentlyContinue
            Set-AutoDiscoverVirtualDirectory -Identity "$($CASServer.Name)\autodiscover (Exchange Back End)" -WSSecurityAuthentication:$True -ErrorAction SilentlyContinue

            if($Error)
            {
                write-Host -ForegroundColor Red "Failed"
            } else {
                write-Host -ForegroundColor Green "Success"
            }    
            $match = $true
        }
    }

    if(!$match)
    {
        Write-Host -ForegroundColor Red "You have entered an invalid server name. Please run the script again"
    }
}

Function ConfirmAnswerYN
{
	$Confirm = "" 
	while ($Confirm -eq "") 
	{ 
		switch (Read-Host "(Y/N)") 
		{ 
			"yes" {$Confirm = "yes"} 
			"no" {$Confirm = "No"} 
			"y" {$Confirm = "yes"} 
			"n" {$Confirm = "No"} 
			default {Write-Host "Invalid entry, please answer question again " -NoNewline} 
		} 
	} 
	return $Confirm 
}

Function ConfirmAnswerNumerical($min,$max)
{
    $answer = 0
    while (($answer -lt $min) -or ($answer -gt $max))
    {
        $answer = Read-Host "Enter your selection ($min-$max)"
    }

    return $answer

}

#Main script function
#====================================================
$Error.Clear()
Write-Host -NoNewline "Adding the Exchange Management Snapin..."
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Snapin -ErrorAction SilentlyContinue
if($error)
{
    #try 2010
    $error.Clear()
    Add-PSSnapin Microsoft.Exchange.Management.Powershell.E2010 -ErrorAction SilentlyContinue
    if(!$error)
    {
        $runningFrom2010 = $true
    }
 }
 if($Error)
 {
    Write-Host -ForegroundColor Red "Failed"
    exit
} else {
    Write-Host -ForegroundColor Green "Success"
}

if($runningFrom2010)
{

    write-host -ForegroundColor Red "If you have any 2013/2016 servers, you cannot run this script from Exchange 2010."
    write-host "Do you have any Exchange 2013/2016 servers in your environment?" -NoNewline
    $answer = ConfirmAnswerYN
    if ($answer -eq "yes")
	{
		exit
	}
}


Write-Host -NoNewline "Verifying local server..."
$localSite = GetLocalServerInfo
Write-Host -NoNewline "Getting list of organization CAS Servers..."
$CASServers = GetClientAccessServers
Write-Host -NoNewline "Getting each CAS Server's information..."
$ServerPool = GetRemoteServerInfo($CASServers)

#### there is likely a much better way to output this info

write-Host "The following supported servers have been discovered:"
write-Host ""

foreach($object in $ServerPool)
{
    write-Host "$($object.Name)`t$($object.Site)`t$($object.Version)"
}
write-Host ""
write-Host "Which servers would you like to reset WSSecurity on?"
write-Host ""
write-Host "1) ALL servers" -NoNewline
Write-Host -ForegroundColor Yellow " (not recommended if you have multiple geographical sites)*"
write-Host "2) ALL servers in the $localSite site"
write-Host "3) All servers of a specific version" -NoNewline
Write-Host -ForegroundColor Yellow " (not recommended if you have multiple geographic sites)*"
write-host "4) All specific version servers in the $localSite site"
write-host "5) A specific server"
write-host ""
write-host -ForegroundColor Yellow "* If you have multiple geographical sites, applying this change across site boundries may take a very long time"
Write-Host -ForegroundColor Yellow "  It is recommended, instead, that you run this script from each site"
write-host ""
$selection = ConfirmAnswerNumerical 1 5

switch($selection)
{

    1 {ResetWSSecurityAll $ServerPool}
    2 {ResetWSSecurityLocalSite $localSite $ServerPool}
    3 {ResetWSSecurityVersion $ServerPool}
    4 {ResetWSSecurityVersionSite $ServerPool $localSite}
    5 {ResetWSSecurityServerName $ServerPool}

}



