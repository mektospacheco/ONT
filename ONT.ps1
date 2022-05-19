
# +-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
# |                                                                                                                      Script Info                                                                                                                      |
# +-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
# | This powershell script will perform a range of connectivity tests to Office 365 to ensure that TCP connections via port 80 and 443 are happening without any issues. The tool achieves this by downloading PSPing and running it from the machine.    |
# | This script will also check the machine settings for proxy servers and at the end provide articles that can be helpfull for further understanding of recommendations and best practices related to networking configurations in Office 365            |
# | Created by Ricardo Pacheco (ricardo.pacheco@microsoft.com) - This script is provided as is and it is not guaranteed to work 100% of the time. Please review all the code before running it to avoid any code execution that you are note confortable. |
# +-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+





#Collecting Desktop Path
    $desktop = [System.Environment]::GetFolderPath('Desktop')
 #Setting Output directory
    $output = $desktop + '\OutlookNetworkTests'

#Test Path
    $PathExists = Test-Path -Path $output


#Creating Path to store files / logs
    If ($PathExists -eq $False) {
        New-Item -Path $desktop -Name "OutlookNetworkTests" -ItemType directory
    }
 #Starting Transcript
    Start-Transcript -OutputDirectory $output
        
#Run script as Administrator uncomment if needed
#if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Load assembly for GUI messages
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    cls
    #Display Tool descriptionGUI
    Write-Host "This is the Outlook Network Test Tool, this tool will perform a range of network tests to Office 365`n`nThis tool performs a range of TCP and connections to Office 365 trough port 80 and 443 to check the connectivity status.`n`nA log file will be generated at the end." -ForegroundColor Yellow





#Office 365 url variables
    $Office365SSL = "outlook.office365.com:443"
    $Office365Http = "outlook.office365.com:80"
    $O365 = "outlook.office365.com"

#Files to be downloaded
    $EndpointsScript = "https://gallery.technet.microsoft.com/office/Get-Office-365-Endpoint-b533554a/file/216426/1/Invoke-O365EndpointDataGathering.ps1"
    $PSPingurl = "https://live.sysinternals.com/tools/psping.exe"
    $ValidateO365Script = "https://gallery.technet.microsoft.com/Validate-Office-365-3517ae88/file/220645/1/Test-O365%20V1.2.zip"

#Import BITS Module
  Import-Module BitsTransfer

     
     
  #Changing to work directory
  cd $output

  #Asking for domain name being tested

     $rootdomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your company domain name.", "Company Domain Name")

     #Dump Date for logging purposes
        Get-Date

        cls

       #Checking If Proxy exists / is enabled
        Write-Host "Collecting Proxy settings from your system..." -ForegroundColor Yellow
         $IsProxyEnabled = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyEnable
            $ProxyServer = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyserver

    #Dumping Proxy info
        Write-Host "Is proxy enabled? 0 = No / 1 = Yes"
         $IsProxyEnabled
        Write-Host "Address Found (It doenst mean is active, just that we have it stored on the machine)"
         $ProxyServer

         #Proxy GUI Dump
       If ($IsProxyEnabled -eq 0) {
        Write-Host "The Proxy is DISABLED" -ForegroundColor Green
            } 
    Else
        {
            Write-Host "The Proxy is ENABLED and we found the following address $ProxyServer" -ForegroundColor Red
        } 

  #Download PSPing
  Write-Host " "
  Write-Host " "
    Write-Host "Downloading PSPing" -ForegroundColor Yellow

#Clear pending BitsTransfer Jobs
Get-BitsTransfer | Remove-BitsTransfer

    #Starting BITS Download of PSPING
        Start-BitsTransfer -Source $PSPingurl -Destination $output

    
  <#
  
  Download O365 Validation Script
    Write-Host "Downloading Office 365 IP's Validation script" -ForegroundColor Yellow

     #Starting BITS Download of O365 Validation Script
        Start-BitsTransfer -Source $ValidateO365Script -Destination $output

            #Unziping file
               Expand-Archive -Path Test-O365%20V1.2.zip -DestinationPath C:\OutlookNetworkTest

                #Removing unnacessary file
                    Remove-Item $output\Test-O365%20V1.2.zip
                    
                    #>





    #TCP Testing to Outlook.office365.com w/ PSPing
    Write-Host " "
             Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
        Write-Host "Testing TCP Connections" -ForegroundColor Yellow
        Write-Host " "

        #IPv4 Testing Port 443 
        Write-Host " "
             Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
             Write-Host "Port 443 IPv4" -ForegroundColor Yellow
                 .\psping.exe -4 -q -n 25 -i 0 $Office365SSL

             
             #IPv4 Testing Port 80
             Write-Host " "
             Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
                 Write-Host "Port 80 IPv4" -ForegroundColor Yellow
                    .\psping.exe -4  -n 25 -i 0 $Office365Http


         #IPv6 Testing Port 443 
          Write-Host " "
                Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
            Write-Host "Port 443 IPv6" -ForegroundColor Yellow
                .\psping.exe -6  -n 25 -i 0 $Office365SSL
                            
             #IPv6 Testing Port 80
                Write-Host " "
                Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
                Write-Host "Port 80 IPv6" -ForegroundColor Yellow
                    .\psping.exe -6  -n 25 -i 0 $Office365Http

    #ICMP Testing
    Write-Host " " 
    Write-Host " "
    Write-Host "Testing ICMP Connections" -ForegroundColor Yellow
    Write-Host " "
        
        #ICMP IPv4 Testing to Outlook.office365.com w/ PSPing
        Write-Host " " 
        Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
            Write-Host "ICMP Office365 IPv4" -ForegroundColor Yellow
                 .\psping.exe -4  -n 25 -i 0 $O365

        #ICMP IPv4 Testing to autodiscover domain w/ PSPing
        Write-Host " " 
        Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
            Write-Host "ICMP Autodiscover IPv4" -ForegroundColor Yellow
                .\psping.exe -4  -n 25 -i 0 autodiscover.$rootdomain

        #ICMP IPv6 Testing to Outlook.office365.com w/ PSPing
        Write-Host " " 
        Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
        Write-Host "ICMP Office365 IPv6" -ForegroundColor Yellow
             .\psping.exe -6  -n 25 -i 0 $O365

        #ICMP IPv6 Testing to autodiscover domain w/ PSPing
        Write-Host " " 
        Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" 
        Write-Host "ICMP Autodiscover IPv6" -ForegroundColor Yellow
             .\psping.exe -4  -n 25 -i 0 autodiscover.$rootdomain

             

  <#
  
  Need to add check if we want to  download ip's and urls for later check if yes
  Start-BitsTransfer -Source $EndpointsScript -Destination $output
  .\Invoke-O365EndpointDataGathering.ps1
   Write-Host " " 
    Write-Host " "
   Write-Host "Please check the if the IP's and url's downloaded " -ForegroundColor Yellow

  #>

   

   $a = new-object -comobject wscript.shell 
$intAnswer = $a.popup("Do you wish to open articles with relevant information related to Office 365 / Exchange Online / Outlook Networking?", ` 
0,"Open Support Articles?",4) 
If ($intAnswer -eq 6) { 
   start https://blogs.technet.microsoft.com/exovoice/2016/10/24/basic-troubleshooting-of-outlook-connectivity-in-office-365-from-network-perspective/
   start https://docs.microsoft.com/en-us/office365/enterprise/office-365-network-connectivity-principles
   start https://blogs.technet.microsoft.com/onthewire/2017/03/22/__guidance/
   start https://docs.microsoft.com/en-us/office365/enterprise/setup-overview-for-enterprises
   start https://docs.microsoft.com/pt-pt/office365/enterprise/network-planning-and-performance
} 

   [System.Windows.Forms.Messagebox]::Show("Please collect the log file in $output `n`nHave a nice day", "Finished")

  Stop-Transcript
 
 
  
  
