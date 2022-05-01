# ONT
Outlook Network Tests

This powershell script will perform a range of connectivity tests to Office 365 to ensure that TCP connections via port 80 and 443 are happening without any issues. The tool achieves this by downloading PSPing and running it from the machine.    
This script will also check the machine settings for proxy servers and at the end provide articles that can be helpfull for further understanding of recommendations and best practices related to networking configurations in Office 365            

The tests for custom domains is always performed against autodiscover.<domain.com>

Created by Ricardo Pacheco (ricardo.pacheco@microsoft.com) - This script is provided as is and it is not guaranteed to work 100% of the time. Please review all the code before running it to avoid any code execution that you are note confortable. 

The script can be easily accessed / run by running the following powershell cmdlets:

$ONTScript = Invoke-WebRequest https://aka.ms/ONT

Invoke-Expression ($ONTScript.Content)


                                                          Issues Found
                                                            

Issue 1

![Issue 1](https://raw.githubusercontent.com/mektospacheco/ONT/master/Images/image.png)


If you find the following error message, you can work around it by downloading the file directly from GitHub and executing the file.
![Issue 1](https://raw.githubusercontent.com/mektospacheco/ONT/master/Images/SaveAs.jpg)
