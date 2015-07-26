# Source: http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
function Get-MSIInformation
{
    Param
    (
        [parameter(Mandatory=$true)][IO.FileInfo]$Path,
        [parameter(Mandatory=$true)][ValidateSet("ProductCode","ProductVersion","ProductName")][string]$Property
    )

    try 
    {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        
        $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase","InvokeMethod",$Null,$WindowsInstaller,@($Path.FullName,0))
        
        $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
        
        $View = $MSIDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$null,$MSIDatabase,($Query))
        
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        
        $Record = $View.GetType().InvokeMember("Fetch","InvokeMethod",$null,$View,$null)
        
        $Value = $Record.GetType().InvokeMember("StringData","GetProperty",$null,$Record,1)
        
        return $Value
    } 
    catch
    {
        Write-Output $_.Exception.Message
    }
}

function Get-InstallationStatus
{
    [CmdletBinding()]
	Param
    (
        [Parameter(Mandatory=$true)]$ProductCode
    )

    if (Test-Path HKLM:\SOFTWARE\Wow6432Node)
    {
        [System.Object]$Product = Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where { $_.PSChildName -eq $ProductCode }
    }

    if (!($Product))
    {
        [System.Object]$Product = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where { $_.PSChildName -eq $ProductCode }
    }
    
    # Depricated by KB974524
    # https://support.microsoft.com/en-us/kb/974524
    #[System.Object]$Product = Get-WmiObject -Class Win32_Product -Filter "IdentifyingNumber='$ProductCode'"

    return $Product
}

function Install-MSI
{
    [CmdletBinding()]
	Param
    (
        [Parameter(Mandatory=$true)][ValidateScript({Test-Path $_ -PathType 'Container'})][string]$AIP,
        [Parameter(Mandatory=$true)][string]$Vendor,
        [Parameter(Mandatory=$true)][string]$Application,
        [Parameter(Mandatory=$true)][string]$Version,
        [Parameter(Mandatory=$false)][string]$Arguments
    )

    Process
    {
        $process = "msiexec.exe"

        Write-Verbose "Attempting to install $application $version from $vendor"

        Write-Verbose "Validating parameters"

        if (!(Test-Path ($AIP + $Vendor + "\" + $Application + "\" + $Version)))
        {
            #Log failure
            #Write-Log -Body "Unable to find path: $AIP$Vendor\$Application\$Version"

            exit
        }

        $installers = Get-ChildItem ($aip + $vendor + "\" + $application + "\" + $version + "\*") -Include *.msi

        Write-Verbose ("Found " + $installers.Count + " installers")

        $i = 1

        Write-Verbose "Begining itteration through installers"

        foreach ($installer in $installers)
        {
            Write-Progress -Activity "Installing $application" -CurrentOperation ("Installing " + $installer.Name) -Status "Current Status" -PercentComplete (($i/(($installers | measure).count)*100))

            Write-Verbose ("Installer itteration $i | " + $installer.Name)

            Write-Verbose "Testing if product is already installed"

            Write-Verbose "Retreiving Product Code"

            $ProductCode = (Get-MSIInformation -Path $installer -Property ProductCode)

            Write-Verbose "Product Code is: $ProductCode"

            if (!(Get-InstallationStatus -ProductCode $ProductCode[1]))
            {
                Write-Verbose "Product is not already installed, begining install"

                Write-Verbose ("Start-Process $process -ArgumentList '/i `"" + $installer.FullName + "`" $arguments -Verb runAs -Wait")

                Start-Process $process -ArgumentList ('/i "' + $installer.FullName + '" ' + $arguments) -Verb runAs -Wait

                if (Get-InstallationStatus -ProductCode $ProductCode[1])
                {
                    Write-Host -ForegroundColor Green ("Installation of " + $installer.Name + " Succeeded")
                }
                else
                {
                    #Log failure
                    #Write-Log ("Installation of " + $installer.Name + " Failed")
                }
            }
            else
            {
                Write-Verbose "Product is already installed, skipping"
            }

            Write-Verbose "Incrementing counter"

            $i++
        }
    }
}
