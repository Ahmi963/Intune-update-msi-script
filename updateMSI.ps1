<#
    .DESCRIPTION
    uninstall older versions of applications that uses MSI installer and install the newer version directly.
    this script is intended and tested for use in Endpoint Manager but could probably be used in similar client management systems 

  
    .NOTES
    filename:               updateMSI.ps1
    author:                 Ahmi963
    created:                22/04/2022
    last updated:           see Git
    #Copyright/License:     free to use as long as header are kept intact
#>


#######################################################################
# replace the variable values with the correct values for your use-case
#######################################################################
# to check application name run following command on a system where application is installed and copy the name from the list > get-CimInstance -ClassName Win32_Product | Format-Table IdentifyingNumber, Name, LocalPackage -AutoSize
$Global:appname = "replace with application name"               # example: "Adobe Acrobat DC" (or any other app that needs to be uninstalled)
$Global:msifilename = "replace with new msi package name"       # example: AcroRead.msi (or any other MSI file that needs to be installed) 





# uninstall the specified old version of an application
function UninstallOldVersion {
    try {
        $identifyingnumber = Get-CimInstance -ClassName Win32_Product | Where-Object {$_.Name -like "*$Global:appname*"} | Select-Object -ExpandProperty IdentifyingNumber
        if ($null -ne $identifyingnumber) 
        {
            Write-Output "application detected. uninstalling..."

            # installation parameters might be different for specific msi files
            msiexec /x $identifyingnumber /qn

            #will pause script for 15 seconds to insure uninstall is fully completed
            Start-Sleep -s 15
            return $true
            
        } else {
            Write-Output " old application was not detected. installation can begin"
            return $true
        }
    }
    catch {
        Write-Host "Something went wrong! application presence couldn't be detected $($_)"
        return $false
        
    }
}

# install the newer version of the specified files
function InstallNewVersion {
   try {
    msiexec /i "$Global:msifilename" /qn
    return $true
   } 
   catch {
    Write-Host "something went wrong while installing the application $($_)"
    return $false
   }
}

#MAIN
try {
  if (UninstallOldVersion) 
  {
    if (InstallNewVersion)
    {
      Write-Output "deployment was successful"
      exit 0   
    } else {
        Write-Output "newer version couldn't be installed"
        exit 1
    }

  } else {
    Write-Output "older version couldn't be uninstalled"
    exit 1
  }
}
catch {
    Write-Host "deployment encountered following error and was not successful $($_)"
}