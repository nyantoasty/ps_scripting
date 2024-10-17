



   <#
    .DESCRIPTION
        This script is designed to quickly download a specific driver, and map to two specific printers.

        These can be adjusted in the variables at the top as needed.

    .FUNCTIONS
        * add-COMPrinters
            This function will loop through $printerInfo and add whatever printers/ports are listed
        * check-successful
            This function checks to see if the desired printers have been added and configured according to $printerInfo and the desired $printerINF

    .NOTES
        Created by: KJA
        Modified: 2024-09-06

    .CHANGELOG
        * 9.6.24 - created script       

#>

#Modules ===================================================================================================================
if(!(get-installedmodule -name 'PSWRiteColor' -EA ignore)) {
    set-psrepository -name psgallery -InstallationPolicy Trusted
    install-module -name pswritecolor
    import-module -name pswritecolor 
}
if(!(get-installedmodule -name '7zip4powershell' -ea ignore)){
    set-psrepository -name psgallery -installationpolicy trusted
    install-module -name 7zip4powershell
    import-module -name 7zip4powershell
}
#Variables ===================================================================================================================

$script:printerInfo = @{
    "COMM110C-1" = "https://ipp.shsu.edu/printers/COMM110C-1"
    "COM408-2" = "https://ipp.shsu.edu/printers/COM408-2"
} #printer name + address
$script:folderName = "PrinterInstalls"
$script:printerFolder = "C:\" + "$script:foldername" + "\" #where to save .inf and reinstall script
$script:printerINF = "CNLB0MA64.inf" #current .inf file obtained from running driver installer (Canon UFR Plus Driver, located in Drivers folder after install)

$script:printModelsDriers = @{

}

$script:url = "https://downloads.canon.com/bicg2024/drivers/Generic_Plus_UFRII_v3.00_Set-up.exe"

$script:allPrinters = @()
$script:reportPrinters = @()


$printerObj = [pscustomobject]@{
    Name = $printer.printername
    Driver = $printer.drivername
    Port = $printer.portname
    uNCName = $printer.uncname
    url = $printer.url
    Location = $printer.location
    Description = $printer.description
    Created = $printer.whencreated
    Modified = $printer.whenchanged
}


$drivers = @{
    "Canon Generic Plus UFR II" = "https://downloads.canon.com/bicg2024/drivers/Generic_Plus_UFRII_v3.00_Set-up.exe"
    "HP Universal Printing PCL 6" = "https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-pcl6-x64-7.2.0.25780.exe" 
    "HP Universal Printing PS" = "https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-ps-x64-7.2.0.25780.exe"
    "Kyocera Classic Universaldriver PCL6" = "https://www.kyoceradocumentsolutions.us/content/download-center-americas/us/drivers/drivers/P3145_P3150_P3155_PCL_zip.download.zip"
    "RICOH PCL6 UniversalDriver V4.38" = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001340/0001340025/V43800/z00960L1b.exe"
    "Xerox Global Print Driver PCL6" = "https://www.support.xerox.com/en-us/product/global-printer-driver/downloads?language=en#"
    Unlisted = ''
}
#Functions ===================================================================================================================

function check-printDriver {
    param ([string]$pname)

    $printer = get-adobject -filter {printername -like $pname} -properties *
    write-host $printer

    $printerObj = [pscustomobject]@{
        Name = $printer.printername
        Driver = $printer.drivername
        Port = $printer.portname
        uNCName = $printer.uncname
        url = $printer.url
        Location = $printer.location
        Description = $printer.description
        Created = $printer.whencreated
        Modified = $printer.whenchanged
    }
    $check = get-printerdriver
    $res = $check | where-object Name -like $printerobj.driver
    if($res -eq $null) {
        Write-Color "Warning!"," ${$printerobj.driver} not found. Attempting to download." -color red,yellow

        if($printerobj.driver -in $drivers.keys) {
            write-color "Downloading driver from ${$printerobj.url}"
            New-Item -path "C:\" -name $script:foldername -itemtype "directory"
            $clnt = new-object system.net.webclient
            $file = $script.printerfolder + "driver.exe"
            $dest = $script:printerfolder
            $clnt.downloadfile($printerobj.url,$file)

        }
    }
    else {
        Write-Color "${$printerobj.driver} already installed."
    }
}

########WIP
 add-printerport -name $printerobj.port -printerhostaddress $printerobj.uncname
 
 add-printer -drivername $printerobj.driver -name $printerobj.name -portname $printerobj.uncname

function add-COMPrinters {
    param(
        [string]$pname,
        [string]$paddress        
    )
    $dname = "Canon Generic Plus UFR II"
    $check = get-printerdriver
    $res = $check | where-object Name -like $dname

    if($res -eq $null) {
    
        New-Item -path "C:\" -Name $script:foldername -itemtype "directory"
        $clnt = new-object system.net.webclient
        $file = $script:printerfolder + "driver.exe"
        $dest = $script:printerfolder

        $clnt.downloadfile($script:url,$file)

        expand-7zip -archiveFilename $file -TargetPath $dest

        $path = $script:printerfolder + "x64\Driver\"
        $path = $path + $printerINF

        Invoke-Command {C:\Windows\System32\pnputil.exe -a $path } -erroraction silentlycontinue
        Add-PrinterDriver -name $dname
        
        $check = get-printerdriver
        $res = $check | where-object Name -like $dname
    }

    write-color "Driver"," $($dname) ","(Major version"," $($res.majorversion))"," found." -color blue,green,blue,green,blue

    Add-PrinterPort -name $paddress -PrinterHostAddress $paddress -erroraction silentlycontinue
    add-printer -drivername $dname -name $pname -portname $paddress -erroraction silentlycontinue

    $tmpObj = [pscustomobject]@{
        Printer = $pname
        Port = $paddress
        Driver = $dname
    }
    $script:allPrinters += $tmpObj
}

function check-successful {
    param(
        [array]$desired
    )
    $flagArray = @{
        prntflag = $null
        drvflag = $null
        portflag = $null
    }

    $present = get-printer
    foreach($printer in $desired) {
        $check = $present | where-object name -eq $printer.Printer
        if($check -eq $null) {
            write-color "Printer ","$($printer.printer)"," not present; try adding manually."
            $flagArray.prntflag = $false
            $flagArray.drvflag = $false
            $flagArray.portflag = $false            
        }
        else {
            $flagArray.prntflag = $true
            if($check.drivername -ne $printer.driver) {
                write-color "Driver mismatch!" -color red
                Write-color "Expected: ","$($printer.driver)" -color yellow, red
                Write-color "Detected: ","$($check.drivername)" -color yellow,red
                $flagArray.drvflag = $false
            }
            else {
                $flagArray.drvflag = $true
            }
            if($check.portname -ne $printer.port) {
                write-color "Port mismatch!" -color red
                Write-color "Expected: ","$($printer.port)" -color yellow, red
                Write-color "Detected: ","$($check.portname)" -color yellow,red
                $flagArray.portflag = $false
            }
            else {
                $flagArray.portflag = $true
            }
        }
        $printername = $printer.printer
        $printername = -join($printername,", $($flagarray.prntflag)")
        $driver = $printer.driver
        $driver = -join($driver, ", $($flagarray.drvflag)")
        $port = $printer.port
        $port = -join($port,", $($flagarray.portflag)")

        $report = [pscustomobject]@{
            Printer = $printername
            Driver = $driver
            Port = $port
        }
        $script:reportPrinters += $report
    }
}

#Main program =================================================================================================================

foreach($printer in $printerinfo.keys) {

    add-comprinters $printer $printerinfo[$printer]
}

check-successful $script:allprinters

remove-item $script:printerfolder -recurse

$script:reportPrinters | fl
