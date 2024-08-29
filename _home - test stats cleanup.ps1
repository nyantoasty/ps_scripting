
function write-excel {

    $preview = preview-obj

    if(test-path $script:checkpath) {

        Write-Color "${script:checkpath} already exists!" -color red
        $script:checkpath = "${Directory}\New_${Report}_${TaskName}.xlsx"
        Write-Color "Adjusting report name to '${script:checkpath}'" -color yellow
    }

    if ($preview -like 'Y') {
      $script:xlpkg = $script:trash       | export-excel -path $script:checkpath    -worksheetname 'TRASH'        -tablename 'T_T'              -autosize -passthru 
      $script:xlpkg = $script:Basics      | Export-excel -excelpackage $xlpkg       -worksheetname 'Basics'       -tablename 'Get_Basic'        -autosize -passthru 
      $script:xlpkg = $script:Battery     | Export-excel -excelpackage $xlpkg       -worksheetname 'Battery'      -tablename 'Get_Battery'      -autosize -passthru 
      $script:xlpkg = $script:OS          | Export-excel -excelpackage $xlpkg       -worksheetname 'OS'           -tablename 'Get_OS'           -autosize -passthru 
      $script:xlpkg = $script:BIOS        | Export-excel -excelpackage $xlpkg       -worksheetname 'BIOS'         -tablename 'Get_BIOS'         -autosize -passthru 
      $script:xlpkg = $script:CPU         | Export-excel -excelpackage $xlpkg       -worksheetname 'CPU'          -tablename 'Get_CPU'          -autosize -passthru 
      $script:xlpkg = $script:GPU         | Export-excel -excelpackage $xlpkg       -worksheetname 'GPU'          -tablename 'Get_GPU'          -autosize -passthru 
      $script:xlpkg = $script:RAM         | Export-excel -excelpackage $xlpkg       -worksheetname 'RAM'          -tablename 'Get_RAM'          -autosize -passthru 
      $script:xlpkg = $script:Network     | Export-excel -excelpackage $xlpkg       -worksheetname 'Network'      -tablename 'Get_Network'      -autosize -passthru 
      $script:xlpkg = $script:Space       | Export-excel -excelpackage $xlpkg       -worksheetname 'Space'        -tablename 'Get_Space'        -autosize -passthru 
      $script:xlpkg = $script:Health      | Export-excel -excelpackage $xlpkg       -worksheetname 'Health'       -tablename 'Get_Health'       -autosize -passthru -MoveToStart
      

      $script:xlpkg.workbook.worksheets.delete('TRASH')
      set-excelrange    -worksheet $script:xlpkg.Health -Range "E:E" -wraptext
      set-excelrange    -worksheet $script:xlpkg.Health -Range "G:G" -wraptext
      set-excelrange    -worksheet $script:xlpkg.Health -Range "H:H" -wraptext
      set-excelrange    -worksheet $script:xlpkg.OS -Range "H:H" -wraptext
      set-excelrange    -worksheet $script:xlpkg.OS -Range "Q:Q" -wraptext
      close-excelpackage $script:xlpkg
    }
    else {
      Write-color "Please try rerunning with a working object." -color red
    }
}

Function Get-FriendlySize {
    Param([bigint]$bytes)
    switch($bytes){
        {$_ -gt 1PB}{"{0:N2} PB" -f ($_ / 1PB);break}
        {$_ -gt 1TB}{"{0:N2} TB" -f ($_ / 1TB);break}
        {$_ -gt 1GB}{"{0:N2} GB" -f ($_ / 1GB);break}
        {$_ -gt 1MB}{"{0:N2} MB" -f ($_ / 1MB);break}
        {$_ -gt 1KB}{"{0:N2} KB" -f ($_ / 1KB);break}
        default {"{0:N2} Bytes" -f $_}
    }
}

function Get-Space {
  param(
  [string]$comp) 

  if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag

    $script:volumes = get-volume | where-object {$_.filesystemtype -ne 'Unknown' -and $_.driveletter -ne $null}

    foreach($vol in $volumes) {

        
        $drive = $vol.driveletter

        $infov = get-volume | where driveletter -like $drive | get-partition | select *
        $infops = get-psdrive $drive}
        $infodrive = get-disk | where disknumber -like $infov.disknumber | select *
        $infohealth = get-physicaldisk | where {$_.driveletter -like $infov.driveletter} | select *
        $free = (($infops.free / $infov.size)).tostring("P")
        $health = $infodrive.healthstatus

        $script:res = [pscustomobject]@{
          Room = $script:Room
          Host = $comp
          OT = $OT
          Query = 'Get-Space'
          Drive = $drive
          DriveModel = $infodrive.friendlyname
          DriveType = $vol.drivetype
          Health = $health
          BusType = $infodrive.bustype
          AdapterSerial = $vol.adapterserialnumber
          SerialNumber = $infodrive.serialnumber
          Signature = $infodrive.signature
          FileSystem = $vol.filesystem
          FileSystemLabel = $vol.filesystemlabel
          DiskNumber = $infodrive.disknumber
          NumberPartitions = $infodrive.numberofpartitions
          PartitionStyle = $infodrive.partitionstyle
          OperationalStatus = $infodrive.operationalstatus
          isBoot = $infodrive.isboot
          FirmwareVersion = $infodrive.firmwareversion
          Location = $infodrive.location
          Size = $infov.size
          Used = $infops.used
          Free = $infops.free
          Size_Friendly = get-friendlysize $infov.size
          Used_Friendly = get-friendlysize $infops.used
          Free_Friendly = get-friendlysize $infops.free
          '%_Free' = $free 
          Status = $status
        }
        $fsl = $vol.filesystemlabel
        $btype = $infodrive.bustype
        $friendly = $infodrive.friendlyname
        $name = "Label - "
        $name = -join($name,$fsl)
        $name = -join($name," `r`n")
        $name = -join($name,$btype)
        $name = -join($name," - ")
        $name = -join($name,$friendly)
        $tmp = get-friendlysize $infops.free
        $detail = $tmp
        $detail = -join($detail,' free of ')
        $tmp = get-friendlysize $infov.size
        $detail = -join($detail,$tmp)
        $detail = -join($detail,' total.')

        $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $comp
            OT = $OT
            Query = 'Get-Health'
            Component = 'Space'
            Health = $health
            Name = $name
            Detail = $detail
            Serial_Mac = $infodrive.serialnumber            
            
        }

        $script:health += $script:healthres
        $script:space += $res
        
    }
    return $script:space | out-null
}

function Get-Basic {
      param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
    $Servicetag = (get-ciminstance -computername $comp win32_bios).serialnumber
    
    $script:obj = get-computerinfo | select *

      $script:res = [pscustomobject]@{
            Room = $script:Room
            Host = $comp
            OT = $OT
            Query = 'Get-Basic'
            Chassis_Type = $script:obj.cschassisskunumber
            ChassisBootstate = $script:obj.cschassisbootupstate
            PowerSupplyState = $script:obj.cspowersupplystate
            Manufacturer = $script:obj.csmanufacturer
            ServiceTag = $servicetag
            Family = $script:obj.cssystemfamily
            SystemSKU = $script:obj.cssystemskunumber
            Model = $script:obj.csmodel
            PC_Type = $script:obj.cspcsystemtype
            Basic_Status = $script:obj.csstatus
            Thermal_Status = $script:obj.csthermalstate
      }

      $cshealth = "PSU - "
      $cshealth = -join($cshealth,$script:res.PowerSupplyState)
      $cshealth = -join($cshealth," `r`n")
      $cshealth = -join ($cshealth,"ThermalState - ")
      $cshealth = -join($cshealth,$script:res.thermal_status)
      $details = $script:res.family
      $details = -join($details," ")
      $details = -join($details,$script:res.PC_Type)
      $details = -join($details," ")
      $details = -join($details,$script:res.Chassis_Type)

      $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $OT
            Query = 'Get-Basic'
            Component = 'Basic'
            Health = $cshealth
            Name = $script:res.model
            Detail = $details
            Serial_Mac = $servicetag
      }
      $script:Basics += $script:res
      $script:health += $script:healthres   
}

function Get-Battery {
      param([string]$comp)

      if($script:err -ne $null) {
            $status = $script:Err
      }
      else {
            $status = "Successful"
      }

      $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag

      $script:obj = get-ciminstance -computername $comp win32_battery | select *

      if($script:obj -ne $null) {

            $runtime = $script:obj.Estimatedruntime
            if($runtime -eq 71582788) {
                  $runtime = 'AC Power'
            }
            else {
                  $runtime = (new-timespan -minutes $runtime).tostring()      
            }

            $bstatus = $script:BatteryStatus_map[[int]$script:obj.BatteryStatus]
            $chem = $script:Chemistry_map[[int]$script:obj.chemistry]

            $bathealth = "Battery Status - "
            $bathealth = -join($bathealth,$bstatus)
            $bathealth = -join($bathealth,", `r`n ")
            $bathealth = -join($bathealth,"Est. Charge - ")
            $bathealth = -join($bathealth,$script:obj.estimatedchargeremaining)
            $bathealth = -join($bathealth,", `r`n ")
            $bathealth = -join($bathealth,"Est. Runtime - ")
            $bathealth = -join($bathealth,$runtime)

            $batdetail = "Chemistry - "
            $batdetail = -join($batdetail,$chem)
            $batdetail = -join ($batdetail," `r`n")
            $batdetail = -join($batdetail,"Voltage - ")
            $batdetail = -join($batdetail,$script:obj.designvoltage)


            $script:res = [pscustomobject]@{
                  Room = $script:Room
                  Host = $comp
                  OT = $OT
                  Query = "Get-Battery"
                  Name = $script:obj.caption
                  Model = $script:obj.name
                  DeviceID = $script:obj.deviceid
                  Battery_Status = $bstatus
                  Availability = $script:availability_map[[int]$script:obj.availability]
                  Chemistry = $chem
                  Design_Voltage = $script:obj.designvoltage
                  Estimated_Charge = $script:obj.estimatedchargeremaining
                  Estimated_Runtime = $runtime
                  Status = $status

            }

            $script:healthres = [pscustomobject]@{
                  Room = $script:Room
                  Host = $comp
                  OT = $OT
                  Query = 'Get-Health'
                  Component = 'Battery'
                  Health = $bathealth
                  Name = $script:res.name
                  Detail = $batdetail
                  Serial_Mac = $script:res.deviceid

            }
      }
      else {
            $script:res = [pscustomobject]@{
                  Room = $script:Room
                  Host = $comp
                  OT = $OT
                  Query = 'Get-Battery'
                  Name = 'NA'
                  Model = 'NA'
                  DeviceID = 'NA'
                  Battery_Status = 'No Batteries Detected'
                  Availability = 'NA'
                  Chemistry = 'NA'
                  Design_Voltage = 'NA'
                  Estimated_Charge = 'NA'
                  Estimated_Runtime = 'NA'
                  Status = $status

            }

            $script:healthres = [pscustomobject]@{
                  Room = $script:Room
                  Host = $comp
                  OT = $OT
                  Query = 'Get-Health'
                  Component = 'Battery'
                  Health = 'No Batteries Detected'
                  Name = 'NA'
                  Detail = 'NA'
                  Serial_Mac = 'NA'

            }

      }
      $script:Battery += $script:res
      $script:Health += $script:healthres   
}

function Get-OS {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag

    $script:obj = get-computerinfo | select-object '*os*'
    $windows = get-computerinfo | select-object '*Windows*'
    $hotfixes = $script:obj.oshotfixes | sort-object -property hotfixid -descending | select-object -first 5
    $hotfixes = $hotfixes | select-object -expandproperty hotfixid
    $admins = get-localgroupmember -group "Administrators" | select-object Name
    $admin = "Admins - "

    foreach($a in $admins) {
      $admin = -join($admin," `r`n")
      $admin = -join($admin,$a.name)
    }
    
    $string = "Hotfixes - "
    foreach($h in $hotfixes){
        $string = -join($string,";`r`n ")
        $string = -join($string,$h)
    }

      $script:res = [pscustomobject]@{
        Room = $script:Room
        Host = $comp
        OT = $OT
        Query = 'Get-OS'
        OS_Name = $script:obj.osName
        OS_Version = $script:obj.osversion
        OS_Build = $script:obj.osbuildnumber
        OS_Serial = $script:obj.osserialnumber
        Last_5_Hotfixes = $string
        OS_LocalTime = $script:obj.oslocaldatetime
        OS_BootTime = $script:obj.oslastbootuptime
        OS_InstallDate = $script:obj.osinstalldate
        OS_Org = $script:obj.osorganization
        OS_Arch = $script:obj.osarchitecture
        OS_SP_Major_Version = $script:obj.osservicepackmajorversion
        OS_SP_Minor_Version = $script:obj.osservicepackminorversion
        OS_Status = $script:obj.osstatus
        Admin_Group = $admin
        Status = $status
      }
      $detail = "OS_Version - "
      $detail = -join($detail,$script:res.os_version)
      $detail = -join($detail," `r`n")
      $detail = -join($detail,"OS_Build - ")
      $detail = -join($detail,$script:res.os_build)
      $detail = -join($detail," `r`n")
      $detail = -join($detail,$admin)

      $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $script:res.ot
            Query = 'Get-Health'
            Component = 'OS'
            Health = $script:res.os_status
            Name = $script:res.os_Name
            Detail = $detail
            Serial_Mac = $script:res.os_serial            
            
      }

      $script:health += $script:healthres
      $script:OS = $script:res
      return $script:os | out-null
}

function Get-BIOS {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
    
    $script:obj = get-computerinfo | select-object '*BIOS*'
    $mobo = get-ciminstance -computername $comp win32_baseboard
    $version = ($script:obj.biosversion)

    $script:res = [pscustomobject]@{
        Room = $script:Room
        Host = $comp
        OT = $OT
        Query = 'Get-BIOS'
        BIOS_Version = $version
        BIOS_Release = $script:obj.biosreleasedate
        BIOS_Primary = $script:obj.biosprimarybios
        BIOS_Serial = $script:obj.biosseralnumber
        BIOS_Firmware_Type = $script:obj.biosfirmwaretype
        SMBIOS_Present = $script:obj.biossmbiospresent
        SMBIOS_Version = $script:obj.biossmbiosbiosversion
        BIOS_Status = $script:obj.biosstatus
        Motherboard_Status = $mbo.status
        MB_Serial = $mobo.serialnumber
        MB_Version = $mobo.version
        MB_Product = $mobo.product
        Status = $status
    }
    $detail = "BIOS FirmwareType - "
    $detail = -join($detail, $script:res.BIOS_Firmware_Type)
    $detail = -join($detail," `r`n")
    $detail = -join($detail,"BIOS Release - ")
    $detail = -join($detail,$script:res.bios_release)

    $bhealth = "BIOS Status - "
    $bhealth = -join($bhealth, $script:res.bios_status)
    $bhealth = -join($bhealth," `r`n")
    $bhealth = -join($bhealth,"Motherboard Status - ")
    $bhealth = -join($bhealth,$mbo.status)

    $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $script:res.ot
            Query = 'Get-Health'
            Component = 'BIOS'
            Name = $script:res.BIOS_Version
            Detail = $detail
            Serial_Mac = $script:res.bios_serial            
            Health = $bhealth
        }

        $script:health += $script:healthres

    $script:BIOS += $script:res
    return $script:bios | out-null
}

Function Get-CPU {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
    $script:obj = get-ciminstance -classname win32_processor | select-object *
    $fan = get-ciminstance -classname win32_fan | select-object *

    $fandetail = $script:availability_map[[int]$fan.availability]
    $fandetail = -join($fandetail,", ")
    $fandetail = -join($fandetail,$fan.status
      )
    $health = $script:cpustatus_map[[int]$script:obj.cpustatus]
    $health = -join($health,"; ")
    $health = -join($health,$script:obj.status)
    $health = -join($health," `r`nFan_Status - ")
    $health = -join($health,$fandetail)

    $serial = $script:obj.SerialNumber
    if($serial -eq $null) {
      $serial = "Unavailable"
    }

    $script:res = [pscustomobject]@{
        Room = $script:Room
        Host = $comp
        OT = $OT
        Query = 'Get-CPU'
        Model = $script:obj.name
        Caption = $script:obj.caption
        Device_ID = $script:obj.deviceid
        PartNumber = $script:obj.partnumber
        ProcessorID = $script:obj.processorid
        SerialNumber = $script:obj.serialnumber
        Type = $script:ProcessorType_map[[int]$script:obj.processortype]
        Architecture = $script:architecture_map[[int]$script:obj.architecture]
        Family = $script:family_map[[int]$script:obj.family]
        CPU_Status = $script:cpustatus_map[[int]$script:obj.cpustatus]
        CPU_Fan = $fandetail
        Availability = $script:availability_map[[int]$script:obj.availability]
        Socket = $script:obj.socketdesignation
        Upgrade_Method = $script:UpgradeMethod_map[[int]$script:obj.upgrademethod]
        '#_Cores' = $script:obj.numberofcores
        Current_Voltage = $script:obj.currentvoltage
        Load_Percent = ($script:obj.loadpercentage)
        Max_ClockSpeed = $script:obj.Maxclockspeed
        Current_ClockSpeed = $script:obj.currentclockspeed
        L2_CacheSize = $script:obj.l2cachesize
        L3_CacheSize = $script:obj.l3cachesize
        ThreadCount = $script:obj.threadcount
        Status = $status

    }
    $detail = "Device_ID - "
    $detail = -join($detail, $script:res.device_id)
    $detail = -join($detail," `r`n")
    $detail = -join($detail,"Caption - ")
    $detail = -join($detail,$script:res.caption)
    $serial = "ProcessorID - "
    $serial = -join($serial,$script:res.ProcessorID)
    $serial = -join($serial," `r`n")
    $serial = -join($serial,"Serial - ")
    $serial = -join($serial,$script:res.serialnumber)

    $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $script:res.ot
            Query = 'Get-Health'
            Component = 'CPU'
            Health = $health
            Name = $script:res.Model
            Detail = $detail
            Serial_Mac = $serial       
            
        }
        $script:health += $script:healthres
    $script:CPU += $script:res
    return $script:cpu | out-null
}

Function Get-RAM {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
    $script:mem = get-ciminstance win32_physicalmemory
    $script:memarray = get-ciminstance win32_physicalmemoryarray | select *

    foreach($m in $script:mem) {

      $manu = $script:ManufacturerRAM_Map[[string]$m.manufacturer]
      if($manu -eq $null) {
            $manuMap = $m.manufacturer
      }
      
        $script:res = [pscustomobject]@{
            Room = $script:Room
            Host = $comp
            OT = $OT
            Query = 'Get-RAM'
            Name = $m.tag
            Part_Number = $m.partnumber
            Serial_Number = $m.serialnumber
            Manufacturer = $manumap
            Form_Factor = $script:FormFactor_Map[[int]$m.formfactor]
            Capacity = $m.capacity
            Capacity_friendly = get-friendlysize $m.capacity
            Data_Width = $m.datawidth
            Memory_Type = $script:MemoryType_map[[int]$m.memorytype]
            Type_Detail = $m.typedetail
            Memory_ErrorCorrection = $script:MemoryErrorCorrection_map[[int]$memarray.memoryerrorcorrection]
            Memory_Use = $script:use_map[[int]$memarray.use] 
            Speed = $m.speed
            Config_Clockspeed = $m.configuredclockspeed
            Config_Voltage = $m.configuredvoltage
            Location = $m.devicelocator
            Status = $status
        }

        $name = $script:res.name
        $name = -join($name,"; ")
        $name = -join($name,$script:res.location)
        $detail = "Manufacturer - "
        $detail = -join ($detail,$manumap)
        $detail = -join($detail," `r`n")
        $detail = -join($detail,"Part_Number - ")
        $detail = -join($detail,$script:res.part_number)

        $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $script:res.ot
            Query = 'Get-Health'
            Component = 'RAM'
            Health = 'Unreported; memtest for memory problems'
            Name = $name
            Detail = $detail
            Serial_Mac = $script:res.Serial_number            
            
        }
        $script:health += $script:healthres
        $script:ram += $script:res
    }
    return $script:ram | out-null
}

Function Get-GPU {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag

    $script:vid = get-ciminstance win32_videocontroller

    foreach($v in $script:vid) {

        $script:res = [pscustomobject]@{
            Room = $script:Room
            Host = $comp
            OT = $OT
            Query = 'Get-GPU'
            GPU = $v.name
            Driver = $v.driverversion
            DriverDate = $v.driverdate
            Adapter_RAM_GB = $v.adapterram/1GB
            RAM_Type = $script:VideoMemoryType_map[[int]$v.videomemorytype]
            Adapter_DAC_Type = $v.adapterdacetype
            Current_VideoMode = $v.videomodedescription
            Video_Processor = $v.videoprocessor
            Availability = $script:availability_map[[int]$v.availability]
            GPU_Status = $v.status
            Dither_Type = $script:DitherType_map[[int]$v.dithertype]
            Video_Architecture = $script:videoarchitecture_map[[int]$v.videoarchitecture]
            Status = $status
        }
        $detail = "Driver - "
        $detail = -join($detail, $script:res.driver)
        $detail = -join($detail," `r`n")
        $detail = -join($detail,"DriverDate - ")
        $detail = -join($detail,$script:res.driverdate)
        $hdetail = "Status - "
        $hdetail = -join($hdetail,$script:res.GPU_Status)
        $hdetail = -join($hdetail," `r`n")
        $hdetail = -join($hdetail,"Availability - ")
        $hdetail = -join($script:res.availability)

        $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $script:res.ot
            Query = 'Get-Health'
            Component = 'GPU'
            Health = $hdetail
            Name = $script:res.gpu
            Detail = $detail
            Serial_Mac = 'Unavailable for GPUs'           
            
        }
        $script:health += $script:healthres

        $script:GPU += $script:res
    }
    return $script:gpu | out-null
}

function Get-Network {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:net = get-ciminstance win32_networkadapter | where {$_.physicaladapter -like 'True'} | select *
    $OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
    
    foreach($script:n in $script:net){ 

        $script:a = get-netadapter | where InterfaceIndex -eq $script:n.interfaceindex | select *
        $script:c = get-ciminstance win32_networkadapterconfiguration | where interfaceindex -eq $script:n.interfaceindex | select * 

        $script:res = [pscustomobject]@{
            Room = $script:Room
            Host = $comp
            OT = $OT
            Query = 'Get-Network'
            Name = $script:n.productname
            NetConnectionID = $script:n.netconnectionid
            NetEnabled = $script:n.netenabled
            Health = $script:NetConnectionStatus_map[[int]$script:n.netconnectionstatus]
            Device_ID = $script:n.deviceid
            Availability = $script:availability_map[[int]$script:n.availability]
            LinkSpeed = $script:a.linkspeed
            Speed = $script:n.speed
            Speed_Friendly = get-friendlysize $script:n.speed
            Adapter_TypeID = $script:AdapterTypeID_map[[int]$script:n.AdapterTypeID]
            Installed = $script:n.installed
            InterfaceIndex = $script:n.interfaceindex
            MAC = $script:n.macaddress
            Manufacturer = $script:n.manufacturer
            PhysicalAdapter = $script:n.physicaladapter
            ServiceName = $script:n.servicename
            Driver_Version = $script:a.driverversion
            Driver_Date = $script:a.driverdate
            SCSI_Interface = $script:a.iscsiinterface
            DHCP_Enabled = $script:c.DHCPEnabled
            DHCP_LeaseObtained = $script:c.dhcpleaseobtained
            DHCP_LeaseExpires = $script:c.DHCPLeaseExpires
            Status = $status
        }
        $detail = "Connection_ID - "  
        $detail = -join($detail, $script:res.netconnectionid)
        $detail = -join($detail," `r`n")
        $detail = -join($detail,"Adapter_Type - ")
        $detail = -join($detail,$script:res.adapter_typeid)

        $script:healthres = [pscustomobject]@{
            Room = $script:Room
            Host = $script:res.host
            OT = $script:res.ot
            Query = 'Get-Health'
            Component = 'Network'
            Health = $script:res.health
            Name = $script:res.name
            Detail = $detail
            Serial_Mac = $script:n.macaddress            
            
        }
        $script:health += $script:healthres
        $script:network += $script:res          
    }
    return $script:network | out-null
}

function run-stats {
      param(
  [string]$comp)

  $script:Task = 'Get-OS'
  get-OS -comp $comp

  $script:Task = 'Get-Battery'
  get-battery -comp $comp

  $script:Task = 'Get-BIOS'
  get-BIOS -comp $comp

  $script:Task = 'Get-CPU'
  get-CPU -comp $comp

  $script:Task = 'Get-RAM'
  get-RAM -comp $comp

  $script:Task = 'Get-GPU'
  get-GPU -comp $comp

  $script:Task = 'Get-Network'
  get-Network -comp $comp

  $script:Task = 'Get-Space'
  get-space -comp $comp

  $script:Task = 'Get-Basic'
  Get-Basic -comp $comp

  $script:Task = 'Run-All'

  $trash = [pscustomobject]@{
      This = 'is just'
      The = 'hacky'
      Garbage = 'to delete'
      Query = 'TRASH'
  }

  $all = @(
    $trash  
    $script:Health
    $script:Basics
    $script:Battery
    $script:OS
    $script:BIOS
    $script:CPU
    $script:RAM
    $script:GPU
    $script:Network
    $script:space
    
    )
  $script:allstats += $all
  $script:results = $script:allstats
    
  return $script:allstats | out-null
}