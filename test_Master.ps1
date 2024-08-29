<#
    .DESCRIPTION
        This is a master script to perform some of our common duties. The script is set up to allow looping through single nodes or lists of node names.

        There are various functions to choose from. Currently it is set up to generate .csv reports when being run against a list of nodes, and to output to the screen if looping through individual nodes. Looping will allow you to set a Room/Location or Timestamp to help organize results, and you can choose to save the results to file..

    .FUNCTIONS
        * Select-Task 
            * Decide which task to run

        * Set-Values
            * Gathers main info, like where and how to save reports, sets the room/location column, and whether or not to use transcripts
        
        * Get-Hostfile
            * Collects the list of nodes to run tasks against
        
        * Start-Log
        
        * Stop-Log
        
        * Cleanup
            * This will check a report and remove any blank lines or duplicates
        
        * Export-Data
            * This handles exporting report data to a csv
        
        * Append-Data
            * This handles running multi-part reports, such as for different rooms
        
        * Skip-Bad [-bad <Computer Name>] 
            * Creates a $Res object denoting that the computer was unabled to be contacted; this means it will still show up in the reported $Results
        
        * Get-Mode
            * Loop one node at a time or run against a list
        
        * Check-AD [-comp <Computer Name>]
            * Checks if node (or similar string) is in AD and stores information about it
        
        * Check-WinRM
            * Confirms that WinRM and winrm authentication are working on the remote machine
        
        * Check-Connection
            * Handles checking that nodes are in AD and contactable
            * Also calls requested functions once it determines if a node can or needs to be contacted
        
        * Get-Groups [-comp <Computer Name>]
        
        * Get-ADinfo [-comp <Computer Name>]
            * similar to check-ad, but pulls description, os, and ou formatted for reports
        
        * Get-OUMembers
            * Allows user to enter fully distinguished OU name to view [computer] members
        
        * Get-Mac [-comp <Computer Name>]
            * Gets MAC address from online machine
        
        * Generate-CSV [-comp <Computer Name>]
            * Adds node to a .csv for use with Veyon
        
        * Get-Software [-comp <Computer Name>]
            * Checks nodes for ALL or SPECIFIC pieces of software
        
        * Invoke-Script [-comp <Computer Name>]
            ** HIDDEN **
            * Invokes a user-defined script against the node
        
        * Invoke-Line [-comp <Computer Name> -line <Command>]
            ** HIDDEN **
            * Invokes a user-defined command against the node
        
        * Get-Linked [-comp <Computer Name>]
            * Gets linked group policy objects for a node's OUs
        
        * Update-Veyon
            * Uses the veyon-cli to import computer locations/MAC addresses via a .csv file; the .csv file can be created by Generate-CSV

        * Get-Space [-comp <Computer Name>]
            * Uses {get-psdrive C} and invoke-command to check the available space on the C:/

        * Get-FriendlySize [$bytes]
            * Converts bytes to appropriate values (kb/mb/gb/tb/pb)
            * Found on Stack Overflow [ https://stackoverflow.com/questions/63965384/how-to-convert-the-output-value-stored-in-variable-which-is-in-bytes-into-kb-mb ]

        * Get-Stats [-comp <Computer Name>]
            * gets a basic breakdown of hardware stats and OS info, focusing on what ITAM has asked about

        * Get-OS [-comp <Computer Name>]
            * Shows more detailed OS information about a computer, including the last 5 hotfixes, service pack versions, builds, and install date

        * Get-BIOS [-comp <Computer Name>]
            * Gathers basic BIOS info, not including configuration settings

        * Get-CPU [-comp <Computer Name>]
            *

        * Get-RAM [-comp <Computer Name>]
            * Gathers information about RAM

        * Get-GPU [-comp <Computer Name>]

        * Get-Network [-comp <Computer Name>]
            * Gathers information about network adapters, including bluetooth, wired, and wireless

        * Get-ConnectedDev [-comp <Computer Name>]
            * Checks for currently-connected PNP devices; test

    .NOTES
        Created by: KJA
        Modified: 2024-07-17

    .CHANGELOG

        7/31
            - Continuing adding separate functions for stats + adding hashtables for identifying specific numeric codes
        7/30
            - Began breaking out functions for getting computer stats

        7/29
            - Adjusted Get-Stats to gather info requested by ITAM - checking to see if there are further adjustments that need to be made
        7/26
            - Adjusted the [D]isplay preview while looping nodes to use a gridview that will open in a separate window; this allows one to more easily work with results without having to save everything to a .csv and open in excel
        7/25
            - Completedly reworked get-space to pull all pertinent info for boot disk and currently connected storage
        7/23
            - Added get-friendlysize function to convert bytes to more readable numbers
            - Began working on get-stats function to pull live info from machines
        7/17
            - Added forward and reverse NSlookup to Get-ADinfo
            - Added 'Get-Space' function for checking available space on C:/ for PCs
        6/27
            - Added ability to enter comma separated string for looping nodes or searching for specific pieces of software
        6/26
            - Began working on get-bios function
            - Added the ability to use (MM.dd_hh:mm) timestamps for the room/location label
                - Makes appending updated results to working report/task easier to organize by newest
        6/25
            - Added try/catch to update-veyon (for psversions older than 7)
            - Updated #Parameters and #Functions sections
        6/24 
            - Improved Check-Connection function
            - Added Cleanup function
        

#>
#Modules ===================================================================================================================
if(!(get-installedmodule -name 'PSWRiteColor' -EA ignore)) {

    set-psrepository -name psgallery -InstallationPolicy Trusted
    install-module -name pswritecolor
    import-module -name pswritecolor 
}
if(!(get-installedmodule -name 'ImportExcel' -EA ignore)) {
    set-psrepository -name psgallery -InstallationPolicy Trusted
    install-module -name importexcel
    import-module -name importexcel

}

#Hashtables ===================================================================================================================
$script:AcceleratorCapabilities_map = @{
      0 = 'Unknown'
      1 = 'Other'
      2 = 'Graphics Accelerator'
      3 = '3D Accelerator'
}

$script:AdapterTypeID_map = @{
      0 = 'Ethernet 802.3'
      1 = 'Token Ring 802.5'
      2 = 'Fiber Distributed Data Interface (FDDI)'
      3 = 'Wide Area Network (WAN)'
      4 = 'LocalTalk'
      5 = 'Ethernet using DIX header format'
      6 = 'ARCNET'
      7 = 'ARCNET (878.2)'
      8 = 'ATM'
      9 = 'Wireless'
     10 = 'Infrared Wireless'
     11 = 'Bpc'
     12 = 'CoWan'
     13 = '1394'
}

$script:Architecture_map = @{
      0 = 'x86'
      1 = 'MIPS'
      2 = 'Alpha'
      3 = 'PowerPC'
      6 = 'ia64'
      9 = 'x64'
}

$script:Availability_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'Running/Full Power'
      4 = 'Warning'
      5 = 'In Test'
      6 = 'Not Applicable'
      7 = 'Power Off'
      8 = 'Off Line'
      9 = 'Off Duty'
     10 = 'Degraded'
     11 = 'Not Installed'
     12 = 'Install Error'
     13 = 'Power Save - Unknown'
     14 = 'Power Save - Low Power Mode'
     15 = 'Power Save - Standby'
     16 = 'Power Cycle'
     17 = 'Power Save - Warning'
     18 = 'Paused'
     19 = 'Not Ready'
     20 = 'Not Configured'
     21 = 'Quiesced'
}

$script:BusType_map = @{
      0 = 'Internal'
      1 = 'ISA'
      2 = 'EISA'
      3 = 'MicroChannel'
      4 = 'TurboChannel'
      5 = 'PCI Bus'
      6 = 'VME Bus'
      7 = 'NuBus'
      8 = 'PCMCIA Bus'
      9 = 'C Bus'
     10 = 'MPI Bus'
     11 = 'MPSA Bus'
     12 = 'Internal Processor'
     13 = 'Internal Power Bus'
     14 = 'PNP ISA Bus'
     15 = 'PNP Bus'
     16 = 'Maximum Interface Type'
     17 = 'NVMe'
}

$script:Capabilities_map = @{
      0 = 'Unknown'
      1 = 'Other'
      2 = 'Sequential Access'
      3 = 'Random Access'
      4 = 'Supports Writing'
      5 = 'Encryption'
      6 = 'Compression'
      7 = 'Supports Removeable Media'
      8 = 'Manual Cleaning'
      9 = 'Automatic Cleaning'
     10 = 'SMART Notification'
     11 = 'Supports Dual Sided Media'
     12 = 'Predismount Eject Not Required'
}

$script:ConfigManagerErrorCode_map = @{
      0 = 'This device is working properly.'
      1 = 'This device is not configured correctly.'
      2 = 'Windows cannot load the driver for this device.'
      3 = 'The driver for this device might be corrupted, or your system may be running low on memory or other resources.'
      4 = 'This device is not working properly. One of its drivers or your registry might be corrupted.'
      5 = 'The driver for this device needs a resource that Windows cannot manage.'
      6 = 'The boot configuration for this device conflicts with other devices.'
      7 = 'Cannot filter.'
      8 = 'The driver loader for the device is missing.'
      9 = 'This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly.'
     10 = 'This device cannot start.'
     11 = 'This device failed.'
     12 = 'This device cannot find enough free resources that it can use.'
     13 = 'Windows cannot verify this device''s resources.'
     14 = 'This device cannot work properly until you restart your computer.'
     15 = 'This device is not working properly because there is probably a re-enumeration problem.'
     16 = 'Windows cannot identify all the resources this device uses.'
     17 = 'This device is asking for an unknown resource type.'
     18 = 'Reinstall the drivers for this device.'
     19 = 'Failure using the VxD loader.'
     20 = 'Your registry might be corrupted.'
     21 = 'System failure: Try changing the driver for this device. If that does not work, see your hardware documentation. Windows is removing this device.'
     22 = 'This device is disabled.'
     23 = 'System failure: Try changing the driver for this device. If that doesn''t work, see your hardware documentation.'
     24 = 'This device is not present, is not working properly, or does not have all its drivers installed.'
     25 = 'Windows is still setting up this device.'
     26 = 'Windows is still setting up this device.'
     27 = 'This device does not have valid log configuration.'
     28 = 'The drivers for this device are not installed.'
     29 = 'This device is disabled because the firmware of the device did not give it the required resources.'
     30 = 'This device is using an Interrupt Request (IRQ) resource that another device is using.'
     31 = 'This device is not working properly because Windows cannot load the drivers required for this device.'
}

$script:CpuStatus_map = @{
    0 = 'Unknown'
    1 = 'CPU Enabled'
    2 = 'CPU Disabled by User via BIOS Setup'
    3 = 'CPU Disabled By BIOS (POST Error)'
    4 = 'CPU is Idle'
    5 = 'Reserved'
    6 = 'Reserved'
    7 = 'Other'
}

$script:CPUStatusInfo_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'Enabled'
      4 = 'Disabled'
      5 = 'Not Applicable'
}

$script:DriveType_map = @{
    0 = 'Unknown'
    1 = 'No_Root_Directory'
    2 = 'Removable Disk'
    3 = 'Local Disk'
    4 = 'Network Drive'
    5 = 'Compact Disk'
    6 = 'RAM Disk'
}

$script:DitherType_map = @{
      1 = 'No dithering'
      2 = 'Dithering with a coarse brush'
      3 = 'Dithering with a fine brush'
      4 = 'Line art dithering'
      5 = 'Device does gray scaling'
}

$script:Family_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = '8086'
      4 = '80286'
      5 = '80386'
      6 = '80486'
      7 = '8087'
      8 = '80287'
      9 = '80387'
     10 = '80487'
     11 = 'Pentium(R) brand'
     12 = 'Pentium(R) Pro'
     13 = 'Pentium(R) II'
     14 = 'Pentium(R) processor with MMX(TM) technology'
     15 = 'Celeron(TM)'
     16 = 'Pentium(R) II Xeon(TM)'
     17 = 'Pentium(R) III'
     18 = 'M1 Family'
     19 = 'M2 Family'
     24 = 'K5 Family'
     25 = 'K6 Family'
     26 = 'K6-2'
     27 = 'K6-3'
     28 = 'AMD Athlon(TM) Processor Family'
     29 = 'AMD(R) Duron(TM) Processor'
     30 = 'AMD29000 Family'
     31 = 'K6-2+'
     32 = 'Power PC Family'
     33 = 'Power PC 601'
     34 = 'Power PC 603'
     35 = 'Power PC 603+'
     36 = 'Power PC 604'
     37 = 'Power PC 620'
     38 = 'Power PC X704'
     39 = 'Power PC 750'
     48 = 'Alpha Family'
     49 = 'Alpha 21064'
     50 = 'Alpha 21066'
     51 = 'Alpha 21164'
     52 = 'Alpha 21164PC'
     53 = 'Alpha 21164a'
     54 = 'Alpha 21264'
     55 = 'Alpha 21364'
     64 = 'MIPS Family'
     65 = 'MIPS R4000'
     66 = 'MIPS R4200'
     67 = 'MIPS R4400'
     68 = 'MIPS R4600'
     69 = 'MIPS R10000'
     80 = 'SPARC Family'
     81 = 'SuperSPARC'
     82 = 'microSPARC II'
     83 = 'microSPARC IIep'
     84 = 'UltraSPARC'
     85 = 'UltraSPARC II'
     86 = 'UltraSPARC IIi'
     87 = 'UltraSPARC III'
     88 = 'UltraSPARC IIIi'
     96 = '68040'
     97 = '68xxx Family'
     98 = '68000'
     99 = '68010'
    100 = '68020'
    101 = '68030'
    112 = 'Hobbit Family'
    120 = 'Crusoe(TM) TM5000 Family'
    121 = 'Crusoe(TM) TM3000 Family'
    122 = 'Efficeon(TM) TM8000 Family'
    128 = 'Weitek'
    130 = 'Itanium(TM) Processor'
    131 = 'AMD Athlon(TM) 64 Processor Family'
    132 = 'AMD Opteron(TM) Family'
    144 = 'PA-RISC Family'
    145 = 'PA-RISC 8500'
    146 = 'PA-RISC 8000'
    147 = 'PA-RISC 7300LC'
    148 = 'PA-RISC 7200'
    149 = 'PA-RISC 7100LC'
    150 = 'PA-RISC 7100'
    160 = 'V30 Family'
    176 = 'Pentium(R) III Xeon(TM)'
    177 = 'Pentium(R) III Processor with Intel(R) SpeedStep(TM) Technology'
    178 = 'Pentium(R) 4'
    179 = 'Intel(R) Xeon(TM)'
    180 = 'AS400 Family'
    181 = 'Intel(R) Xeon(TM) processor MP'
    182 = 'AMD AthlonXP(TM) Family'
    183 = 'AMD AthlonMP(TM) Family'
    184 = 'Intel(R) Itanium(R) 2'
    185 = 'Intel Pentium M Processor'
    190 = 'K7'
    200 = 'IBM390 Family'
    201 = 'G4'
    202 = 'G5'
    203 = 'G6'
    204 = 'z/Architecture base'
    250 = 'i860'
    251 = 'i960'
    260 = 'SH-3'
    261 = 'SH-4'
    280 = 'ARM'
    281 = 'StrongARM'
    300 = '6x86'
    301 = 'MediaGX'
    302 = 'MII'
    320 = 'WinChip'
    350 = 'DSP'
    500 = 'Video Processor'
}

$script:FormFactor_Map = @{
    0 = 'Unknown'
    1 = 'Other'
    2 = 'SIP'
    3 = 'DIP'
    4 = 'ZIP'
    5 = 'SOP'
    6 = 'Proprietary'
    7 = 'SIMM'
    8 = 'DIMM'
    9 = 'TSOP'
    10 = 'PGA'
    11 = 'RIMM'
    12 = 'SODIMM'
    13 = 'SRIMM'
    14 = 'SMD'
    15 = 'SSMP'
    16 = 'QFP'
    17 = 'TQFP'
    18 = 'SOIC'
    19 = 'LCC'
    20 = 'PLCC'
    21 = 'BGA'
    22 = 'FPBGA'
    23 = 'LGA'
}

$script:MemoryType_map = @{
      0 = 'Unknown'
      1 = 'Other'
      2 = 'DRAM'
      3 = 'Synchronous DRAM'
      4 = 'Cache DRAM'
      5 = 'EDO'
      6 = 'EDRAM'
      7 = 'VRAM'
      8 = 'SRAM'
      9 = 'RAM'
     10 = 'ROM'
     11 = 'Flash'
     12 = 'EEPROM'
     13 = 'FEPROM'
     14 = 'EPROM'
     15 = 'CDRAM'
     16 = '3DRAM'
     17 = 'SDRAM'
     18 = 'SGRAM'
     19 = 'RDRAM'
     20 = 'DDR'
     21 = 'DDR2'
     22 = 'DDR2 FB-DIMM'
}

$script:NetConnectionStatus_map = @{
      0 = 'Disconnected'
      1 = 'Connecting'
      2 = 'Connected'
      3 = 'Disconnecting'
      4 = 'Hardware Not Present'
      5 = 'Hardware Disabled'
      6 = 'Hardware Malfunction'
      7 = 'Media Disconnected'
      8 = 'Authenticating'
      9 = 'Authentication Succeeded'
     10 = 'Authentication Failed'
     11 = 'Invalid Address'
     12 = 'Credentials Required'
}

$script:PowerManagementCapabilities_map = @{
      0 = 'Unknown'
      1 = 'Not Supported'
      2 = 'Disabled'
      3 = 'Enabled'
      4 = 'Power Saving Modes Entered Automatically'
      5 = 'Power State Settable'
      6 = 'Power Cycling Supported'
      7 = 'Timed Power On Supported'
}

$script:ProcessorType_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'Central Processor'
      4 = 'Math Processor'
      5 = 'DSP Processor'
      6 = 'Video Processor'
}

$script:SecurityStatus_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'None'
      4 = 'External interface locked out'
      5 = 'External interface enabled'
}

$script:StatusInfo_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'Enabled'
      4 = 'Disabled'
      5 = 'Not Applicable'
}

$script:TypeDetail_map = @{
      1 = 'Reserved'
      2 = 'Other'
      4 = 'Unknown'
      8 = 'Fast-paged'
     16 = 'Static column'
     32 = 'Pseudo-static'
     64 = 'RAMBUS'
    128 = 'Synchronous'
    256 = 'CMOS'
    512 = 'EDO'
   1024 = 'Window DRAM'
   2048 = 'Cache DRAM'
   4096 = 'Non-volatile'
}

$script:UpgradeMethod_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'Daughter Board'
      4 = 'ZIF Socket'
      5 = 'Replacement/Piggy Back'
      6 = 'None'
      7 = 'LIF Socket'
      8 = 'Slot 1'
      9 = 'Slot 2'
     10 = '370 Pin Socket'
     11 = 'Slot A'
     12 = 'Slot M'
     13 = 'Socket 423'
     14 = 'Socket A (Socket 462)'
     15 = 'Socket 478'
     16 = 'Socket 754'
     17 = 'Socket 940'
     18 = 'Socket 939'
}

$script:VideoArchitecture_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'CGA'
      4 = 'EGA'
      5 = 'VGA'
      6 = 'SVGA'
      7 = 'MDA'
      8 = 'HGC'
      9 = 'MCGA'
     10 = '8514A'
     11 = 'XGA'
     12 = 'Linear Frame Buffer'
    160 = 'PC-98'
}

$script:VideoMemoryType_map = @{
      1 = 'Other'
      2 = 'Unknown'
      3 = 'VRAM'
      4 = 'DRAM'
      5 = 'SRAM'
      6 = 'WRAM'
      7 = 'EDO RAM'
      8 = 'Burst Synchronous DRAM'
      9 = 'Pipelined Burst SRAM'
     10 = 'CDRAM'
     11 = '3DRAM'
     12 = 'SDRAM'
     13 = 'SGRAM'
}
#Parameters ===================================================================================================================

$adfixflag = $null # indicates the listed hostname was corrected (ie 146828 to D146828)
$Append = $null # indicator for appending to veyon csv
$Appended = $null # indicator that Append function has been used
$appendResults = @()
$addesc = $null # stores description from AD
$adDescFlag = $null # indicator for pulling info from AD description for Veyon
$authflag = $null # indicator for winrm failing due to authentication
$checkADFlag = $null # indicates Task is only checking AD, not computers
$Checkpath = $null # variable to hold file loc information
$ChkChoice = $null # Whether to Overwrite/Rename/Append
$cmdLine = $null # scriptblock for invoke-command
$Confirm = $null # user confirmation
$ConfirmWrite = $null # indicates that user wants to write looped single node results to csv
$connflag = $null # indicator for winrm being configured
$Continue = $null # Repeat Task
$Count = $null # counter
$Directory = 'C:\Transcripts' # File path to save logs/transcripts
$dname = $null # stores distinguished name for get-ou search
$Err = $null #stores error message
$Filename = $null # Generated name of Transcript
$host = $null # Name of active node
$Hostfile = $null # List of node names
$inadflag = $null # shows if $host name was found in AD
$job = # stores info from running invoke-command as a job
$lastlogon = $null # stores last logon date (from AD)
$Length = $null # gets count of how many names are in a csv or array
$Logging = $null # Whether or not to enable logging
$mem = $null # stores info about RAM
$members = @() # stores members of OU or group
$memsum = $null #total memory in mem
$Mode = $null # [1] Looping single nodes or [2]against host file
$net = $null # stores info about connected net adapters
$Node = $null # Computer Name
$NodeList = @() # Multiple Computer Names
$nsforward = $null # holds information for nslookup
$Obj = @() # Working ps object; added to $Results
$oldname = $null # originally listed host name from hostfile
$os = $null # Stores OS info pulled from AD
$osdrive = $null # stores drive info for bootdisk
$osfree = $null # stores free space of bootdrive
$ou = $null # Stores OU from AD
$percentfree = $null # stores % free space
$pnp = $null # stores info about plug and play devices
$psver = $null #checks the version of powershell
$Report = $null # Location or specific title
$Res = @() # Another working ps object
$Results = @() # Final array of all $Obj after loop
$Reverse = $null # holds info for reverse nslookup
$Room = $null # Location information required by generate-csv
$Script = $null # path to the script for invoke-script
$SoftwareChoice = $null #Checking for [A]ll or [S]pecific
$SoftwareList = @() # string array of software names
$status = $null #stores status message
$Task = $null # Which main  function the user wants
$tcflag = $null # stores results from test-connection
$Test = $null
$tmpname = $null # stores the name pulled from AD to compare with the original search string
$updateTime = $null # allows the user to replace Location/Room variable with the date and time. Useful for when rerunning the same task against the same hostfile when Appending
$updateTimeFlag = $null # Indicator for if updateTime will be used for $room
$utilization = $null # var in get-space; holds used/free space. <10% space will result in a Status warning
$vid = $null # var to hold video/gpu/resolution info
$wsman = $null #checks if winrm is running
$wsmanauth = $null # stores results from checking winrm authentication
$wsmanconnect = $null # stores results from checking winrm

#Functions ====================================================================================================================

function Select-Task {

    write-Color "=================== ","Available Tasks"," ===================" -Color Cyan,Magenta,Cyan

    Write-Color "1:"," Get MAC ","address(es) of a computer/computers" -Color Yellow,Green,Blue
    Write-Color "2:"," Get installed software"," from a computer/computers" -Color Yellow,Green,Blue
    Write-Color "3:"," Get AD groups"," for a computer/computers" -Color Yellow,Green,Blue
    #Write-Color "Invoke-Script:"," Invoke a script"," against a computer/computers" -Color Yellow,Green,Blue
    Write-Color "4:"," Generate a .csv"," file for"," Veyon" -Color Yellow,Green,Blue
    #Write-Color "Invoke-Line:"," Invoke a specific command"," against a computer/computers" -Color Yellow,Green,Blue
    Write-Color "5:"," Get inherited GPO Links"," for a computer/computers" -Color Yellow,Green,Blue
    Write-Color "6:"," Update Veyon"," by uploading a .csv with node information" -Color yellow,green,blue
    Write-Color "7:"," Get OU Members ","of specific OU" -color yellow,green,blue
    Write-Color "8:"," Get AD Info"," of specific node names." -color yellow,green,blue
    Write-Color "9:"," Cleanup CSV"," to ensure there are no duplicates or blanks in a report." -color yellow,green,yellow
    Write-Color "10:"," Get Drive Space"," to check for used/free space on the C:/ of Windows computers." -color yellow,green,yellow
    Write-Color "11:"," Get Computer Stats"," to check for basic computer info [test case for ITAM]" -color yellow,green,yellow
    Write-Color "12:"," Get OS Info"," to find detailed OS info + last 5 hotfixes" -color yellow,green,yellow
    Write-Color "13:"," Get BIOS"," to pull basic BIOS information;"," NOTE: Does not check BIOS configuration." -color yellow,green,yellow
    Write-Color "14:"," Get CPU"," to retrieve detailed CPU information" -color yellow,green,yellow
    Write-Color "15:", "Get RAM"," to retrieve detailed RAM information" -color yellow,green,yellow
    Write-Color "16:"," Get GPU"," to retrieve detailed information about about both discrete and integrated GPUs" -color yellow,green,yellow
    Write-Color "17:"," Get Network"," to retrieve detailed information about network and bluetooth adapters" -color yellow,green,yellow
    Write-Color "Q:"," Quit" -Color Yellow,Red

    $input = Read-Host "Please make a selection"

    switch ($input) {

        '1' {

            Write-Color "You have selected ","1:"," Get MAC ","address(es) of a computer/computers" -Color Green,Yellow,Blue,Green
            $script:Task = 'Get-Mac'
            return
        }
        '2' {
            
            Write-Color "You have selected ","2:"," Get installed software"," from a computer/computers" -Color Green,Yellow,Blue,Green
            $script:Task = 'Get-Software'
            return
        }
        '3' {
            
            Write-Color "You have selected ","3:"," Get AD groups"," for a computer/computers" -Color Green,Yellow,Blue,Green
            $script:Task = 'Get-Groups'
            $script:checkADFlag = $true
            return
        }
        'Invoke-Script' {
            
            Write-Color "You have selected ","Invoke-Script:"," Invoke a script"," against a computer/computers" -Color Green,Yellow,Blue,Green
            $script:Task = 'Invoke-Script'
            return
        }
        '4' {
            
            Write-Color "You have selected ","4:"," Generate a .csv"," file for"," Veyon" -Color Green,Yellow,Blue,Green
            $script:Task = 'Generate-CSV'

            Write-Color "Note: Most nodes have descriptions in AD to denote their"," physical location." -color blue,yellow

            do {
                write-color "Would you like to use the ","AD descriptions"," as labels in your config?" -color blue,yellow,blue

                $script:ADdescFlag = Read-Host -prompt "[Y]es or any key for No"

                Write-color "You entered ","${script:ADdescFlag}",", is this correct?" -color blue,green,blue
                $script:Confirm = Read-Host -prompt "[Y]es or any key for No"
            } while ($script:Confirm -ne 'Y')
            return
        }
        'Invoke-Line' {
            
            Write-Color "You have selected ","Invoke-Line:"," Invoke a specific command"," against a computer/computers" -Color Green,Yellow,Blue,Green
            $script:Task = 'Invoke-Line'
            return
        }
        '5' {
            
            Write-Color "You have selected ","5:"," Get inherited GPO Links"," for a computer/computers" -Color Green,Yellow,Blue,Green
            write-color "Warning! ","Some versions of PowerShell may not report as expected. `nIf this report fails or returns blanks or Powershell objects for the linked GPO's, try switching to a different version of Powershell. ","`nCurrent Version: ","${script:psver}" -color red,yellow,blue,magenta
            $script:Task = 'Get-Linked'
            $script:checkADFlag = $true
            return
        }
        '6' {
            Write-Color "You have selected ","6:"," Update Veyon"," by uploading a .csv with node information" -Color green,yellow,blue,green
            $script:Task = 'Update-Veyon'
            return

        }
        '7' {
            Write-Color "You have selected","7:"," Get OU Members"," of a specific OU" -Color green,yellow,blue,green
            $script:Task = 'Get-OUMembers'
            $script:checkADFlag = $true
            return

        }
        '8' {
            Write-Color "You have selected","8:"," Get AD Info"," of specific nodes." -color green,yellow,blue,green
            $script:Task = 'Get-ADinfo'
            $script:checkADFlag = $true
            return
        }
        '9' {
            Write-color "You have selected ","3:"," Cleanup CSV" -color green,yellow,blue
            $script:Task = 'Cleanup'
            return
        }
        '10' {
            Write-color "You have selected ","10:"," Get Drive Space" -color green,yellow,blue
            $script:Task = 'Get-Space'
            return
        }
        '11' {
            Write-color "You have selected ","11:"," Get Computer Stats" -color green,yellow,blue
            $script:Task = 'Get-Stats'
            return
        }
        '12' {
            Write-color "You have selected ","12:"," Get OS Info" -color green,yellow,blue
            $script:Task = 'Get-OS'
            return
        }
        '13' {
            Write-color "You have selected ","13:"," Get BIOS" -color green,yellow,blue
            $script:Task = 'Get-BIOS'
            return
        }
        '14' {
            Write-color "You have selected ","14:"," Get CPU" -color green,yellow,blue
            $script:Task = 'Get-CPU'
            return
        }
        '15' {
            Write-color "You have selected ","15:"," Get RAM" -color green,yellow,blue
            $script:Task = 'Get-RAM'
            return
        }
        '16' {
            Write-color "You have selected ","16:"," Get GPU" -color green,yellow,blue
            $script:Task = 'Get-GPU'
            return
        }
        '17' {
            Write-color "You have selected ","17:"," Get Network" -color green,yellow,blue
            $script:Task = 'Get-Network'
            return
        }
        'Q' {

            Write-Color "Goodbye." -Color Red
            exit
            return
        }
    }        
}

function Set-Values {

    Write-Color "Would you like to save your output to ","${script:Directory}" -Color Yellow,Green
    $script:Confirm = Read-Host -Prompt '[Y]es or any key for No'
    Write-Color "${script:Confirm}" -Color Cyan
    if(!($script:Confirm -eq 'Y')) {
        $script:Directory = Read-Host -Prompt "Where would you like to save your output?"
        Write-Color "${script:Directory}" -Color Cyan
    }
    $repChar = '_'

    $tmpName = Read-Host -Prompt 'Please enter report title; note that spaces will be replaced by undescores'
    $script:Report = $tmpName -replace ' ', $repChar

    $script:CheckPath = "${script:Directory}\${script:Report}_${script:Task}.csv"

    if(test-path $script:Checkpath) {

        Write-Color "A file already exists by this name. Would you like to ","[O]","verwrite or ","[A]","ppend to the old file, or have the script automatically ","[R]","ename the new file?" -Color Yellow,Red,Yellow,Blue,Yellow,Green,Yellow
        $script:ChkChoice = Read-Host -prompt "[O]verwrite, [A]ppend, Automatically [R]ename, or any other key to cancel and end the script"
    }
    Write-Color -foregroundcolor Green "Report will be saved as ${script:Report}_${script:Task}.csv in directory ${script:Directory}"

    if($script:Mode -eq 2) {
        Get-Hostfile
        #$script:Room = Read-Host -prompt "Please enter a location for the device(s)"
        
        $script:Logging   = Read-Host -Prompt "Would you like a transcript? [Y]es or any key for no"
        Write-Color "${script:Logging}" -Color Cyan
        if(!($script:Logging -eq 'Y')) {
            Write-Color "If the outputs are not working as expected, rerun the script with a transcript." -Color blue
            Write-Color "You should be able to see where the error is happening." -Color blue
        }
    }    
}

function Get-Hostfile {
    $check = $false
    $current = pwd

    do {

        Write-color "Current directory is ","$current" -color green,blue        
        
        if($script:Task -eq 'Update-Veyon') {

            Write-Color "Your hostfile should be formatted as follows" -color yellow
            Write-Color "HOSTS.csv"
            Write-Color "Location",", ","Asset Tag",", ","Hostname",", ","MAC" -color green,yellow,green,yellow,green,yellow,green

            $script:Hostfile  = Read-Host -Prompt "Please enter your hostfile"

        }
        else {
            $script:Hostfile  = Read-Host -Prompt "Please enter your hostfile"
        }

        if(test-path $script:Hostfile) {
            $script:Length = (get-content $script:Hostfile).length
            Write-Color "${script:Hostfile} has ${script:Length} objects." -Color Cyan
            $preview = gc $script:Hostfile -first 5
            Write-Color "`nPreview of first five lines: `n"
            foreach ($i in $preview) {
                write-host($i)
            }

            $confirm = read-host -prompt "`nDoes this look like the correct file? [Y]es or any key for No"

            if($confirm -eq 'Y') {
                $check = $true
            }

            else {
                $check = $false
            }
        }

        else {
            
            Write-Color "${script:Hostfile}"," not found" -color red,yellow
            Write-color "Current directory is ","$current" -color yellow,red
            Write-Color "You may need to write the full file path. (ex: ","C:\Reports\Classroom.txt",", Do not include single or double quotes.)" -color yellow,blue,yellow
            $check = $false
        }
    } While ($check -eq $false)
    $script:NodeList = gc $script:Hostfile
    $script:Length = (get-content $script:Hostfile).length
}

function Start-Log {

    $Username   = $env:USERNAME
    $Hostname   = hostname
    $Datetime   = get-date -Format "HHmm_MM.dd.yy"
    $Fname   = "Transcript_${script:Report}_${script:Task}_[${Username}]-[${Hostname}]-[${Datetime}].txt"
    $LPathName = Join-Path -Path "$script:Directory" -ChildPath "$Fname"

    Write-Color "Checking intended filename and path: ","${LPathname}" -Color green,Blue

    if(Test-Path $LPathName) {

        Write-Color "A transcript already exists with this name. Appending to existing transcript."
        Start-Transcript -LiteralPath ("$LPathName") -Append
    }
 
    Start-Transcript -LiteralPath ("$LPathName") -NoClobber
}

function Stop-Log {

    Stop-Transcript
}

function Cleanup {

    $cleaning = gc $script:Checkpath 
    $tempCSVIn = $script:Checkpath
    $cleaner = $cleaning | where {$_.trim(',') -ne ""}
    $clean = $cleaner | convertfrom-csv | sort-object * -unique
    remove-item $tempcsvin
    $clean | convertto-csv | set-content $script:checkpath  
}

function Export-Data {

    if ($script:Results -ne $null -and $script:Results -is [System.Collections.IEnumerable] -and $script:Results | Select-Object -First 1 | Get-Member -Name Host,Room -MemberType Properties) {
        $script:Results = $script:Results | Sort-Object -Property Room,Host
    }

    if($script:Append -eq 'Y') {

        try {
            $script:Results | sort-object * -unique | Export-CSV -Path $script:CheckPath -usequotes Never -append -NoTypeInformation -force
        }
        catch {
            $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -append -NoTypeInformation -force
        }
    }

    else {
        if(test-path $script:checkpath) {

            switch($script:chkChoice) {

                'O' {
                    Write-Color "Forcing overwrite of ","${checkpath}" -Color Yellow,Green
                    remove-item $script:checkpath
                    try {
                        $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -usequotes Never -NoTypeInformation
                    }
                    catch {
                        $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -NoTypeInformation
                    }
                }
                'A' {
                    $header = gc $script:Checkpath | select-object -first 1
                    $script:results = $script:results | select-object -skip 1

                    try {
                        Write-Color "Appending to existing report ","${checkpath}" -Color yellow,green
                        try {
                            $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -append -usequotes Never -NoTypeInformation -force
                        }
                        catch {
                            $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -append -NoTypeInformation -force
                        }
                    }
                    catch {
                        Write-Color "Error. Unable to append to existing file." -Color Red
                        $timestamp = get-date -Format "HHmm_MM.dd.yy"
                        $script:Checkpath = $script:CheckPath.TrimEnd(".csv")
                        $newPath = -join ($checkpath,"_",$timestamp,".csv")
                        Write-Color "Saving file as"," ${newPath}" -color yellow,green
                        $script:Checkpath = $newPath
                        try {
                            $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -usequotes Never -NoTypeInformation
                        }
                        catch {
                            $script:Results | sort-object * -unique | Select * | Export-CSV -Path $script:CheckPath -NoTypeInformation
                        }
                    }
                }
                'R' {

                    $timestamp = get-date -Format "HHmm_MM.dd.yy"
                    $checkpath = $checkpath.TrimEnd(".csv")
                    $newPath = -join ($checkpath,"_",$timestamp,".csv")
                    Write-Color "Saving file as "," ${checkpath}" -color yellow,green
                    try {
                        $script:Results | Select * | Export-CSV -Path $script:CheckPath -usequotes Never -NoTypeInformation
                    }
                    catch {
                        $script:Results | Select * | Export-CSV -Path $script:CheckPath -NoTypeInformation
                    } 
                }
                Default {
                    Write-Color "Goodbye" -Color red
                }
            }
        }
        
        else {

            try {
                $script:Results | sort-object * -unique | Export-CSV -Path $script:CheckPath -usequotes Never -NoTypeInformation
            }
            catch {
                $script:Results | sort-object * -unique | Export-CSV -Path $script:CheckPath -NoTypeInformation
            } 
        }  
    } 
}

function Append-Data {

    $script:Continue = $null

    do {
            
        Write-Color "Would you like to add another batch to your file?" -color Yellow

        if($script:room -ne $null) {
            Write-Color "Previous batch: ","${script:Room}" -color blue,green
        }

        $script:Confirm = read-host -prompt "[Y] for yes, any key for no"

        if($script:Confirm -eq 'Y') {
            $script:Continue = 'Y'
            $script:Append = 'Y'
            
            Write-Color "Would you like to ","[R]","euse your previous hostfile or use a ","[N]","ew one?" -color Yellow,Green,Yellow,Blue,Yellow
            Write-Color "Previous hostfile: ","${script:Hostfile}" -color blue,green
            $temp = Read-Host -prompt "[R]euse Hostfile or [N]ew Hostfile, or any other key to cancel"

            switch($temp) {
                'R' {

                    $script:NodeList = gc $script:Hostfile
                    $script:Length = (get-content $script:Hostfile).length
                    Write-Color "Total Objects: ","${script:Length}" -Color blue,green 
                    Check-Connection
                    Export-Data
                    $script:Continue = 'Y'                         
                }
                'N' {
                    Get-Hostfile
                    Write-Color "Total Objects: ","${script:Length}" -Color blue,green 
                    Check-Connection
                    Export-Data
                    $script:Continue = 'Y'

                }
                Default {
                    Write-Color "Cancelling." -Color red
                }
            }
        }
        else { 
            $script:Append = $null
            $script:Continue = $null
        
        }
    } while($script:Continue -eq 'Y')
}

<#function Export-Data {

    if ($script:Results -ne $null -and $script:Results -is [System.Collections.IEnumerable] -and $script:Results | Select-Object -First 1 | Get-Member -Name Host,Room -MemberType Properties) {
        $script:Results = $script:Results | Sort-Object -Property Room,Host
    }

    if($script:Append -eq 'Y') {

        $script:xlpkg = $script:results | export-excel -path $script:checkpath -Append -worksheetname $task -tablename $task -passthru -autosize
        
    }

    else {
        if(test-path $script:checkpath) {

            switch($script:chkChoice) {

                'O' {
                    Write-Color "Forcing overwrite of ","${checkpath}" -Color Yellow,Green
                    remove-item $script:checkpath
                    $script:xlpkg = $script:results | export-excel -path $script:checkpath -worksheetname $task -tablename $task -passthru -autosize
                    
                }
                'A' {
                    $header = gc $script:Checkpath | select-object -first 1
                    $script:results = $script:results | select-object -skip 1

                    try {
                        Write-Color "Appending to existing report ","${checkpath}" -Color yellow,green
                        $script:xlpkg = $script:results | export-excel -path $script:checkpath -Append -worksheetname $task -tablename $task -passthru -autosize
                        
                    }
                    catch {
                        Write-Color "Error. Unable to append to existing file." -Color Red
                        $timestamp = get-date -Format "HHmm_MM.dd.yy"
                        $script:Checkpath = $script:CheckPath.TrimEnd(".xlsx")
                        $newPath = -join ($checkpath,"_",$timestamp,".xlsx")
                        Write-Color "Saving file as"," ${newPath}" -color yellow,green
                        $script:Checkpath = $newPath
                        $script:xlpkg = $script:results | export-excel -path $script:checkpath -Append -worksheetname $task -tablename $task -passthru -autosize
                    }
                }
                'R' {

                    $timestamp = get-date -Format "HHmm_MM.dd.yy"
                    $checkpath = $checkpath.TrimEnd(".xlsx")
                    $newPath = -join ($checkpath,"_",$timestamp,".xlsx")
                    Write-Color "Saving file as "," ${checkpath}" -color yellow,green
                    $script:xlpkg = $script:results | export-excel -path $script:checkpath -Append -worksheetname $task -tablename $task -passthru -autosize
                     
                }
                Default {
                    Write-Color "Goodbye" -Color red
                }
            }
        }
        
        else {

            $script:xlpkg = $script:results | export-excel -path $script:checkpath -Append -worksheetname $task -tablename $task -passthru -autosize
        }  
    }
    close-excelpackage $script:xlpkg
}

function Append-Data {

    $script:Continue = $null

    do {
            
        Write-Color "Would you like to add another batch to your file?" -color Yellow

        if($script:room -ne $null) {
            Write-Color "Previous batch: ","${script:Room}" -color blue,green
        }

        $script:Confirm = read-host -prompt "[Y] for yes, any key for no"

        if($script:Confirm -eq 'Y') {
            $script:Continue = 'Y'
            $script:Append = 'Y'
            
            Write-Color "Would you like to ","[R]","euse your previous hostfile or use a ","[N]","ew one?" -color Yellow,Green,Yellow,Blue,Yellow
            Write-Color "Previous hostfile: ","${script:Hostfile}" -color blue,green
            $temp = Read-Host -prompt "[R]euse Hostfile or [N]ew Hostfile, or any other key to cancel"

            switch($temp) {
                'R' {

                    $script:NodeList = gc $script:Hostfile
                    $script:Length = (get-content $script:Hostfile).length
                    Write-Color "Total Objects: ","${script:Length}" -Color blue,green 
                    Check-Connection
                    Export-Data
                    $script:Continue = 'Y'                         
                }
                'N' {
                    Get-Hostfile
                    Write-Color "Total Objects: ","${script:Length}" -Color blue,green 
                    Check-Connection
                    Export-Data
                    $script:Continue = 'Y'

                }
                Default {
                    Write-Color "Cancelling." -Color red
                }
            }
        }
        else { 
            $script:Append = $null
            $script:Continue = $null
        
        }
    } while($script:Continue -eq 'Y')
}#>

function Skip-Bad {

    param ( [string]$bad)
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Unable to connect"
    }

    switch ($script:Task) {

        'Get-Mac' {
            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $bad
                OT = 'NA'
                MAC = 'NA'
                IP = 'NA'
                Status = $status
            }
            $script:Results += $script:Res
        }
        'Generate-CSV' {
            if($script:adDescFlag -eq 'Y') {
                if($script:adDesc -ne $null) {
                    $desc = $script:adDesc
                }
            }
            else {
                $desc = $bad
            }
            
            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                ComputerName = $desc
                Host = $bad
                MAC = 'NA'
                Status = $status     
            } 
            $script:Results += $script:Res
        }
        'Get-Software' {

            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $bad
                ProductName = 'NA'
                ProductVersion = 'NA'
                Status = $status
            }
            $script:Results += $script:Res
        }

        'Get-Groups' {
            if($script:adDescFlag -eq 'Y') {
                if($script:adDesc -ne $null) {
                    $desc = $script:adDesc
                }
            }
            else {
                $desc = $bad
            }

            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $bad
                Group = 'NA'
                CanonicalName = 'NA'
                DistinguishedName = $script:ou
                Status = $status
            }
            $script:Results += $script:Res
        }

        'Invoke-Script' {
            $script:Res = [PSCustomObject]@{
                Host = $bad
                Output = 'NA'
                Status = $status
            }
            $script:Results += $script:Res            
        }

        'Invoke-Line' {
            $script:Res = [PSCustomObject]@{
                Host = $bad
                State = 'NA'
                Line = $line
                Output = 'NA'
                Status = $status
            }
            $script:Results += $script:Res            
        }

        'Get-Linked' {
            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $bad
                DisplayName = 'NA'
                Enabled = 'NA'
                Enforced = 'NA'
                Target = 'NA'
                Status = $status
            }
            $script:Results += $script:Res
        }
        'Get-ADinfo' {
            $nsforward = resolve-dnsname -name $bad -erroraction silentlycontinue
            if($nsforward -eq $null) {
                $reverse = $null
            }
            else {
                $reverse = resolve-dnsname -name $nsforward.ipaddress.tostring() -erroraction silentlycontinue
            }
            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $bad
                Desc = $script:adDesc
                OS = $script:OS
                OU = $script:ou
                Forward = $nsforward.ipaddress
                Reverse = $reverse.namehost
                Status = $status
            }
            $script:Results += $script:Res
        }
        'Get-Space' {
            $script:Res = [PSCustomObject]@{
                Host = $bad
                Drive = ''
                DriveModel = ''
                FriendlyName = ''
                DiskNumber = ''
                Connected = ''
                Health = ''
                Bus = ''
                IsBoot = ''
                PartitionStyle = ''
                BusType = ''
                FirmwareVersion = ''
                Location = ''
                Serial_Number = ''
                Size = ''
                Used = ''
                Free = ''
                '%_Free' = ''
                Status = $status
            }
            $script:Results += $script:Res            
        }
        'Get-Stats' {

            $script:res = [pscustomobject]@{
                Host = $comp
                OS = ''
                CurrentVersion = ''
                Version = ''
                Build = ''
                InstallDate = ''
                LastBoot = ''
                Boot_Disk = ''
                Boot_Disk_Health = ''
                Boot_Disk_Size = ''
                Boot_Disk_Free = ''
                'Boot_Disk_%Free' = ''
                Boot_Disk_Status = ''
                BIOSCaption = ''
                BIOSReleaseDate = ''
                SMBIOSVersion = ''
                SMBIOS_Major = ''
                SMBIOS_Minor = ''
                BIOS_Version = ''
                Model = ''
                Chassis = ''
                Processor = ''
                ProcessorDesc ='' 
                ProcessorArch = ''
                Cores = ''
                TotalRAM ='' 
                RAMCount = ''
                GPU1 = ''
                GPU1Driver ='' 
                GPU1DriverDate = '' 
                GPU1_RAM = ''
                Current_VideoMode = '' 
                GPU2 = ''
                GPU2Driver ='' 
                GPU2DriverDate = '' 
                GPU2_RAM = ''
                Monitor = ''
                CD_Drive = ''
                AdapterType = ''
                AdapterName = ''
                AdapterDriver = ''
                MAC = ''
                Speed = ''
                Status = $status
            }
        }

        'Get-OS' {
            $script:res = [pscustomobject]@{
                Host = $bad
                OS_Name = ''
                OS_Version = ''
                OS_Build = ''
                OS_Serial = ''
                Last_5_Hotfixes = ''
                OS_LocalTime = ''
                OS_BootTime = ''
                OS_InstallDate = ''
                OS_Org = ''
                OS_Arch = ''
                OS_SP_Major_Version = ''
                OS_SP_Minor_Version = ''
                OS_Status = ''
                Status = $status
            }
        }

        'Get-BIOS' {
            $script:res = [pscustomobject]@{
                Host = $bad
                BIOS_Version = ''
                BIOS_Release = ''
                BIOS_Primary = ''
                BIOS_Serial = ''
                BIOS_Firmware_Type = ''
                SMBIOS_Present = ''
                SMBIOS_Version = ''
                BIOS_Status = ''
                Status = $status
            }
        }

        'Get-CPU' {
            $script:res = [pscustomobject]@{
                Host = $bad
                Model = ''
                Caption = ''
                Device_ID = ''
                PartNumber = ''
                ProcessorID = ''
                SerialNumber = ''
                Type = ''
                Architecture = ''
                Family = ''
                CPU_Status = ''
                Availability = ''
                Socket = ''
                Upgrade_Method = ''
                '#_Cores' = ''
                Current_Voltage = ''
                Load_Percent = ''
                Max_Clockspeed = ''
                Current_Clockspeed = ''
                L2_CacheSize = ''
                L3_CacheSize = ''
                Threadcount = ''
                BaseBoard_Serial = ''
                BaseBoard_Product = ''
                Status = $status
            }
        }

        'Get-Ram' {
            $script:res = [pscustomobject]@{
                Host = $bad
                Name = ''
                Part_Number = ''
                Serial_Number = ''
                Form_Factor = ''
                Capacity = ''
                Data_Width = ''
                Memory_Type = ''
                Type_Detail = ''
                Speed = ''
                Config_Clockspeed = ''
                Config_Voltage = ''
                Location = ''
                Status = $status
            }
        }

        'Get-GPU' {
            $script:res = [pscustomobject]@{
                Host = $bad
                GPU = ''
                Driver = ''
                DriverDate = ''
                Adapter_RAM_GB = ''
                RAM_Type = ''
                Adapter_DAC_Type = ''
                Current_VideoMode = ''
                Video_Processor = ''
                Availability = ''
                GPU_Status = ''
                Dither_Type = ''
                Video_Architecture = ''
                Status = $status
            }
        }

        'Get-Network' {
            $script:res = [pscustomobject]@{
                Host = $bad
                Name = ''
                NetConnectionID = ''
                NetEnabled = ''
                Device_ID = ''
                Availability = ''
                LinkSpeed = ''
                Adapter_TypeID = ''
                Installed = ''
                InterfaceIndex = ''
                MAC = ''
                Manufacturer = ''
                PhysicalAdapter = ''
                ServiceName = ''
                Driver_Version = ''
                Driver_Date = ''
                SCSI_Interface = ''
                DHCP_Enabled = ''
                DHCP_LeaseObtained = ''
                DHCP_LeaseExpires = ''                
                Status = $status
            }
        }
    }
}

function Get-Mode {

    $First = 0

    if($script:Task -eq 'Get-OUMembers') { 
        $script:Mode = 1
    }

    else {

        do {

            if($First -ne 0) {
                Write-Color "You must enter ","'1'"," or ","'2'",", please try again." -Color Yellow,Red,Yellow,Red,Yellow
            }

            $script:Mode = Read-Host -Prompt '[1] Loop Single Nodes or [2] use a host file'

            if($script:Mode -eq 1) {
                Write-Color "You have chosen"," [1] ",", loop single nodes." -color Green, Yellow, Green
            }
            else {
                Write-Color "You have chosen"," [2] ",", to use a host file." -color Green,Yellow,Green
            }
            $First++

        } While (!($script:Mode -in ('1','2')))
    }
    if($script:Mode -eq 2) {
        $script:ConfirmWrite = 'Y'
        Set-Values
    }
}

function Check-AD {
    param ([string]$comp)

    $script:ADFixFlag = $false
    $script:inADFlag = $null
    $script:LastLogon = $null
    $script:tmpName = $null
    $script:OldName = $null
    $numeric = ($comp -match "^[\d\.]+$")


    #if $comp is numeric (OT) add wildcards to searchstr
    if($numeric -eq $true) {
        $searchstr = -join("*",$comp,"*")
    }
    else {
        $searchstr = $comp
    }

    #filter search AD for names that contain $comp
    $script:inAD = get-adcomputer -filter {name -like $searchstr} -properties *

    if($script:inAD -ne $null) {
        $script:tmpName = $script:inAD.Name
        $script:LastLogon = $script:inAD.lastlogondate
        $script:inADFlag = $true
        $script:os = $script:inAD.OperatingSystem
        $script:ou   = ($script:inAD.distinguishedName).split(',')[-6].split('=')[1]
        $script:adDesc = $script:inAD.description

        if($script:ou -eq $comp) {

                $script:ou = ($script:inAD.distinguishedName).split(',')[-5].split('=')[1]
            }

            if(!($script:os -like '*Windows*')) {
                $script:os = "Non-PC"
                $script:Err += "Non-PC"
            }

            #if found in AD, check if $comp and AD name are the same
            if($script:tmpName -ne $comp) {
                #only set fix flag to true if searchstr was for OT
                if ($numeric -eq $true) { 
                    $script:OldName = $script:Node
                    $script:Node = $script:tmpName
                    $script:ADFixFlag = $true
                }
            }
        }

    else {
        $script:inADFlag = $false
    }

    return $script:inADFlag
}

function Check-WinRM {
    param ([string]$comp)
    $script:ConnFlag = $null
    $script:AuthFlag = $null

    $script:wsmanConnect = test-wsman -computer $comp -erroraction silentlycontinue
    if($script:wsmanConnect -ne $null) {
        $script:ConnFlag = $true
        $script:WsmanAuth = test-wsman -computer $comp -authentication default -erroraction silentlycontinue
        if($script:WsmanAuth -ne $null) {
            $script:AuthFlag = $true
        }
        else {
            
            $script:AuthFlag = $false
        }
    }
    else {
        
        $script:ConnFlag = $false
    }
    return $script:AuthFlag
}

function Check-Connection {
    $script:tcFlag = $null
    $script:Continue = $null

    function Invoke-CommandAndStoreResults {
        $script:cmdLine = -join ($script:Task, " -comp ", $script:Node)
        Invoke-Expression $script:cmdLine | Tee-Object -Variable temp | Select-Object * | Write-Output
        $script:Res = $temp
        $script:Results += $script:Res
        $script:Results | Out-Null
    }
    function Handle-WinRM {
        $w_check = Check-Winrm -comp $script:Node
        if ($w_check) {
            Write-Color "`tWinRM Configured ","Correctly." -Color Blue, Green
            Invoke-CommandAndStoreResults
        } else {
            if ($script:ConnFlag -eq $true) {
                Write-Color "`tWinRM Authentication ","Failed." -Color Blue, Red
                $script:Err += " WinRM Auth Failure."
            } else {
                Write-Color "`tWinRM Connection ","Failed." -Color Blue, Red
                $script:Err += " Unable to connect to WinRM."
            }
            if($script:checkADFlag) {
                Write-Color "`tContinuing AD query." -color yellow
                Invoke-CommandAndStoreResults
            }
            else {
                Skip-Bad -bad $script:Node
            }
        }
    }
    function Handle-Offline {

        if ($script:LastLogon -ne $null) {
            $script:Err += " Offline. Last Logon: ${script:LastLogon}"
            Write-Color "`tOffline."," Last Logon: ","${script:LastLogon}" -Color Red, Yellow, Red
        } else {
            $script:Err += " Offline. Last Logon unknown."
            Write-Color "`tOffline. Last Logon unknown." -Color Red
        }
        if ($script:checkadflag -eq $true) {
            Write-Color "`tContinuing AD query." -color yellow
            Invoke-CommandAndStoreResults
        } else {
            Skip-Bad -bad $script:Node
        }
    }
    function Handle-NotInAD {
        if ($script:os -eq 'Non-PC') {
            Write-Color "`t${script.Node} ","may be a Mac; unable to process script commands." -Color Red, Yellow
            Invoke-CommandAndStoreResults
        } else {
            Write-Color "`t${script:Node} not found in AD." -Color Red
            $script:Err = "Not in AD."
            Skip-Bad -bad $script:Node
        }
    }
    function Handle-ADCheckSuccess {
        if ($script:ADFixFlag -eq $true) {
            Write-Color "`tMismatch between"," ${script:OldName} ","and ","${script:tmpName}" -Color Yellow, Red, Yellow, Green
            Write-Color "`tAdjusting search name to ","${script:tmpName}" -Color Yellow, Green
            $script:Err = "Name Corrected."
        }
        $tc_check = Test-Connection -ComputerName $script:Node -Quiet -ErrorAction SilentlyContinue
        
        if ($tc_check) {
            $script:tcFlag = $true
            Write-Color "`tTest Connection ","Successful." -Color Blue, Green
            if ($script:os -ne 'Non-PC') {
                Handle-WinRM

            } 
            else {
                
                if ($script:checkadflag -eq $true) {
                    Write-Color "`t${script.Node}"," is Non-PC, but AD query can continue." -Color yellow, blue
                    Invoke-CommandAndStoreResults
                } else {
                    Write-Color "`t${script.Node}"," is Non-PC; cannot interpret script commands." -color red, yellow
                    Skip-Bad -bad $script:Node
                }
            }
        } else {
            Handle-Offline
        }
    }
    function Set-Room {
        $script:updateTime = get-date -format "MM.dd_hh:mm"
        Write-Color "Auto Timestamp value is currently ","${script:updateTime}" -color blue,cyan
        do {
            Write-color "Would you like to use the ","[T]","imestamp ","(${script:UpdateTime})"," or enter a ","[C]","ustom Location?" -color blue,cyan,blue,cyan,blue,magenta,green
            $script:Choice = read-host -prompt "[T]imestamp or [C]ustom Location"
            }while(($script:Choice -ne 'T') -and ($script:Choice -ne 'C'))
            if($script:choice -eq 'C') {
                $script:Room = Read-Host -Prompt "Please enter Location"
            }
            else {
                $script:room = $script:updateTime
            }
    }

    # main logic

    if ($script:Mode -eq 1) {
        $script:First = 0
        Set-Room
        
        do {
            $script:Err = $null
            if (!($script:Continue -eq 'S') -and !($script:Continue -eq 'D') -and !($script:Continue -eq 'A') -and ($script:First -gt 0)) {
                Write-Color "You must enter ","[D]","isplay, ","[A]","dd, ","[C]","hange Room ","or ","[S]","top!" -Color Red, Yellow, Red, Yellow, Red, Yellow, Red, Yellow, Red
            } else {
                $script:NodeList = @()
                Write-Color "You may enter a ","single node"," [ex: D123456]"," or a ","list of nodes separated by commas ","[ex: 123456,D145345]","`nDo not use spaces after your commas." -color blue,green,cyan,blue,green,cyan,yellow
                $searchstr = Read-Host -Prompt 'Please enter your search string'
                $script:Nodelist +=($searchstr.split(','))

                Write-Color "You entered the following: " -color green
                foreach($script:node in $script:nodelist) {
                    write-color "${script:node}" -color cyan
                }
                foreach($script:Node in $script:NodeList) {
                    $script:Err = $null
                    Write-Color "Working on ","${script:Node}" -Color Yellow, Green   
                    $ad_check = Check-AD -comp $script:Node
                    if ($ad_check -eq $true) {
                        Handle-ADCheckSuccess
                    } else {
                        Handle-NotInAD
                    }
                }
            }
            $script:Continue = Read-Host -Prompt "[D]isplay Results, [A]dd new node, [C]hange Room, [S]top"
            if ($script:Continue -eq 'D') {
                Write-Color "Displaying current results" -Color Cyan
                Write-COlor "Results will appear in another window; closing the window will not affect this session." -color blue
                $script:Results | out-gridview
                $script:Continue = Read-Host -Prompt "[A]dd new node, [S]top"
            }
            if ($script:Continue -eq 'C') {
                set-room
                $script:Continue = Read-Host -Prompt "[D]isplay Results, [A]dd new node, [C]hange Room, [S]top"
            }
            $script:First++
        } while ($script:Continue -ne 'S')
    } else {
        $script:Count = 1
        Set-Room
        foreach ($script:Node in $script:NodeList) {
            $script:Err = $null
            Write-Color "Working on ","${script:Node}"," - ","${script:Count}"," out of ","${script:Length}" -Color Blue, Green, Blue, Yellow, Blue, Yellow   
            $ad_check = Check-AD -comp $script:Node
            if ($ad_check -eq $true) {
                Handle-ADCheckSuccess
            } else {
                Handle-NotInAD
            }
            $script:Count++
        }
    }
}

function Get-Groups {
    param ([string] $comp)
   
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:Obj = (Get-ADComputer -Identity $comp -properties canonicalname, MemberOf)
    $ou   = ($script:Obj.distinguishedName).split(',')[-6].split('=')[1]
    $cname   = $script:Obj.canonicalName
    $hname  = $script:Obj.Name
    $groups = $script:Obj.MemberOf -replace '^CN=([^,]+).+$','$1'

    if($ou -eq $hname) {

            $ou = ($script:Obj.distinguishedName).split(',')[-5].split('=')[1]
        }

    foreach($group in $groups) {
        $script:Res = [PSCustomObject]@{
            Room = $script:Room
            Host = $hname
            Group = $group
            CanonicalName = $cname
            DistinguishedName = $ou
            Status = $status

        }
        $script:Results += $script:Res
    }
    $script:Results | out-null
}

function Get-ADinfo {
    
    param ( [string]$comp)

    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $nsforward = resolve-dnsname -name $comp -erroraction silentlycontinue
        if($nsforward -eq $null) {
            $reverse = $null
        }
        else {
            $reverse = resolve-dnsname -name $nsforward.ipaddress.tostring() -erroraction silentlycontinue
        }

    $script:Res = [PSCustomObject]@{
        Room = $script:Room
        Host = $script:Node
        Desc = $script:adDesc
        OS = $script:os
        OU = $script:ou
        NSlookup = $nsforward.IPAddress
        ReverseNS = $reverse.Namehost
        Status = $status
    }
    $script:results += $script:res
    $script:results | out-null
}

function Get-OUMembers {
    
    $script:dname = read-host -prompt "Please enter the fully distinguished OU name `n('OU=FAR 220,OU=Labs,OU=Workstations,OU=Domain Computers,DC=SHSU,DC=EDU')`n"

    $script:members = Get-adcomputer -filter * -searchbase $script:dname -properties * | select name,description,distinguishedname
    foreach ($mem in $script:Members) {
        $ou = ($mem.distinguishedName).split(',')[-6].split('=')[1]
        $name = ($mem.name)
        $desc = ($mem.description)

        if($ou -eq $name) {

            $ou = ($mem.distinguishedName).split(',')[-5].split('=')[1]
        }

        $script:Res = [PSCustomObject]@{
            Name = $name
            Desc = $desc
            OU = $ou
        }
        $script:Results += $script:Res
        $script:Results | out-null
    }
}

function Get-Mac {
    param ( [string] $comp)
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:Obj = get-netipconfiguration -computername $comp -Detailed | where-object {$_.NetProfile.Name -like 'SHSU.EDU' }
    
    $script:Res = [PSCustomObject]@{
        Room = $script:Room
        Host = $comp
        OT = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
        MAC = $script:Obj.NetAdapter.LinkLayerAddress
        IP = $script:Obj.IPv4Address.IPAddress
        Status = $status
    }
    
    $script:Results += $script:Res
    $script:Results | out-null
    #return $script:Res
}

function Generate-CSV {
    param ( [string] $comp)

    $status = if ($script:Err -ne $null) { $script:Err } else { "Successful" }

    $desc = if ($script:ADdescFlag -eq 'Y') { 
        $script:adDesc 
    } else { 
        (Get-CimInstance -ComputerName $comp win32_systemenclosure).smbiosassettag 
    }

    $script:Obj = get-netipconfiguration -computername $comp -Detailed | where-object {$_.netprofile.name -like 'SHSU.EDU' }

    $script:Res = [PSCustomObject]@{
        Room = $script:Room
        ComputerName = $desc
        Host = $comp
        MAC = $script:Obj.NetAdapter.LinkLayerAddress
        Status = $status

    }
    $script:Res | out-null
    $script:Results += $script:Res
    $script:Results | out-null
}

function Get-Software {
    param ( [string] $comp)
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    if ($script:SoftwareChoice -eq 'A') {

        $script:Obj = Get-CimInstance -Namespace root/cimv2/sms -ClassName SMS_InstalledSoftware -ComputerName $comp | Select-Object -Property ProductName, ProductVersion

        $Prods = $script:Obj.ProductName
        $cname = $comp       

        foreach ($product in $Prods) {
            $test = $script:Obj | where-object {$_.ProductName -like $product}
            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $cname
                ProductName = $test.ProductName
                ProductVersion = $test.ProductVersion
                Status = $status
            }
            $script:Results += $script:Res
        }
    }

    else {

        foreach ($product in $script:SoftwareList) {
            $product = -join('*',$product,'*')
            $script:Obj = Get-CimInstance -Namespace root/cimv2/sms -ClassName SMS_InstalledSoftware -ComputerName $comp | Where-Object {$_.ProductName -like $product} | Select-Object -Property ProductName, ProductVersion

            if(!($script:Obj -eq $null)) {

                foreach ($ob in $script:Obj) {
        
                    $script:Res = [PSCustomObject]@{
                        Room = $script:Room
                        Host = $comp
                        ProductName = $ob.ProductName
                        ProductVersion = $ob.ProductVersion
                        Status = $status
                    }
                    $script:Results += $script:Res
                }
            }
            else {

                $script:Res = [PSCustomObject]@{
                    Room = $script:Room
                    Host = $comp
                    ProductName = -join($product, ' NA')
                    ProductVersion = 'Not Installed'
                    Status = $status
                }
                $script:Results += $script:Res
            }
        }
    }
    $script:Results | out-null
}

function Invoke-Script {
    param([string]$comp)
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    Invoke-command -computername $comp -Filepath $script:Script | tee-object -variable temp | Select * | out-null
    
    $script:Res = [PSCustomObject]@{
        Host = $comp
        Output = $temp
        Status = $status
    }
    $script:Results += $script:res
    $script:Results | out-null
}

function Invoke-Line {
    param(
        [string]$comp,
        [string]$line)
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $jname = -join("tmp_"+$comp)

    $script:job = invoke-command -computername $comp -scriptblock {$using:line} -asjob -jobname $jname |  get-job | wait-job

    $script:r = $script:job.childjobs
    $check1 = $script:r.output.value
    $check2 = $script:r.output

    if ($check1 -eq $null) {
        $script:Res = [PSCustomObject]@{
            Host = $script:job.location
            State = $script:job.State
            Line = $line
            Output = $check2
            Status = $status
        }
    }

    else {
        $script:Res = [PSCustomObject]@{
            Host = $script:job.location
            State = $script:job.State
            Line = $line
            Output = $check1
            Status = $status
        }
    }


    $script:Results += $script:Res
    $script:Results | out-null

    remove-job $jname
}

function Get-Linked {
    param([string]$comp)
    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:Obj = Get-ADComputer -identity $comp
    $script:Test = ($script:Obj.distinguishedName -split '(?=OU=)',2)[1]

    $dnames = (invoke-command -session $(get-pssession -name winpscompatsession) -scriptblock {(Get-GPInheritance -target $script:Test).inheritedGPOLinks})
    
    foreach($dname in $dnames) {
        $script:Res = [PSCustomObject]@{
            Room = $script:Room
            Host = $comp
            DisplayName = $dname.DisplayName
            Enabled = $dname.Enabled
            Enforced = $dname.Enforced
            Order = $dname.Order
            Target = $dname.Target
            Status = $status
        }
        $script:Results += $script:Res
    }
    $script:Results | out-null
}

function Update-Veyon {

    if($script:Task -ne 'Update-Veyon') {
        $file = $script:Checkpath
    }
    else {
        $file = $script:Hostfile
    }

    $vpath = 'C:\Program Files\Veyon\Veyon-cli.exe'
    $tmpFile = $file.trimend(".csv")
    $tempCSVeyon = "C:\Temp\tmpconfig.csv"
    $vset = "%location%,%name%,%host%,%mac%"
    #$vargs = "networkobjects import ${tempCSVeyon} format ${vset}"

    if($script:psver.Major -lt 7) {
        (Get-Content $script:hostfile) | % {$_ -replace '"', ""} | out-file -FilePath $script:hostfile -Force -Encoding ascii
    }

    if (test-path $vpath) {

        $check = $true

    }

    else {
        $check = $false

        do {

            Write-Color "Warning, veyon-cli.exe not found at expected location"," ${vpath}" -color yellow,red
            $vpath = Read-Host -Prompt "Please enter the full path for veyon-cli.exe"
            
            if (test-path $vpath) {

                $check = $true

            }

        } while(!($check = $true))

    }

    if($check = $true) {

        $temp = import-csv $file | where-object {!($_.status -like '*Non-PC*')}
        $temp = $temp | sort-object -property host | select Room,ComputerName,Host,MAC
        try { $temp | export-csv -path $tempCSVeyon -usequotes never -notypeinformation -noheader }
        catch {$temp | % {$_ -replace '"',""} | out-file -path $tempCSVeyon -encoding ascii }

        Start-Process -filepath $vpath -argumentlist "networkobjects import ${tempCSVeyon} format ${vset}" -verbose -wait -passthru -nonewwindow
        start-process -filepath $vpath -argumentlist "service restart" -nonewwindow

    }
    remove-item $tempCSVeyon    
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

    $space = @()

    $script:volumes = invoke-command -computername $comp -scriptblock {get-volume | where-object {$_.filesystemtype -ne 'Unknown' -and $_.driveletter -ne $null}}

    foreach($vol in $volumes) {

        $drive = $vol.DriveLetter

        $infov = invoke-command -computername $comp -scriptblock {get-volume | where driveletter -like $using:drive | get-partition | select *}
        $infops = invoke-command -computername $comp -scriptblock {get-psdrive $using:drive}
        $infodrive = invoke-command -computername $comp -scriptblock {get-disk | where disknumber -like $using:infov.disknumber | select *}
        $free = (($infops.free / $infov.size)).tostring("P")

        $script:res = [pscustomobject]@{
          Host = $comp
          Query = 'Get-Space'
          Drive = $drive
          DriveModel = $infodrive.friendlyname
          DriveType = $vol.drivetype
          Health = $vol.healthstatus
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
        $space += $res
    }
    return $space
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

function Get-Stats {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:obj = invoke-command -computername $comp -scriptblock {get-computerinfo | select-object *}
    $script:vid = invoke-command -computername $comp -scriptblock {get-ciminstance win32_videocontroller}

    $baseboard = invoke-command -computername $comp -scriptblock {get-ciminstance win32_baseboard -property *}
    $enclosure = invoke-command -computername $comp -scriptblock {get-ciminstance win32_systemenclosure | select *}

    $osdrive = ($script:obj.ossystemdrive).split(":")[0]
    $osfree = (invoke-command -computername $comp -scriptblock {get-psdrive $using:osdrive | select-object -expandproperty free})

    $script:bootdisk = invoke-command -computername $comp -scriptblock {get-volume | where driveletter -like $using:osdrive | get-partition | get-disk}

    $cd = invoke-command -computername $comp -scriptblock {get-ciminstance win32_cdromdrive | select-object -expandproperty Caption}

    $percentFree = ($osfree / $script:bootdisk.size).tostring("P")

    if($vid[1] -ne $null) {
        $GPU2 = $vid[1].name
        $GPU2Driver = $vid[1].driverversion
        $GPU2DriverDate = $vid[1].driverdate
        $GPU2_RAM = $vid[1].adapterram/1GB
    }
    else {
        $GPU2 = ''
        $GPU2Driver = ''
        $GPU2DriverDate = ''
        $GPU2_RAM = ''
    }

    $script:mem = invoke-command -computername $comp -scriptblock {get-ciminstance win32_physicalmemory }
    $script:net = invoke-command -computername $comp -scriptblock {get-netadapter | where {$_.status -eq 'Up' -and $_.virtual -ne 'True' } | select-object *}
    
    
    $script:memSum = $script:mem | measure-object -property Capacity -sum
    $script:pnp = invoke-command -computername $comp -scriptblock {get-pnpdevice | select-object *} 
#this res is wip new version
    $script:res = [pscustomobject]@{
        Host = $comp
        AssetTag = $enclosure.smbiosassettag
        Serial_Number = $enclosure.serialnumber
        LockPresent = $enclosure.lockpresent
        OS = $script:obj.windowsproductname
        CurrentVersion = $script:obj.windowscurrentversion
        Version = $script:obj.osversion
        Build = $script:obj.osbuildnumber
        InstallDate = $script:obj.windowsinstalldatefromregistry
        LastBoot = $script:obj.oslastbootuptime
        Boot_Disk = $script:bootdisk.friendlyname
        Boot_Disk_Health = $script:bootdisk.healthstatus
        Boot_Disk_Size = get-friendlysize $script:bootdisk.size
        Boot_Disk_Free = get-friendlysize $osfree
        'Boot_Disk_%Free' = $percentfree
        Boot_Disk_Status = $script:bootdisk.healthstatus
        BIOSCaption = $script:obj.bioscaption
        BIOSReleaseDate = $script:obj.biosreleasedate
        SMBIOSVersion = $script:obj.biossmbiosbiosversion
        SMBIOS_Major = $script:obj.biossmbiosmajorversion
        SMBIOS_Minor = $script:obj.biossmbiosminorversion
        BIOS_Version = $script:obj.biosversion
        Model = $script:obj.csmodel
        Chassis = $script:obj.cschassisskunumber
        BaseBoard = $baseboard.name
        BaseBoard_Model = $baseboard.model
        BaseBoardProd = $baseboard.product
        BaseBoardPartNum = $baseboard.partnumber
        BaseBoardSerialNum = $baseboard.serialnumber
        BaseBoardVersion = $baseboard.version
        BaseBoardConfig = $baseboard.configoptions
        Processor = $script:obj.csprocessors.name
        ProcessorDesc = $script:obj.csprocessors.description
        ProcessorArch = $script:obj.csprocessors.architecture
        Cores = $script:obj.csprocessors.numberofcores
        TotalRAM = get-friendlysize $script:memSum.sum
        RAMCount = $script:memSum.count
        GPU1 = $script:vid[0].name
        GPU1Driver = $script:vid[0].driverversion
        GPU1DriverDate = $script:vid[0].driverdate
        GPU1_RAM_GB = $script:vid[0].adapterram/1GB
        Current_VideoMode = $script:vid[0].videomodedescription
        GPU2 = $GPU2
        GPU2Driver = $GPU2Driver 
        GPU2DriverDate = $GPU2DriverDate
        GPU2_RAM_GB = $GPU2_RAM
        Monitor = $script:pnp.Monitor.friendlyname
        CD_Drive = $cd.caption -join('; ',$cd.drive)
        AdapterType = $script:net.ifalias -join [System.Environment]::NewLine
        AdapterName = $script:net.ifDesc -join [System.Environment]::NewLine
        AdapterWake = $script:net.devicewakeupenable -join [System.Environment]::NewLine
        AdapterDriver = $script:net.driverversionstring -join [System.Environment]::NewLine
        MAC = $script:net.macaddress -join [System.Environment]::NewLine
        Speed = $script:net.linkspeed -join [System.Environment]::NewLine
        Status = $status

    }
    $script:results += $script:res
}

function Get-OS {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:obj = invoke-command -computername $comp -scriptblock {get-computerinfo | select-object '*os*'}
    $hotfixes = $script:obj.oshotfixes | sort-object -property hotfixid -descending | select-object -first 5
    $hotfixes = $hotfixes | select-object -expandproperty hotfixid
    $string = $null
    foreach($h in $hotfixes){
        $string += -join ($h.hotfixid,"; ")
    }
    $string = $string.trimend("; ")

    $script:res = [pscustomobject]@{
        Host = $comp
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
        Status = $status

    }

    $script:results += $script:res
}

function Get-BIOS {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    
    $script:obj = invoke-command -computername $comp -scriptblock {get-computerinfo | select-object '*BIOS*'}
    $version = ($script:obj.biosversion)
    $asset = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag

    $script:res = [pscustomobject]@{
        Host = $comp
        BIOS_Version = $version
        BIOS_Release = $script:obj.biosreleasedate
        BIOS_Primary = $script:obj.biosprimarybios
        BIOS_Serial = $script:obj.biosseralnumber
        BIOS_Firmware_Type = $script:obj.biosfirmwaretype
        BIOS_AssetTag = (get-ciminstance -computername $comp win32_systemenclosure).smbiosassettag
        SMBIOS_Present = $script:obj.biossmbiospresent
        SMBIOS_Version = $script:obj.biossmbiosbiosversion
        BIOS_Status = $script:obj.biosstatus
        Status = $status
    }

    $script:BIOS += $script:res
}

Function Get-CPU {

    
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    
    $script:obj = invoke-command -computername $comp -scriptblock {get-ciminstance -classname win32_processor | select-object *}

    $script:res = [pscustomobject]@{
        Host = $comp
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
    $script:CPU += $script:res
}

Function Get-RAM {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $ram = @()

    $script:mem = invoke-command -computername $comp -scriptblock {get-ciminstance win32_physicalmemory}

    foreach($m in $script:mem) {
        $script:res = [pscustomobject]@{
            Host = $comp
            Query = 'Get-RAM'
            Name = $m.tag
            Part_Number = $m.partnumber
            Serial_Number = $m.serialnumber
            Form_Factor = $m.formfactor
            Capacity = $m.capacity
            Capacity_friendly = get-friendlysize $m.capacity
            Data_Width = $m.datawidth
            Memory_Type = $m.memorytype
            Type_Detail = $m.typedetail
            Speed = $m.speed
            Config_Clockspeed = $m.configuredclockspeed
            Config_Voltage = $m.configuredvoltage
            Location = $m.devicelocater
            Status = $status
        }
        $ram += $res
    }
    return $ram
}

Function Get-GPU {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:vid = invoke-command -computername $comp -scriptblock {get-ciminstance win32_videocontroller}

    foreach($v in $script:vid) {

        $script:res = [pscustomobject]@{
            Host = $comp
            GPU = $v.name
            Driver = $v.driverversion
            DriverDate = $v.driverdate
            Adapter_RAM_GB = $v.adapterram/1GB
            RAM_Type = $script:VideoMemoryType_map[[int]$v.videomemorytype]
            Adapter_DAC_Type = $v.adapterdacetype
            Current_VideoMode = $v.videomodedescription
            Video_Processor = $v.videoprocessor
            Availability = $script:availability_map[[int]$v.availability]
            GPU_Status = $script:v.status
            Dither_Type = $script:DitherType_map[[int]$v.dithertype]
            Video_Architecture = $script:videoarchitecture_map[[int]$v.videoarchitecture]
            Status = $status
        }
        $script:GPU += $script:res
    }
}

function Get-Network {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
    $network = @()

    $script:net = invoke-command -computername $comp -scriptblock {get-ciminstance win32_networkadapter | where {$_.physicaladapter -like 'True'} | select *}

    
    foreach($script:n in $script:net){ 

        $script:a = invoke-command -computername $comp -scriptblock {get-netadapter | where InterfaceIndex -eq $script:n.interfaceindex | select *}  
        $script:c = invoke-command -computername $comp -scriptblock {get-ciminstance win32_networkadapterconfiguration | where interfaceindex -eq $script:n.interfaceindex | select * }

        $script:res = [pscustomobject]@{
            Host = $comp
            Query = 'Get-Network'
            Name = $script:n.productname
            NetConnectionID = $script:n.netconnectionid
            NetEnabled = $script:n.netenabled
            Device_ID = $script:n.deviceid
            Availability = $script:n.availability
            LinkSpeed = $script:a.linkspeed
            Speed = $script:a.speed
            Speed_Friendly = get-friendlysize $script:a.speed
            Adapter_TypeID = $script:n.AdapterTypeID
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
        $network += $res          
    }
    return $network
}

Function Get-ConnectedDev {
    param([string]$comp)

    if($script:err -ne $null) {
        $status = $script:Err
    }
    else {
        $status = "Successful"
    }
}

#Main program =================================================================================================================
cls
$script:psver = $psversiontable.PSVersion

if (!($script:psver.Major -gt 6)) {

    Write-color "Warning! ","This script was written for PowerShell version 7 or greater." -color red,yellow
    Write-Color "You have version ","${script:psver}" -color yellow,green
    Write-Color "Exporting to CSV may not behave as expected. " -color yellow
    Write-Color "For the best experience, please make sure you run this script in Powershell version 7.0 or greater. " -color cyan
}

Select-Task

if(($script:task -ne 'Update-Veyon') -and ($script:task -ne 'Cleanup')) {
    get-mode
    if ($script:Logging -eq 'Y') {

        Start-Log   
    }
}

switch ($script:Task) {

    'Get-Mac' {

        Check-Connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }

    }

    'Get-Software' {
        $script:Continue = $null

        $script:SoftwareChoice = Read-Host -Prompt 'Are you checking for [A]ll software, or [S]pecific software'
        Write-Host $script:SoftwareChoice

        if ($script:SoftwareChoice -eq 'S') {
            do {
                write-color "You may enter a ","single search string"," [ex: Acrobat]"," or a ","list of search strings separated by commas [ex: Acrobat,Veyon,Adobe Photoshop,Chrome]","`nDo not use spaces after your commas." -color blue,green,cyan,blue,green,cyan,yellow
                $searchstr = Read-Host 'Please enter a search string or list of search strings'
                $script:SoftwareList += ($searchstr.split(','))
                
                Write-Color 'You entered the following search string(s)' -color green
                foreach($sft in $script:SoftwareList) {
                    Write-Color "${sft}" -color cyan
                }
                write-color "Is this list correct?" -color yellow
                $script:Continue = Read-host -prompt "[Y]es or any key for No"
                
            } while ($script:Continue -ne 'Y')
            
            
        }
        Check-Connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }

    'Get-Groups' {

        Check-Connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }

    }

    'Invoke-Script' {
        $script:Continue = $null

        do {

            $script:Script = Read-Host -Prompt 'Please enter the script you wish to run. If it is not in the same directory, enter the full filepath, ie C:\Scripts\SCRIPT.PS1'
            Write-Color -foregroundcolor Green $script:Script
            $script:Continue = Read-Host -Prompt "Is this the correct filepath? [Y]es, any key for no"
            Write-Color $script:Continue
            } While(!($script:Continue -eq 'Y'))

        Check-Connection
        if($script:Mode -eq 2) {
            Export-Data
        }

    }

    'Invoke-Line' {
        Write-Color "Warning!"," It is suggested that you test your command against one computer before running it against multiple machines!" -Color yellow,red -backgroundcolor red,black
        $First = 0
        $script:Confirm = $null
        $script:Continue = 'Y'

        do {

            if($First -eq 0) {                
                Write-Color "Please note that this command will be tested against your local computer with a ","-whatif"," appended. Not all commands accept the ","-whatif"," argument, so there may be an error message." -color yellow,blue,yellow,blue,yellow
                Write-Color "This does not mean that there will be an error on the remote computer(s)." -color blue
                $First++
            }
            else {

                do {
                    $script:targetCmd = Read-Host -prompt 'Please enter the precise command you wish to run'
                    Write-Color "This is the potential result of your command on a computer: "-color blue
                    $script:Test = -join ($script:targetCmd, " -whatif")

                    Write-Color "${script:Test}" -color Green
                    Invoke-Expression $script:Test

                    $script:Confirm = Read-Host -prompt "Is this the command you want to run? [Y] for yes, any key for no"
                    $First++
                    Write-Host $script:Confirm

                    if(!($script:Confirm -eq 'Y')) {
                        $script:Continue = read-host -prompt "Would you like to try again? [Y] to retype your command, any key to terminate script"
                        Write-Host $script:Continue
                    
                        if(!($script:Continue -eq 'Y')) {
                            Write-Color "Terminating script." -Color Red
                            exit
                        }
                    }
                    $script:Continue = $null
                } While ($Script:Continue -eq 'Y')
            }
        } While (!($script:Confirm -eq 'Y')) 

        Check-Connection
        if($script:Mode -eq 2) {
            Export-Data
        }
    }
    'Get-Linked' {
        
        Check-Connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }   
    }
    'Generate-CSV' {

        Check-Connection
        if($script:Mode -eq 2) {

            export-data
            append-data
        }

        if($script:Checkpath -ne $null) {
            cleanup
        }
    }
    'Update-Veyon' {
        Get-Hostfile
        Update-Veyon
    }
    'Get-OUMembers' {
        Get-OUMembers
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }  
    'Get-ADinfo' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
    'Cleanup' {
        $script:checkpath = read-host -prompt 'Please enter the filepath for your .csv'
        Write-color "Removing duplicates and blank lines from ","$script:CheckPath" -color yellow,green
        cleanup
    }  
    'Get-Space' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    } 

    'Get-Stats' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }  
    'Get-OS' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
    'Get-BIOS' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
    'Get-CPU' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
    'Get-RAM' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
    'Get-GPU' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
    'Get-Network' {
        check-connection
        if($script:Mode -eq 2) {
            Export-Data
            Append-Data
        }
    }
}

if(($script:Mode -eq 1) -and ($script:Results -ne $null)) {
    $script:ConfirmWrite = Read-Host -prompt "Would you like to save your results to .xlsx? [Y]es or any key for No"
    Write-Host $script:ConfirmWrite

    if($script:ConfirmWrite -eq 'Y') {
        Set-Values
        Export-Data
    }
}

if(($script:task -ne 'Generate-CSV') -and ($script:Checkpath -ne $null)) {
    cleanup
}

if($script:Task -eq 'Generate-CSV') {
    $script:Confirm = Read-Host -Prompt 'Would you like to add your configuration to Veyon? [Y]es or any key for No'

    if($script:Confirm -eq 'Y') {
        $script:hostfile = $script:checkpath
        Update-Veyon
    }   
}
if($script:ConfirmWrite -eq 'Y') {
    Write-Color "Report written to ","${script:CheckPath}" -color blue,cyan
}

if($script:Logging -eq 'Y') {
    Stop-Log
}

Write-Color "Goodbye." -color magenta