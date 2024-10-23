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

#Functions ===================================================================================================================

function get-filename {
    $script:Directory = "C:\Transcripts\"
    Write-Color "Would you like to save your output to ","${script:Directory}" -Color Yellow,Green
    $Confirm = Read-Host -Prompt '[Y]es or any key for No'
    Write-Color "${Confirm}" -Color Cyan
    
    if(!($Confirm -eq 'Y')) {
        $script:Directory = Read-Host -Prompt "Where would you like to save your output?"
        Write-Color "${Directory}" -Color Cyan
    }
    $tmpname = set-room

    $repChar = '_'

    if($script:room -ne $null) {
        Write-Color "Would you like to name your report $($tmpname).xlsx?"
        $response = read-host -prompt "[Y]es or any key for no"
        if($response -like 'Y') {
            $tmpname = $script:room
        }
        else {
            $tmpName = Read-Host -Prompt 'Please enter report title; note that spaces will be replaced by undescores'    
        }
    }
    else {
        $tmpName = Read-Host -Prompt 'Please enter report title; note that spaces will be replaced by undescores'
    }
    $script:Report = $tmpName -replace ' ', $repChar
    $script:Directory = $script:Directory.trimend("\")

  
    $script:CheckPath = "${script:Directory}\${script:Report}.xlsx"
      
    write-host "$script:checkpath"
      
    if(test-path $script:Checkpath) {

        Write-Color "A file already exists by this name!" -color red
        Write-Color "`tWould you like to ","[O]","verwrite ","[A]","ppend, ","or have the script ","[R]","ename"," the new file?" -Color Blue,Red,Yellow,red,yellow,blue,cyan,green,blue
        $script:ChkChoice = Read-Host -prompt "[O]verwrite, [R]ename, [A]ppend, or any other key to cancel and end the script"

        switch($script:chkchoice) {
            'O' {
                remove-item -path $script:checkpath -force
                Write-Color "Overwriting ","${script:checkpath}" -color red,yellow
            }

            'R' {
                $script:updateTime = get-date -format "MM-dd-yy"
                $script:Report = -join($script:Report,"_")
                $script:Report = -join($script:Report,$script:updateTime)
                $script:CheckPath = "${script:Directory}\${script:Report}.xlsx"
                write-color "New Filename: ","${script:checkpath}" -color cyan,green
            }
            'A'{
                Write-Color "Will attempt to -append, please be aware this may cause problems."
            }

        }
    }
     
  Write-Color -foregroundcolor Green "Report will be saved as ${script:Report}.xlsx in directory ${script:Directory}"  
  return $script:checkpath
}

function Set-Room {
    $script:updateTime = get-date -format "MM_dd_"
    Write-Color "Auto date value is currently ","${script:updateTime}" -color blue,cyan
    do {
        Write-color "Would you like to use the ","[D]","ate ","(${script:UpdateTime})"," or enter a ","[C]","ustom Location?" -color blue,cyan,blue,cyan,blue,magenta,green
        $script:Choice = read-host -prompt "[D]ate or [C]ustom Location"
        } while(($script:Choice -ne 'D') -and ($script:Choice -ne 'C'))
    if($script:choice -eq 'C') {
        $script:Room = Read-Host -Prompt "Please enter Location"
    }
    else {
        $script:room = $script:updateTime
    }
    return $script:room
}


function write-excel {

    if(test-path $script:checkpath) {

        Write-Color "${script:checkpath} already exists!" -color yellow
        write-color "Attempting to Append" -color blue
    }
   
  $script:xlpkg = $script:trash       | export-excel -path $script:checkpath            -worksheetname 'TRASH'        -tablename 'T_T'              -autosize -passthru 
  $script:xlpkg = $script:AD_Info     | export-excel -excelpackage $xlpkg       -append -worksheetname 'AD_Info'      -tablename 'AD_Info'          -autosize -passthru
  $script:xlpkg = $script:OS          | Export-excel -excelpackage $xlpkg       -append -worksheetname 'OS'           -tablename 'Get_OS'           -autosize -passthru -MoveToStart
  

  $script:xlpkg.workbook.worksheets.delete('TRASH')
  set-excelrange    -worksheet $script:xlpkg.OS -Range "H:H" -wraptext
  set-excelrange    -worksheet $script:xlpkg.OS -Range "Q:Q" -wraptext
      close-excelpackage $script:xlpkg
    
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
    $model = get-computerinfo -property csmodel

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
        Model = $model.csmodel
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

      $script:OS = $script:res
      return $script:os | out-null
}

function Get-ADinfo {
    
    param ( [string]$comp)

    if($script:Err -ne $null) {

        $status = $script:Err
    }
    else {
        $status = "Successful"
    }

    $script:Obj = get-adcomputer -filter {name -like $comp} -properties canonicalname,memberof

    if($script:obj -ne $null) {
        $cname   = $script:Obj.canonicalName
        $hname  = $script:Obj.Name
        $groups = $script:Obj.MemberOf -replace '^CN=([^,]+).+$','$1'

        if($ou -eq $hname) {

                $ou = ($script:Obj.distinguishedName).split(',')[-5].split('=')[1]
            }
            $script:all_groups = "Groups:"

        foreach($g in $groups) {
          $script:all_groups = -join($script:all_groups,"`r`t${g}; ")
           
        } 
        $script:all_groups = $script:all_groups.trimend("; ")

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
            Query = 'Get-ADinfo'
            Desc = $script:adDesc
            OS = $script:ados
            OU = $script:ou
            Groups = $all_groups
            NSlookup = $nsforward.IPAddress
            ReverseNS = $reverse.Namehost
            Created = $script:inad.created
            Last_Logon = $script:inad.lastlogondate
            Status = $status
         }
        }
        else {
            $status = 'Unable to find in AD'

            $script:Res = [PSCustomObject]@{
                Room = $script:Room
                Host = $script:Node
                Query = 'Get-ADinfo'
                Desc = ''
                OS = ''
                OU = ''
                Groups = ''
                NSlookup = ''
                ReverseNS = ''
                Created = ''
                Last_Logon = ''
                Status = $status
            }

        }

        $script:AD_Info += $script:res   
}

function run-stats {
      param(
  [string]$comp)

  $script:Task = 'Get-OS'
  get-OS -comp $comp

  $script:Task = 'Get-ADinfo'
  get-adinfo -comp $comp

  $script:Task = 'Run-All'

  $trash = [pscustomobject]@{
      This = 'is just'
      The = 'hacky'
      Garbage = 'to delete'
      Query = 'TRASH'
  }

  $all = @(
    $trash  
    $script:OS
    $script:ad_info
    
    )
  $script:allstats += $all    
  return $script:allstats | out-null
}

#Main program =================================================================================================================
$script:psver = $psversiontable.PSVersion
if (!($script:psver.Major -gt 6)) {

    Write-color "Warning! ","This script was written for PowerShell version 7 or greater." -color red,yellow
    Write-Color "You have version ","${script:psver}" -color yellow,red
    Write-Color "Some formulas and commands may not behave as expected. " -color yellow
    Write-Color "For the best experience, please make sure you run this script in Powershell version 7.0 or greater. " -color cyan
}

$comp = hostname

get-filename

Write-Color "Gathering data on ","$comp . . ." -color blue,green

run-stats -comp $comp

write-excel