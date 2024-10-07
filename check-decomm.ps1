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

#Variables ===================================================================================================================

$script:results = @()
$script:BothResults = [System.Collections.ArrayList]@()
$script:dnsServers = ("WinAD01.shsu.edu","WinAD02.shsu.edu")
$directory = "C:\Transcripts\"

#Functions ====================================================================================================================

function pullRvs {
    param([string]$server)
    
    Write-Color "Pulling current"," reverse DNS entries"," from ","$($server)" -color blue,green,blue,cyan

    $dnszones = get-dnsserverzone -computername $server | ? zonename -like '*.in-addr.arpa'
    $script:revResults = foreach($zone in $dnszones) {
        $ZoneName = $zone.zonename
        $zoneBreak = $zone.zonename -split '\.'
        $SimpleZone = "$($zonebreak[1]).$($zonebreak[0])"
        $data = Get-DnsServerResourceRecord -computer $server -zonename $zoneName -RRType Ptr

        foreach($i in $data) {
            $hosts = $i.hostname -split '\.'
            $redoHosts = "$($hosts[1]).$($hosts[0])"

            $i | add-member -membertype NoteProperty -name 'ZoneName' -value $ZoneName
            $i | add-member -MemberType NoteProperty -name 'NewZone' -value $simpleZone
            $i | add-member -membertype NoteProperty -name 'NewHosts' -value $redohosts

            $i | select zonename,newzone,newhosts,@{n='RecordData';e={$_.RecordData.PtrDomainName}}, RecordType
        }
    }
    return $script:revResults
}

function fwdrvsNSlookup {
	param([string]$comp)

    $script:inADFlag = $true
    $script:hasnsFwdFlag = $true
    $script:hasnsRvsFlag = $true
    

    $script:inAD = dsquery computer -name $comp

    if($script:inAD -eq $null) {
        $script:inADFlag = $false
        Write-Color "$($comp)"," not found in AD" -color blue,green

        $script:nsLookup = resolve-dnsname $comp -erroraction silentlycontinue

        if($script:nslookup -eq $null) {
            Write-Color "$($comp)"," not found in forward nslookup" -color blue,green
            $script:hasnsFwdFlag = $false
            $searchstr = -join($comp,"*")
            $script:reverse = $script:revResults | where-object -property 'RecordData' -like $searchstr
            
            if($script:reverse -eq $null) {
                Write-Color "$($comp)"," not found in reverse nslookup" -color blue,green
                $script:hasnsRvsFlag = $false
            }
            
            else {
                Write-Color "$($comp)"," found in reverse nslookup" -color yellow,red
            }
        }

        else{
            Write-Color "$($comp)"," found in forward nslookup" -color yellow,red
        }
    }

    else{
        Write-Color "$($comp)"," found in AD" -color yellow,red
        $script:nsLookup = resolve-dnsname $comp -erroraction silentlycontinue

        if($script:nslookup -eq $null) {
            Write-Color "$($comp)"," not found in forward nslookup" -color blue,green
            $script:hasnsFwdFlag = $false
        }

        else{
            Write-Color "$($comp)"," found in forward nslookup" -color yellow,red
        }

        $searchstr = -join($comp,"*")
        $script:reverse = $script:revResults | where-object -property 'RecordData' -like $searchstr
        
        if($script:reverse -eq $null) {
            Write-Color "$($comp)"," not found in reverse nslookup" -color blue,green
            $script:hasnsRvsFlag = $false
        }

        else{
            Write-Color "$($comp)"," found in reverse nslookup" -color yellow,red
        }
    }

    if($script:inad -ne $null) {
        $tmpOU = $script:inad.split(',')[-5].split('=')[1]
    }

    $fwdIP = $script:nslookup.ipaddress
    $rvsIP = [System.Collections.ArrayList]@()

    foreach($i in $script:reverse) {
        $tmpVar = $i.newzone
        $tmpVar = -join($tmpVar,".")
        $tmpVar = -join($tmpVar,$i.newhosts)
        $i | add-member -membertype NoteProperty -Name 'IP' $tmpVar > $null
        $rvsIP.add($i.ip) > $null

    }
    $tmpRvs = ""
    $count = 0
    foreach($i in $rvsIP) {
        if($count -eq 0) {
            $tmpRvs = $i
        }
        else{
            $tmpRvs = -join($tmprvs,", $($i)")
        }
        $count++
    }
    
    $adcheck = $script:inADFlag
    if($script:inADFlag -eq $true) {
        $adcheck = -join($adcheck,", $($tmpOU) OU")
    }

    $script:res = [pscustomobject]@{
        Node = $comp
        "In AD" = $ADcheck
        "in DNS (forward)" = $script:hasnsFwdFlag
        "Forward Value" = $fwdip
        "in DNS (reverse)" = $script:hasnsRvsFlag
        "Reverse Value" = $tmprvs

    }

    $script:results += $script:res
}

function Get-Hostfile {
    $check = $false
    $current = pwd
    $script:nodelist = @()

    do {

      Write-color "Current directory is ","$current" -color green,blue        
        
      $script:Hostfile  = Read-Host -Prompt "Please enter your hostfile"
      

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

    $script:info = gc $script:Hostfile

    foreach($i in $script:info) {
      $script:nodelist  += $i
    }

    $script:Length = (get-content $script:Hostfile).length
}

function get-filename {
    Write-Color "Would you like to save your output to ","${Directory}" -Color Yellow,Green
    $Confirm = Read-Host -Prompt '[Y]es or any key for No'
    Write-Color "${Confirm}" -Color Cyan

    if(!($Confirm -eq 'Y')) {
        $Directory = Read-Host -Prompt "Where would you like to save your output?"
        Write-Color "${Directory}" -Color Cyan
    }

    $repChar = '_'
    $tmpName = Read-Host -Prompt 'Please enter report title; note that spaces will be replaced by undescores'
    $script:Report = $tmpName -replace ' ', $repChar
    $tmpName = Read-Host -Prompt 'Please enter a task or query name; note that spaces will be replaced by undescores (ex: Get Stats becomes Get_Stats)'
    $TaskName = $tmpname -replace ' ', $repChar
    $TaskName = $TaskName -replace ':',''
  
    $script:CheckPath = "${Directory}\${Report}_${TaskName}.xlsx"
      
    write-host "$script:checkpath"
      
    if(test-path $script:Checkpath) {

        Write-Color "A file already exists by this name. Would you like to ","[O]","verwrite or have the script automatically ","[R]","ename the new file?" -Color Yellow,Red,Yellow,Blue,Yellow,Green,Yellow
        $ChkChoice = Read-Host -prompt "[O]verwrite, Automatically [R]ename, or any other key to cancel and end the script"

        switch($script:chkchoice) {
            'O' {
                remove-item -path $script:checkpath -force
            }
            'R' {
                $script:CheckPath = "${Directory}\New_${Report}_${TaskName}.xlsx"
                write-host "$script:checkpath"
            }
        }
    }
     
  Write-Color -foregroundcolor Green "Report will be saved as ${Report}_${Task}.xlsx in directory ${Directory}"  
  return $script:checkpath
}

function write-excel {

    $trash = [pscustomobject]@{
        This = 'is just'
        The = 'hacky'
        Garbage = 'to delete'
        Query = 'TRASH'
        }

    $script:xlpkg = $script:trash   | export-excel -path $script:checkpath -worksheetname 'TRASH'   -tablename 'T_T'     -autosize -passthru 
    $script:xlpkg = $script:Results | Export-excel -excelpackage $xlpkg    -worksheetname 'Results' -tablename 'Results' -autosize -passthru

    $script:xlpkg.workbook.worksheets.delete('TRASH')  

    close-excelpackage $script:xlpkg
}

#Main program =================================================================================================================

Write-Color "Gathering information. This may take a few moments." -color cyan

foreach($srv in $script:dnsServers) {
    $tmp = pullRvs($srv)
    $script:BothResults.add($tmp) > $null
}

if($script:BothResults -ne $null) {
    Write-Color "Reverse Entry List has been generated ","successfully." -color blue,green
    Write-Color "Checking $($script:dnsservers[0]) against $($script:dnsservers[1])" -color yellow

    $matchCheck = compare-object -reference $script:BothResults[0] -difference $script:bothResults[1]

    if($matchCheck -ne $null) {
        Write-Color "Warning!"," Mismatches found between $($script:dnsservers)! ","Please check your script and connection and try again." -color red,yellow,blue
        exit
    }
    
    else {
        Write-Color "DNS Entries match."," Collecting Hostfile." -color green,blue
    }
}

else{
    Write-Color "Warning!"," No data detected! Please check your script and connection and try again!" -color red,yellow
}

Get-Hostfile
$script:count = 1
foreach($n in $script:nodelist) {
    Write-Color "Working on ","${$n}"," - ","${script:Count}"," out of ","${script:Length}" -Color Cyan, Green, Cyan, Yellow, Cyan, Yellow
    fwdrvsNSlookup($n)
    $script:count++
}

$results | out-gridview

get-filename
write-excel