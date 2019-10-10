# Please check GitHub repository for updates to that script: https://github.com/bezibaerchen/powershell_scripts

## Script to check given list of servers for installation status of named application
cls
$servers = Get-Content "<path_to_txt"
$apptocheck = "<name_of_software_returned_via_WMI"
$outfile = "<path_to_csv_outputfile>"

Remove-Item $outfile
echo "Servername;Installation Status" | Out-File -FilePath $outfile -Encoding default -Append
$total=$servers.count
$count_installed=0
$count_notinstalled=0
$i=0

foreach ($server in $servers) {
    $i++
    Write-Progress -Activity "Checking installation status..." -Status "Currently working on $server ($i of $total)." -PercentComplete ($i/$servers.Count*100)
    $apppresent = gwmi win32_product -ComputerName $server | ? {$_.name -match $apptocheck}

     if ($apppresent) {
         $found="installed"
         echo "$server;$found"
         echo "$server;$found" | Out-File -FilePath $outfile -Encoding Default -Append
         $count_installed++
     }
    
    else {
        $found="not installed"
        echo "$server;$found"
        echo "$server;$found" | Out-File -FilePath $outfile -Encoding Default -Append
        $count_notinstalled++
    }
    
}

Echo "Checked the total of $total servers. Number of servers with $apptocheck installed: $count_installed. Number of servers with $apptocheck not installed or not connectable: $count_notinstalled"
Echo "Output in CSV can be found at $outfile"