## Please check GitHub repository for updates to that script: https://github.com/bezibaerchen/powershell_scripts

## Script to determine name and standard gateway for a given list of IP addresses via WMI and output to CSV

## clear screen
cls

## text file containing IP addresses to be checked (one per line)
$check = Get-Content C:\impex\ips.txt

## some counters
$i=0
$notconnectable=0
$gatewaynoname=0
$namenogateway=0
$informationavailable=0
$totalipstocheck=$check.count

## Output file
$outfile = "C:\impex\default_gateways.csv"

## remove possibly existing output file
Remove-Item $outfile -ErrorAction SilentlyContinue

## define headers for CSV
$csvheaders="IP;Name;Gateway;Connectable"

## add headers to output file
echo $csvheaders | Out-File -Append -Encoding default -FilePath $outfile

## Loop through IP addresses and note properties

foreach ($address in $check)
{
$noname=0
$nogateway=0
## progress bar
$i++
Write-Progress -Activity "Query running..." -Status "Trying for $address ($i of $totalipstocheck) Sucessfully queried: $informationavailable No information available: $notconnectable" -PercentComplete ($i/$check.count*100)

	## try to determine name to IP address via WMI. If not possible, set connectable to 0 and name to n/a
    Try
    {
        $name = Get-WmiObject Win32_ComputerSystem -ComputerName $address -ErrorAction Stop | select -ExpandProperty Name
    }
    Catch
    {
        $name = "n/a"
        $noname=1
    }

	## try to determine default gateway via WMI. If not possible, set connectable to 0 and gateway to n/a
    Try
    {
        $gateway = Get-WmiObject win32_networkAdapterConfiguration -ComputerName $address -ErrorAction Stop |  ?{$_.DefaultIPGateway -ne $null} | select -ExpandProperty DefaultIPGateway
    }
    Catch
    {
        $gateway = "n/a"
        $nogateway=1
    }
	
    ## get status and raise counters
    if ($noname -eq "0" -and $nogateway -eq "0")
    {
        $informationavailable++
        $connectable=1
    }
    elseif ($noname -eq "1" -and $nogateway -eq "0")
    {
        $informationavailable++
        $connectable=1
        $gatewaynoname++
    }
    elseif ($noname -eq "0" -and $nogateway -eq "1")
    {
        $informationavailable++
        $connectable=1
        $namenogateway++
    }
    else
    {
        $connectable=0
        $notconnectable++
    }
        
    	
	
	## Build CSV line
    $csvline=$address + ";" + $name + ";" + $gateway + ";" + $connectable
	
	
	
	## append CSV line to output file
    echo $csvline | Out-File -Append -FilePath $outfile
}

## Output statistics to CLI
echo "Statistics`nIPs not connectable: $notconnectable`nIPs with gateway information without name: $gatewaynoname`nIPs with name information without gateway: $namenogateway`nIPs with full information available: $informationavailable"