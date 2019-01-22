## Please check GitHub repository for updates to that script: https://github.com/bezibaerchen/powershell_scripts

## Script to determine name and standard gateway for a given list of IP addresses via WMI and output to CSV

## text file containing IP addresses to be checked (one per line)
$check = Get-Content C:\impex\ips.txt

## set progress counter to 0
$i=0

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

## progress bar
$i++
Write-Progress -Activity "Working on IP addresses" -Status "Currently working on $address." -PercentComplete ($i/$check.count*100)

	## try to determine name to IP address via WMI. If not possible, set connectable to 0 and name to n/a
    Try
    {
        $name = Get-WmiObject Win32_ComputerSystem -ComputerName $address -ErrorAction Stop | select -ExpandProperty Name
        $connectable = 1
    }
    Catch
    {
        $name = "n/a"
        $connectable = 0
    }

	## try to determine default gateway via WMI. If not possible, set connectable to 0 and gateway to n/a
    Try
    {
        $gateway = Get-WmiObject win32_networkAdapterConfiguration -ComputerName $address -ErrorAction Stop |  ?{$_.DefaultIPGateway -ne $null} | select -ExpandProperty DefaultIPGateway
        $connectable = 1
    }
    Catch
    {
        $gateway = "n/a"
        $connectable = 0
    }
	## Build CSV line
    $csvline=$address + ";" + $name + ";" + $gateway + ";" + $connectable
	
	## append CSV line to output file
    echo $csvline | Out-File -Append -FilePath $outfile
}