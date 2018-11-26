## Please check GitHub repository for updates to that script: https://github.com/bezibaerchen/powershell_scripts

## Script to assign Lync / Skype for Business Policies based on group memberships


## Measure duration
$start = Get-Date

## Logging
$outfile="C:\temp\log_policy_assign.txt"
$skippedfile="C:\temp\skipped_users.txt"
$csvall="C:\temp\assigned_policies.csv"

Remove-Item $outfile
Remove-Item $skippedfile
Remove-Item $csvall
Echo "Initializing..." | Out-File -Append -Encoding default $outfile
Echo "Initializing..." | Out-File -Append -Encoding default $skippedfile
$csvheaders="Type;Name;LogonName;mail;Title;Policy before;Assigned Policy;Changed"
Echo $csvheaders | Out-File -Append -Encoding default -FilePath $csvall

## Define E-Mail settings
$smtpServer="<SMTP Server>"
$from = "<from address for mail notification>"
$to = "<recipients> Separate multiple recipients like this: "me@the.net","foo@bar.net""
$subject="[Skype - Policy Assignment] Processing finished"

## Settings needed to execute Lync Commands via Remote Powershell
$username="<username for remote powershell connection"
$pwdTxt = Get-Content "C:\scripts\ps\Lync\cred.txt"
$securePwd = $pwdTxt | ConvertTo-SecureString
$credobject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd
$lync_session = New-PSSession -ConnectionUri "https://<servername for remote powershell>/OcsPowershell" -Credential $credobject
Import-PSSession $lync_session


## Skype settings
$enterprisepolicy="<name of enterprise policy>"
$standardpolicy="<name of standard policy>"

## Helpers
$threshold=5000

## Connect to ARS Proxy
Connect-QADService -Proxy

## Define user groups
$exchangelync="<group 1>"
$mdxlync="<group 2>"

## Reset counters to 0
$exchangecount=0
$mdxcount=0
$exchangeskip=0
$mdxskip=0
$sendmail="no"

## Loop through Exchange Lync users
$exchangegroup=Get-QADGroupMember -Proxy $exchangelync -SizeLimit 9999

echo "Starting assignment for Exchange Users..." | Out-File -FilePath $outfile -Encoding Default -Append
$exchangestart = Get-Date
foreach ($user in $exchangegroup)

{
	## Process next user if threshold isn't reached
    if  ($exchangecount -lt $threshold) {
        ## Get SIP address of current users
		$sip=(Get-QADUser $user -DontUseDefaultIncludedProperties -IncludedProperties msrtcsip-primaryuseraddress)."msrtcsip-primaryuseraddress"
		## Get currently assigned Conferencing Policy
        $setpolicy=(Get-Csuser $sip).ConferencingPolicy
		## Get user details to be included in CSV
        $userdetails = Get-QADUser $user -DontUseDefaultIncludedProperties -IncludedProperties LogonName,Name,mail,Title
		## if current policy isn't the to be assigned one assign it
        if ($setpolicy -ne $enterprisepolicy) {
            ## assign desired policy
			Grant-CsConferencingPolicy –Identity $sip $enterprisepolicy -WhatIf
			## raise counter for changed Exchange users by 1
            $exchangecount=$exchangecount+1
			## Mark changed policy in logfile
            echo "Exchange: Setting policy to $enterprisepolicy for $sip. Policy before:$setpolicy" | Out-File -FilePath $outfile -Encoding Default -Append
			## Build and append details to CSV
            $csvline = "Exchange;" + $userdetails.Name + ";" + $userdetails.LogonName + ";" + $userdetails.mail + ";" + $userdetails.Title + ";" + $setpolicy + ";" + $enterprisepolicy + "; Yes;"
            echo $csvline | Out-File -Append -FilePath $csvall
        }
        ## skip assignment if policy is already the desired one
        else {
            ## mark skipped user in logfile
			echo "Skipping Exchange user $sip as Policy is already $enterprisepolicy" | Out-File -FilePath $skippedfile -Encoding default -Append
			## build and append details to CSV
            $csvline = "Exchange;" + $userdetails.Name + ";" + $userdetails.LogonName + ";" + $userdetails.mail + ";" + $userdetails.Title + ";" + $setpolicy + ";" + $enterprisepolicy + "; No;"
            echo $csvline | Out-File -Append -FilePath $csvall
			## raise counter for skipped Exchange users by 1
            $exchangeskip=$exchangeskip+1
            
        }
           
     ## measure time
		$exchangetime = Get-Date
    }
	## stop processing if threshold is reached
    else {
		## add remark about reached threshold to logfile
        echo "Stopping processing of Exchange users! Threshold reached ($threshold)" | Out-File -FilePath $outfile -Encoding Default -Append
		## change E-Mail content to reflect reached threshold
        $subject="[Skype - Policy Assignment] Error: Threshold reached"
        $body="Error while processing. Please see attachment for details`n`nThreshold set to: $threshold"
		## measure time
		$exchangetime = Get-Date
		## stop further processing
        break
    }
	## Set E-Mail body to successful
    $body="Processing finished. Please see attachment for details"
}


## Loop through MDX Lync users
$mdxgroup=Get-QADGroupMember -Proxy $mdxlync -SizeLimit 9999

echo "Starting assignment for MDX Users..." | Out-File -FilePath $outfile -Encoding Default -Append
$mdxstart = Get-Date
foreach ($user in $mdxgroup)

{
	## Process next user if threshold isn't reached
    if  ($mdxcount -lt $threshold) {
		## Get SIP address of current users
        $sip=(Get-QADUser $user -DontUseDefaultIncludedProperties -IncludedProperties msrtcsip-primaryuseraddress)."msrtcsip-primaryuseraddress"
		## Get currently assigned Conferencing Policy
        $setpolicy=(Get-Csuser $sip).ConferencingPolicy
		## Get user details to be included in CSV
        $userdetails = Get-QADUser $user -DontUseDefaultIncludedProperties -IncludedProperties LogonName,Name,mail,Title
		## if current policy isn't the to be assigned one assign it
        if ($setpolicy -ne $standardpolicy) {
			## assign desired policy
            Grant-CsConferencingPolicy –Identity $sip $standardpolicy -WhatIf
			## raise counter for changed MDX users by 1
            $mdxcount=$mdxcount+1
			## Mark changed policy in logfile
            echo "MDX: Setting policy to $standardpolicy for $sip. Policy before:$setpolicy" | Out-File -FilePath $outfile -Encoding Default -Append
			## Build and append details to CSV
            $csvline = "MDX;" + $userdetails.Name + ";" + $userdetails.LogonName + ";" + $userdetails.mail + ";" + $userdetails.Title + ";" + $setpolicy + ";" + $standardpolicy + "; Yes;"
			echo $csvline | Out-File -Append -FilePath $csvall
        }
		## skip assignment if policy is already the desired one
        else {
            ## mark skipped user in logfile
			echo "Skipping MDX user $sip as Policy is already $standardpolicy"| Out-File -FilePath $skippedfile -Encoding default -Append
			## build and append details to CSV
			$csvline = "MDX;" + $userdetails.Name + ";" + $userdetails.LogonName + ";" + $userdetails.mail + ";" + $userdetails.Title + ";" + $setpolicy + ";" + $standardpolicy + "; No;"
            echo $csvline | Out-File -Append -FilePath $csvall
			## raise counter for skipped MDX users by 1
            $mdxskip=$mdxskip+1
        }
		## measure time
		$mdxtime = Get-Date
    }
	## stop processing if threshold is reached
    else {
		## add remark about reached threshold to logfile
        echo "Stopping processing of MDX users! Threshold reached ($threshold)" | Out-File -FilePath $outfile -Encoding Default -Append
		## change E-Mail content to reflect reached threshold
        $subject="[Skype - Policy Assignment] Error: Threshold reached"
        $body="Error while processing. Please see attachment for details`n`nThreshold set to: $threshold"
		## measure time
		$mdxtime = Get-Date
		## stop further processing
        break
    }
 
## Calculate duration
$end = Get-Date
$exchangeduration = $exchangetime-$exchangestart
$mdxduration = $mdxtime-$mdxstart
$totalduration = $end-$start
$exchangeminutes = [math]::Round($exchangeduration.TotalMinutes)
$exchangeseconds = [math]::Round($exchangeduration.TotalSeconds)
$mdxminutes = [math]::Round($mdxduration.TotalMinutes)
$mdxseconds = [math]::Round($mdxduration.TotalSeconds)
$totaldurationminutes = [math]::Round($totalduration.TotalMinutes)
$totaldurationseconds = [math]::Round($totalduration.TotalSeconds)

## Set mail body to successful and include details
$body="Processing finished - Please see attachment for details.`n`nAdjusted Exchange users: $exchangecount `nSkipped Exchange users: $exchangeskip `nDuration for Exchange adjustments: $exchangeminutes Minutes ( $exchangeseconds Seconds) `n`n`nAdjusted MDX users: $mdxcount `nSkipped MDX user: $mdxskip `nDuration for MDX adjustments: $mdxminutes Minutes ( $mdxseconds Seconds)`n`nOverall script execution time: $totaldurationminutes Minutes ( $totaldurationseconds Seconds )"
    
}

## check if any changes have been made at all
echo "Number of changed MDX accounts: $mdxcount"
echo "Number of changed Exchange accounts: $exchangecount"

## check if any changes have been made. Skip Send E-Mail in case no changes have been detected
if ($mdxcount -gt 0) {
    echo "Greater than 0: $mdxcount"
    $sendmail="yes"
    }
else {
    echo "NOT greater than 0: $mdxcount"
        if ($exchangecount -gt 0) {
            echo "Exchange greater than 0: $exchangecount"
            $sendmail="yes"
        }
        else {
            echo "Exchange NOT greater than 0: $exchangecount"
            $sendmail="no"
        }

}


## send E-Mail to configured recipients with Logfile and CSV attached if changes have been made
if ($sendmail -eq "yes") {
    echo "Sending E-Mail notification to $to" | Out-File -FilePath $outfile -Encoding Default -Append
    Send-Mailmessage -smtpServer $smtpServer -from $from -to $to -subject $subject -body $body -priority High -Attachments $outfile,$csvall
}
else {
    echo "NOT sending E-Mail as no changes have been made" | Out-File -FilePath $outfile -Encoding Default -Append
}

## close session to Skype for Business server
Remove-PSSession $lync_session
echo "Processing finished." | Out-File -FilePath $outfile -Encoding Default -Append
echo "Elapsed time for Exchange users: $exchangeminutes Minutes ( $exchangeseconds Seconds)." | Out-File -FilePath $outfile -Encoding Default -Append
echo "Elapsed time for MDX users: $mdxminutes Minutes ( $mdxseconds Seconds)." | Out-File -FilePath $outfile -Encoding Default -Append
echo "Totally elapsed time: $totaldurationminutes Minutes ( $totaldurationseconds Seconds)." | Out-File -FilePath $outfile -Encoding Default -Append
