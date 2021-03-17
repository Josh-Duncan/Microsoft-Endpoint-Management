# Last tested with MECM Current Branch 2006
# It is unclear how much backwards compatibility it may have.  Use at your own peril
#
# This script needs to be run with an account that has at least one of the roles: Full Administrator, Infrastructure Administrator, Operations Administrator
# This script does not currently need write permissions to anything.
#
# Josh Duncan
# 2021-03-17
# https://github.com/Josh-Duncan
# https://github.com/Josh-Duncan/MEM_MECM/wiki/Quick-Health-Check.ps1
#
# The bulk of the work came from Tao Yang and his blog
#    https://blog.tyang.org/2011/05/01/powershell-functions-get-ipv4-network-start-and-end-address/
#
# This needs to be run from a MECM PowerShell ISE or PowerShell console

$Boundaries = Get-Content 'C:\temp\Boundary_List.csv'
    # Identify the path to the CSV file used to create the boundaries.
    # CSV Format - Name, Comment, IP Subnet
    #     Seattle Office,Wireless,10.0.2.0/24


function Get-IPrangeStartEnd 
{ 
 
    param (  
      [string]$start,  
      [string]$end,  
      [string]$ip,  
      [string]$mask,  
      [int]$cidr  
    )  
      
    function IP-toINT64 () {  
      param ($ip)  
      
      $octets = $ip.split(".")  
      return [int64]([int64]$octets[0]*16777216 +[int64]$octets[1]*65536 +[int64]$octets[2]*256 +[int64]$octets[3])  
    }  
      
    function INT64-toIP() {  
      param ([int64]$int)  
 
      return (([math]::truncate($int/16777216)).tostring()+"."+([math]::truncate(($int%16777216)/65536)).tostring()+"."+([math]::truncate(($int%65536)/256)).tostring()+"."+([math]::truncate($int%256)).tostring() ) 
    }  
      
    if ($ip) {$ipaddr = [Net.IPAddress]::Parse($ip)}  
    if ($cidr) {$maskaddr = [Net.IPAddress]::Parse((INT64-toIP -int ([convert]::ToInt64(("1"*$cidr+"0"*(32-$cidr)),2)))) }  
    if ($mask) {$maskaddr = [Net.IPAddress]::Parse($mask)}  
    if ($ip) {$networkaddr = new-object net.ipaddress ($maskaddr.address -band $ipaddr.address)}  
    if ($ip) {$broadcastaddr = new-object net.ipaddress (([system.net.ipaddress]::parse("255.255.255.255").address -bxor $maskaddr.address -bor $networkaddr.address))}  
      
    if ($ip) {  
      $startaddr = IP-toINT64 -ip $networkaddr.ipaddresstostring  
      $endaddr = IP-toINT64 -ip $broadcastaddr.ipaddresstostring  
    } else {  
      $startaddr = IP-toINT64 -ip $start  
      $endaddr = IP-toINT64 -ip $end  
    }  
      
     $startrange=INT64-toIP -int $startaddr 
     $endrange=INT64-toIP -int $endaddr 
     $temp="$startrange-$endrange"
    
     return $temp 

}


$Boundaries | % {
    $BoundaryName = $_.Split(",")[0]
    $Comment = $_.Split(",")[1]
    $tempSubnet = $_.Split(",")[2]
    $splitCIDR = $tempSubnet.Split("//")[1]
    $splitIPRange = $tempSubnet.Split("//")[0]
    $tempIPRange = Get-IPrangeStartEnd -ip $splitIPRange -cidr $splitCIDR
    New-CMBoundary -Name "$BoundaryName ($Comment)"  -Type IPRange -Value $tempIPRange 
    }