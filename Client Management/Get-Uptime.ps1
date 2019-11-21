$uptime = ((get-date) - (gcim Win32_OperatingSystem).LastBootUpTime)

$Output = ""

if($uptime.Days -gt 1)
{$Output += $uptime.Days.ToString() + " Days "}
elseif($uptime.Days -eq 1)
{$Output += "1 Day "}


if($uptime.Hours -gt 1)
{$Output += $uptime.Hours.ToString() + " Hours "}
elseif($uptime.Hours -eq 1)
{$Output += "1 Hour "}

if($uptime.Minutes -gt 1) 
{$Output += $uptime.Minutes.ToString() + " Minutes "}
elseif($uptime.Minutes -eq 1) 
{$Output += "1 Minute "}

if($uptime.Seconds -gt 1)
{$Output += $uptime.Seconds.ToString() + " Seconds "}
elseif($uptime.Seconds -eq 1)
{$Output += "1 Second "}

"Uptime: $($Output.trim())"
