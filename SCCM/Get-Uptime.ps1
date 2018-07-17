$uptime = ((get-date) - (gcim Win32_OperatingSystem).LastBootUpTime)

$Output = ""

if($uptime.Days -gt 0)
{$Output += $uptime.Days.ToString() + " Days " }

if($uptime.Hours -gt 0)
{$Output += $uptime.Hours.ToString() + " Hours "}

if($uptime.Minutes -gt 0) 
{$Output += $uptime.Minutes.ToString() + " Minutes "}

if($uptime.Seconds -gt 0)
{$Output += $uptime.Seconds.ToString() + " Seconds"}

"Uptime: $Output"
