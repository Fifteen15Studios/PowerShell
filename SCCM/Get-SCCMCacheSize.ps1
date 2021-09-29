$ComputerName = Read-Host "Computer Name "
$Cache = Get-WmiObject -Namespace 'ROOT\CCM\SoftMgmtAgent' -Class CacheConfig

$Cache.Size