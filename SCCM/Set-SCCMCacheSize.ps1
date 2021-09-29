$NewCacheSize = 51200
$Cache = Get-WmiObject -Namespace 'ROOT\CCM\SoftMgmtAgent' -Class CacheConfig

if($Cache.size -lt $NewCacheSize){
    $Cache.Size = $NewCacheSize
    $Cache.Put() | Out-Null
}

# Optional - Forces the change to happen immediately.
#Restart-Service -Name CcmExec