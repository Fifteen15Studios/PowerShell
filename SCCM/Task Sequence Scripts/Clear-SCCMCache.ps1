# Purpose:
# Clear SCCM cache. Useful when installing a lot of applications during a task
# sequence, or when installing applications which require a large download.

$resman = new-object -com "UIResource.UIResourceMgr"
$cacheInfo = $resman.GetCacheInfo()

#Enum Cache elements and delete each item
$cacheinfo.GetCacheElements()  |
foreach
{
    If ($_.ContentSize -gt 0.00) { 
        $cacheInfo.DeleteCacheElement($_.CacheElementID)
    }
}
