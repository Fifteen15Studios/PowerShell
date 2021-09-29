$resman = new-object -com "UIResource.UIResourceMgr"
$cacheInfo = $resman.GetCacheInfo()

$cacheinfo.GetCacheElements()  |
foreach
{
    #$_.ContentID is packageID
    If ($_.ContentSize -gt 0.00) {
        $cacheInfo.DeleteCacheElement($_.CacheElementID)
    }
}
