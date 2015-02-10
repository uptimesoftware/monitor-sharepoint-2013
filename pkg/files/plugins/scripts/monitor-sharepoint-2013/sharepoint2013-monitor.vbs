Set objWSH =  CreateObject("WScript.Shell")
Set WshSysEnv = objWSH.Environment("Process") 
strComputer = WshSysEnv("UPTIME_HOSTNAME")
strUser = WshSysEnv("UPTIME_USERNAME")
strPassword = WshSysEnv("UPTIME_PASSWORD")

On Error Resume Next


Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, "\root\CIMV2", strUser, strPassword, "ms_409")
objSWbemServices.Security_.ImpersonationLevel = 3

if objSWbemServices Is Nothing Then
	WScript.Echo "Access Denied! - Check your username & password"
	WScript.Quit 1
End If
'Pull some useful search counters
Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MicrosoftWindowsSharePointMicrosoftSharePointFoundation4_SharePointFoundation WHERE NOT Name LIKE '%W3SVC%'", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)
'commented out the not so useful ones
For Each objItem In colItems
	WScript.Echo objItem.Name & ".ActiveHeapCount " & objItem.ActiveHeapCount
	WScript.Echo objItem.Name & ".ActiveThreads " & objItem.ActiveThreads
	WScript.Echo objItem.Name & ".CurrentPageRequests " & objItem.CurrentPageRequests
	WScript.Echo objItem.Name & ".ExecutingSqlQueries " & objItem.ExecutingSqlQueries
	WScript.Echo objItem.Name & ".ExecutingTimePerPageRequest " & objItem.ExecutingTimePerPageRequest
	WScript.Echo objItem.Name & ".HealthScore: " & objItem.HealthScore
	WScript.Echo objItem.Name & ".IncomingPageRequestsRate " & objItem.IncomingPageRequestsRate
'	WScript.Echo objItem.Name & ".ObjectCacheAlwaysLiveCount " & objItem.ObjectCacheAlwaysLiveCount
'	WScript.Echo objItem.Name & ".ObjectCacheAlwaysLiveSize " & objItem.ObjectCacheAlwaysLiveSize
'	WScript.Echo objItem.Name & ".ObjectCacheExpiredCount " & objItem.ObjectCacheExpiredCount
'	WScript.Echo objItem.Name & ".ObjectCacheExpiredSize " & objItem.ObjectCacheExpiredSize
'	WScript.Echo objItem.Name & ".ObjectCacheHitCount " & objItem.ObjectCacheHitCount
'	WScript.Echo objItem.Name & ".ObjectCacheLiveSize " & objItem.ObjectCacheLiveSize
'	WScript.Echo objItem.Name & ".ObjectCacheMissCount " & objItem.ObjectCacheMissCount
	WScript.Echo objItem.Name & ".RejectPageRequestsRate " & objItem.RejectPageRequestsRate
	WScript.Echo objItem.Name & ".RespondedPageRequestsRate " & objItem.RespondedPageRequestsRate
	WScript.Echo objItem.Name & ".SqlQueryExecutingtime " & objItem.SqlQueryExecutingtime
'	WScript.Echo objItem.Name & ".TemplateCacheAverageRecordAge " & objItem.TemplateCacheAverageRecordAge
'	WScript.Echo objItem.Name & ".TemplateCacheLowMemoryTrimCount " & objItem.TemplateCacheLowMemoryTrimCount
'	WScript.Echo objItem.Name & ".TemplateCacheMaxSize " & objItem.TemplateCacheMaxSize
'	WScript.Echo objItem.Name & ".TemplateCacheRapidGrowthTrimCount " & objItem.TemplateCacheRapidGrowthTrimCount
'	WScript.Echo objItem.Name & ".TemplateCacheScheduledTrimCount " & objItem.TemplateCacheScheduledTrimCount
'	WScript.Echo objItem.Name & ".TemplateCacheSize " & objItem.TemplateCacheSize
'	WScript.Echo objItem.Name & ".TemplateCacheTotalHitCount " & objItem.TemplateCacheTotalHitCount
'	WScript.Echo objItem.Name & ".TemplateCacheTotalMissCount " & objItem.TemplateCacheTotalMissCount
'	WScript.Echo objItem.Name & ".TemplateCacheXMLHitCount " & objItem.TemplateCacheXMLHitCount
'	WScript.Echo objItem.Name & ".TemplateCacheXMLMissCount " & objItem.TemplateCacheXMLMissCount
	WScript.Echo
Next