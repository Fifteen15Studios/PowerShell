###############################################################################
#
# Gets all PCs in a domain that are running a server OS. The last line 
#   optionally exports the list to a CSV. To do so, uncomment this line and
#   enter a file path in place of <path>
#
###############################################################################

Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' `
-Properties Name,Operatingsystem,OperatingSystemVersion,IPv4Address |
Sort-Object -Property Operatingsystem |
Select-Object -Property Name,Operatingsystem,OperatingSystemVersion,IPv4Address 
# | Export-Csv "<path>\Servers.csv" -NoTypeInformation -Force