<#
.Synopsis
   Run powershell command as admin
.DESCRIPTION
   Run a powershell command with elevated rights without relaunching
   the whole script as admin.
.EXAMPLE
   Invoke-ElevatedCommand -Command Stop-Process -Parameters @{name = 'chrome';force = $true}
.EXAMPLE
   $Params = @{
        name = 'chrome'
        force = $true
   }
   Elevate Stop-Process $Params
.EXAMPLE
   Invoke-ElevatedCommand -Command "ipconfig /flushdns"
#>
function Invoke-ElevatedCommand
{
    [CmdletBinding()]
    [Alias('Elevate')]
    Param
    (
        # Command to run
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]
        $Command,

        # Parameters in hashtable format (like with splatting)
        [Parameter(Position=1)]
        [hashtable]
        $Parameters

    )

    If ($PSBoundParameters['Parameters'])
    {
        $Params = $Parameters.GetEnumerator().foreach{"-$($_.key) '$($_.value)'"} -join ' '
        $Script = "$Command $Params"
    }
    else
    {
        $Script = $Command
    }
    $ProcessParams = @{
        FilePath = 'powershell.exe'
        Verb = 'RunAs'
        ArgumentList = "-command & {$Script | Export-Clixml -Path $env:USERPROFILE\PSoutput.xml}"
        Wait = $true
        WindowStyle = 'Hidden'
    }
    Start-Process @ProcessParams
    If ($Output = Import-Clixml -Path $env:USERPROFILE\PSoutput.xml -ErrorAction SilentlyContinue)
    {
        Remove-Item $env:USERPROFILE\PSoutput.xml
        return $Output
    }
}
