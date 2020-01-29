
<#
.Synopsis
   Measure speed of multiple scriptblocks for comparison
.EXAMPLE
   Measure-ScriptBlockSpeed {$null = [math]::PI},{[math]::PI | Out-Null}
   Measures the difference between using $null vs Out-Null by running the
   scriptblocks 10000 times.

   Output example:

   ScriptBlock           RunTime          Result Speed Count
   -----------           -------          ------ ----- -----
   $null = [math]::PI    00:00:01.1045898 1th    -     10000
   [math]::PI | Out-Null 00:00:02.1625823 2th    +96%  10000
.EXAMPLE
   Measure-ScriptBlockSpeed {$null = [math]::PI},{[math]::PI | Out-Null} -Name Null,Out-Null -Count 1000
   Measures the difference between using $null vs Out-Null by running the
   scriptblocks 10000 times. The name of the first block is 'Null' and the second one 'Out-Null'

   Output example:

   ScriptBlock RunTime          Result Speed Count
   ----------- -------          ------ ----- -----
   Null        00:00:00.1308023 1th    -      1000
   Out-Null    00:00:00.2268845 2th    +73%   1000

.NOTES
   Author: Michaja van der Zouwen
#>
function Measure-ScriptBlockSpeed
{
    [CmdletBinding()]
    Param
    (
        # scriptblock(s) to measure
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [scriptblock[]]
        $ScriptBlock,
        
        # Names of the tests for progress and endresult
        [string[]]
        $Name,

        # Nr of times the scriptblock(s) will be executed
        [Parameter(Mandatory=$false, 
                   Position=1)]
        [int]
        $Count = 10000
    )
    If ($PSBoundParameters['Name'])
    {
        $Activity = "Running comparison [$($Name -join ' vs. ')].."
    }
    else
    {
        $Activity = "Running comparison for $($ScriptBlock.Count) ScriptBlocks.."
    }
    $Output = for ($i = 0; $i -lt $ScriptBlock.Count; $i++)
    {
        If ($PSBoundParameters['Name'])
        {
            $TestName = $PSBoundParameters['Name'][$i]
            $Status = "ScriptBlock [$TestName]"
        }
        else
        {
            $TestName = $ScriptBlock[$i].ToString()
            $Status = "ScriptBlock $($i+1)"
        }
        Write-Progress -Activity $Activity -Status $Status -PercentComplete ((100/($ScriptBlock.Count+1) * $i))
        [pscustomobject]@{
            ScriptBlock = $TestName
            RunTime = Measure-Command {
                for ($y = 0; $y -lt $Count; $y++)
                { 
                    #Write-Progress -Activity Processing -Status "ScriptBlock $($i+1)" -PercentComplete ((100/($Count * $ScriptBlock.Count)) * ($y + ($i*$count)))
                    Invoke-Command $ScriptBlock[$i]
                }
            }
            Result = ''
            Speed = ''
            Count = $Count
        }
    }
    $i++
    Write-Progress -Activity $Activity -Status "Calculating results" -PercentComplete ((100/($ScriptBlock.Count+1) * $i))
    $Output = $Output | Sort RunTime
    for ($i = 0; $i -lt $Output.Count; $i++)
    { 
        $Output[$i].Result = "$($i + 1)th"
        switch ($i){
            0       {$Output[$i].Speed = '-'}
            Default {$Output[$i].Speed = '+' + [int](100 / $Output[0].Runtime.Ticks * $Output[$i].Runtime.Ticks - 100) + '%'}
        }
    }

    $Output | Format-Table -AutoSize
}
