
$RunBank = @{}
$ComputersToCheck = "
Computer1,
Computer2"

foreach ($computer in $ComputersToCheck) {
    $Invoked = Invoke-Command -ComputerName $computer -ScriptBlock {
        $TargetTask = Get-ScheduledTask  | Where TaskName -Like "*filter*"
        $TargetTaskInfo = $TargetTask | Get-ScheduledTaskInfo
        $ComputerName = hostname
        $LastRunTime = $TargetTaskInfo | select -exp LastRunTime
        $NextRunTime = $TargetTaskInfo | select -exp NextRunTime
        $TaskState = $TargetTask | select -exp State
        $Output = [PSCustomObject]@{
            Computer = $ComputerName
            TaskState = $TaskState
            LastRunTime = $LastRunTime
            NextRunTime = $NextRunTime
        }
        return $Output
    }
    [array]$RunBank += $Invoked
}
$RunBank | ft

$table = $RunBank | select Computer, TaskState, LastRunTime, NextRunTime | ConvertTo-Html -Fragment
# HTML formatting
$Title = "Task Status: $(Get-Date)"
$ReportDescription = "Current task status is under TaskState."
$Head = @"
<Title>$Title</Title>
<style>
body { background-color: #white; font-family: Segoe UI, Sans-Serif; font-size: 11pt; }
td, th, table { border:1px solid grey; border-collapse:collapse; }
h1, h2, h3, h4, h5, h6 { font-family Segoe UI, Segoe UI Light, Sans-Serif; font-weight: lighter; }
h1 { font-size: 26pt; }
h4 { font-size: 14pt; }
th { color: #383838; background-color: lightgrey; text-align: left; }
table, tr, td, th { padding: 2px; margin: 0px; }
table { width: 95%; margin-left: 5px; margin-bottom: 20px; }
</style>
<h1>$Title</h1>
<h4>$ReportDescription</h4>
"@
$TaskStateRunning = '(?s)<td>Running</td>'
$TaskStateRunningFormatted = '<td><strong>Running</strong></td>'
$table = $table -replace $TaskStateRunning, $TaskStateRunningFormatted

ConvertTo-Html -Head $Head -Body $table | Out-File "path\report.htm" -Encoding ascii -Force
