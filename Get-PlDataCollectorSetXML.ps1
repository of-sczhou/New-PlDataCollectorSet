Function Get-PlDataCollectorSetXML {
    <#
        .SYNOPSIS
            Export existing data collector set fromm remote computer or local host to Xml file
 
        .DESCRIPTION
            Export existing data collector set fromm remote computer or local host to Xml file
 
        .PARAMETERS
            ComputerName - target computer
            DCSName - Name of data collector set

         .EXAMPLE
            Get-PlDataCollectorSetXML : Lists existing Data Collector Sets on local computer
            Get-PlDataCollectorSetXML -ComputerName srv.contoso.com -$Credential (Get-Credential) -Name Set1
            Get-PlDataCollectorSetXML -ComputerName srv1.contoso.com -$Credential (Get-Credential) -Name Set1 | New-PlDataCollectorSet -ComputerName srv2.contoso.com -$Credential (Get-Credential)

            /Copy DataCollector set from one computer to another
            Get-PlDataCollectorSetXML -Name Set1 | New-PlDataCollectorSet -ComputerName srv.contoso.com -$Credential (Get-Credential)
            Get-PlDataCollectorSetXML -ComputerName srv1.contoso.com -$Credential (Get-Credential) -Name Set1 | New-PlDataCollectorSet -ComputerName srv2.contoso.com -$Credential (Get-Credential)

            /Copy DataCollector set from one computer to another with new name and parameters
            Get-PlDataCollectorSetXML -ComputerName srv1.contoso.com -$Credential (Get-Credential) -Name Set1 | New-PlDataCollectorSet -ComputerName srv2.contoso.com -$Credential (Get-Credential) -DCSName "Set2" -SampleInterval 3 -RotationPeriod 5
    #>

    [CmdletBinding()]
    param (
        [String]$ComputerName = "localhost",
        [PSCredential]$Credential,
        [string]$DCSName
    )

    $Action = {
        param($DataCollectorName)

        $schedule = New-Object -ComObject "Schedule.Service"
        $schedule.Connect()
        $folder = $schedule.GetFolder("Microsoft\Windows\PLA")
        $tasks = @()
        $tasknumber = 0
        $done = $false
        do {
            try {
                $task = $folder.GetTasks($tasknumber)
                $tasknumber++
                if ($task) {
                    $tasks += $task
                }
            }
            catch {
                $done = $true
            }
        } while (-Not $done)
        $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedule)

        if ($DataCollectorName) {
            [__ComObject]$Task = ($tasks | ? {$_.Name -eq $DataCollectorName})
            if ($Task) {
                $DCSXml = ([xml]($Task.Xml)).Task.Data.'#cdata-section'
                $SampleInterval = [int](([xml]$DCSXml).SelectSingleNode("//SampleInterval").'#text')
                $RotationPeriod = [int](([xml]$DCSXml).SelectSingleNode("//MaxFolderCount").'#text')
                [pscustomobject] @{
                    DCSName = $DataCollectorName
                    SampleInterval = $SampleInterval
                    RotationPeriod = $RotationPeriod
                    XML = $DCSXml
                }
            } else {
                $false,$false,$false,$false
            }
        } else {
            $tasks.Name
        }
    }

    if ($ComputerName -ne @("localhost")) {
        $Result = Invoke-Command -Credential $Credentials -ComputerName $ComputerName -SessionOption $SessionOptions -ArgumentList ($DCSName) -ScriptBlock $Action
    } else { #localhost
        $Result = Invoke-Command -ArgumentList ($DCSName) -ScriptBlock $Action
    }

    If ($Result) {
        if (-Not $DCSName) {
            Write-Host "List of collectors:"
            $Result
            Write-Host "`nSelect one and specify it in the parameter -DCSName"
        } else {
            [pscustomobject] @{
              DCSName = $Result.DCSName
              SampleInterval = $Result.SampleInterval
              RotationPeriod = $Result.RotationPeriod
              XML = $Result.XML
            }
        }
    } else {
        Write-Host "$DCSName not found on $ComputerName"
    }
}