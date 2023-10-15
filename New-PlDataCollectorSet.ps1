Function New-PlDataCollectorSet {
    <#
        .SYNOPSIS
            Creates new data collector set from template on remote systems or localhost
 
        .DESCRIPTION
            Creates new data collector set from template on remote systems or localhost
 
        .PARAMETERS
            ComputerNames - single remote computer, array of remote computers, default is localhost
            xmlTemplateName - name of XML Template file, default is first XML file in script folder
            SampleInterval - sets the system polling periodicity in seconds, default is 15 sec
            RotationPeriod - set the rotation period in days, default is 3 days
            StartDataCollector - this is a switch parameter, if it present Data Collector Set start immediatly after creation

         .EXAMPLE
            New-PlDataCollectorSet : creates DCS on the local computer using the first template found in the script's startup folder (Set1.xml coming with script in original release folder)
            New-PlDataCollectorSet -DCSName "DiagSet_3Days_15Sec" -RotationPeriod 3 -SampleInterval 15 -xmlTemplateName Set1.xml -StartUponCreation $false
            New-PlDataCollectorSet -ComputerNames "srv1.contoso.com","srv1.contoso.com" -Credential (Get-Credential)  -DCSName "DiagSet_3Days_15Sec" -RotationPeriod 3 -SampleInterval 15 -xmlTemplateName Set1.xml -StartUponCreation $false
    #>

    [CmdletBinding()]
    param (
        [string[]]$ComputerNames = @("localhost"),
        [parameter(ValueFromPipelineByPropertyName)][string]$DCSName = "Set1",
        [PSCredential]$Credential,
        [string]$xmlTemplateName = ([string[]](Get-ChildItem -Path ".\" -Filter "*.xml").Name)[0],
        [parameter(ValueFromPipelineByPropertyName)][int]$SampleInterval = 15,
        [parameter(ValueFromPipelineByPropertyName)][int]$RotationPeriod = 3,
        [switch]$StartDataCollector,
        [parameter(ValueFromPipelineByPropertyName,DontShow)][xml]$XML
    )

    begin {
        [xml]$xmlTemplate = Get-Content -Path ".\$xmlTemplateName"
        $SessionOptions = New-PSSessionOption -NoMachineProfile -SkipCACheck

        $Action = {
            param( $DataCollectorName, $xml, $Sample, $Rotation, $StartDC )

            # Custimize template by removing some computer-specific nodes or edit nodes with new values according incoming parameters
            "//LatestOutputLocation","//OutputLocation","//Security" | % { try {$xml.ChildNodes.SelectNodes($_) | % {$_.ParentNode.RemoveChild($_)}} catch {} }
            $xml.SelectSingleNode("//Name").'#text' = $DataCollectorName
            $xml.SelectNodes("//*[starts-with(local-name(),'Description')]") | % {$_.'#text' = "This set was created by $env:USERNAME@$env:USERDOMAIN at $((Get-Date).ToString("yyyy.MM.dd-HH:mm:ss"))"}
            $RootPathNode = $xml.SelectSingleNode("//RootPath")
            $RootPathNode.'#text' = $RootPathNode.'#text'.Substring(0,$RootPathNode.'#text'.LastIndexOf("\") + 1) + $DataCollectorName
            $xml.SelectSingleNode("//SampleInterval").'#text' = [string]$Sample
            $xml.SelectSingleNode("//MaxFolderCount").'#text' = [string]$Rotation
            $xml.SelectSingleNode("//Age").'#text' = [string]$Rotation

            # Rewrite 'CounterDisplayName' nodes values with OS System Language Names
            $ENU = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009" -Name "Counter").Counter
            $Current = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage" -Name "Counter").Counter

            $i = 0
            $xml.SelectNodes("//Counter") | % {
                $CounterFull = $_.'#text'.split("\")
                if ($CounterFull[1].IndexOf("(") -ne -1) {
                    $Brackets = $CounterFull[1].Substring($CounterFull[1].IndexOf("("))
                    $CounterPart1 = $CounterFull[1].Replace($Brackets,"")
                } else {
                    $Brackets = ""
                    $CounterPart1 = $CounterFull[1]
                }
    
                $Part1Index = $ENU[$ENU.IndexOf($CounterPart1) - 1]
                $Part1CurrentLanguage = $Current[$Current.IndexOf($Part1Index) + 1] + $Brackets
                if ($CounterFull[2] -eq "*") {
                    $Part2CurrentLanguage = $CounterFull[2]
                } else {
                    $Part2Index = $ENU[$ENU.IndexOf($CounterFull[2]) - 1]
                    $Part2CurrentLanguage = $Current[$Current.IndexOf($Part2Index) + 1]
                }
    
                ($xml.SelectNodes("//CounterDisplayName"))[$i].innertext = $("\" + $Part1CurrentLanguage + "\" + $Part2CurrentLanguage)
                $i++
            }

            $datacollectorset = New-Object -COM Pla.DataCollectorSet
            $datacollectorset.SetXml($xml.OuterXml)

            # Check DataCollector is already exist
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
            }
            while (-Not $done)
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedule)

            if ($tasks | ? {$_.Name -eq $DataCollectorName}) {
                if ($(Read-Host "$DataCollectorName already exist, do you want to overwrite it (y/n)") -eq "y") {
                    logman stop -n $DataCollectorName
                    $sets = New-Object -ComObject Pla.DataCollectorSet
                    $sets.Query($DataCollectorName, $null)
                    $set = $sets.PSObject.Copy()
                    Remove-Item -Path $set.RootPath -Recurse -Force -ErrorAction SilentlyContinue
                    logman delete -n $DataCollectorName

                    $datacollectorset.Commit($DataCollectorName , $null , 0x0003) | Out-Null
                    if ($StartDC) {$datacollectorset.Start($true)}
                } else {"Skip Actions" | Out-Host}
            } else {
                $datacollectorset.Commit($DataCollectorName , $null , 0x0003) | Out-Null
                if ($StartDC) {$datacollectorset.Start($true)}
            }
        }
    }

    Process {
        if ($ComputerNames -ne @("localhost")) {
            $ComputerNames | % {
                $_
                Invoke-Command -Credential $Credentials -ComputerName $_ -SessionOption $SessionOptions -ArgumentList ($DCSName,$xmlTemplate,$SampleInterval,$RotationPeriod,$($StartDataCollector.IsPresent)) -ScriptBlock $Action
            }
        } else { #localhost
            $env:COMPUTERNAME
            Invoke-Command -ArgumentList ($DCSName,$xmlTemplate,$SampleInterval,$RotationPeriod,$($StartDataCollector.IsPresent)) -ScriptBlock $Action
        }
    }
}

Register-ArgumentCompleter -CommandName New-PlDataCollectorSet -ParameterName xmlTemplateName -ScriptBlock {[string[]](Get-ChildItem -Path ".\" -Filter "*.xml").Name}
Register-ArgumentCompleter -CommandName New-PlDataCollectorSet -ParameterName StartUponCreation -ScriptBlock {@($([char]"$" + "true"),$([char]"$" + "false"))}