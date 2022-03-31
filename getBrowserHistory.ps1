param 
(
    [Parameter(Mandatory=$true,HelpMessage="Enter the username of the user whose searches you're looking for.")]
    $username,
    [Parameter(HelpMessage="Enter a keyword to search on in a url in browser history. This is set to wildcard before and after your keyword.")]
    [string]$search,
    [Parameter(HelpMessage="This is a switch, just include it if you'd like only Google searches. This will return the searches and sites that have been referred to from Google.")]
    [switch]$googleSearches = $false,
    [Parameter(HelpMessage="Options include All, IE, Edge, and Chrome. All will gather all 3")]
    [ValidateSet("All", "IE", "Edge", "Chrome")]
    $browser = 'All',
    [Parameter(HelpMessage="Provide a start date and time for a range you'd like to search. ex. 2/2/2022 3:15. Please use UTC.")]
    [Datetime]$startDate,
    [Parameter(HelpMessage="Provide a end date and time for a range you'd like to search. ex.2/2/2022 13:15. Please use UTC.")]
    [Datetime]$endDate,
    [Parameter(HelpMessage="Provide integer parameter for the number of hours from current time you'd like to search in history")]
    [int]$lastHours = 0,
    [Parameter(HelpMessage="Provide path for result output. This will append a folder Browserhistory within the path provided")]
    $outputPath = "c:\temp\"
)

#global variable for chromium browser history locations
$chrome = "C:\Users\$username\AppData\Local\Google\Chrome\User Data\Default\History"
$edge ="C:\Users\$username\AppData\Local\Microsoft\Edge\User Data\Default\History"

Write-Host $username
Write-Host $search
Write-Host $googleSearches
Write-Host $browser
Write-Host $startDate
Write-Host $endDate
Write-Host $lastHours

#check if output path was provided with \ as the last character before appending our folder name
if ($outputPath.Substring($outputPath.Length - 1) -eq '\')
{
    $outputPath = "$($outputPath)browserhistory\"
}
else 
{
    "$($outputPath)\browserhistory\"
}

#clean up from previous runs
if (Test-Path $outputPath)
{
    Remove-Item $outputPath

}
if (Test-Path "$outputPath.zip")
{
    Remove-Item "$outputPath.zip"
}

#function for use with psobject creation to check if value exists else set to null
function checkExists($variable) 
{
    if ($variable)
    {
        return $variable
    }
    else
    {
        return "null"
    }
}

#load sqlite DLL
function loadDLL()
{
    $sqliteDllBase64 = ""
    $sqliteinteropDllBase64 = ""
    $Content = [System.Convert]::FromBase64String($sqliteDllBase64)
    $interopContent = [System.Convert]::FromBase64String($sqliteinteropDllBase64)
    Set-Content -Path "C:\temp\System.Data.SQLite.dll" -Value $Content -Encoding Byte
    Set-Content -Path "C:\temp\SQLite.Interop.dll" -Value $interopContent -Encoding Byte
    Add-Type -Path "C:\temp\System.Data.SQLite.dll"
}


#convert chromium sqlite timestamp storage from webkit to utc time
function convertWebKitToUTC($time)
{
    return (([System.DateTimeOffset]::FromUnixTimeSeconds(($time/1000000 - 11644473600))).DateTime)
}

#utilized for both Chrome/Edge just need a different db
function getChromiumHistory($browser)
{
    $cpFile = "$browser-cp"
    Copy-Item -Path $browser -Destination $cpFile
    if (Test-Path $cpFile)
    {
        $history = @()
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=$cpFile"
        $con.Open()
        $sql = $con.CreateCommand()
        if ([string]::IsNullOrEmpty($search) -and !$googleSearches)
        {
            $query = "select * from urls;"
            $sql.CommandText = $query
        }
        else 
        {
            if ($googleSearches) 
            {
                $query = "select * from urls where url like '%google%';"
                $sql.CommandText = $query
            }
            else 
            {
               $query = "select * from urls where url like '%$search%';"
               $sql.CommandText = $query
            }
        }
        $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
        $data = New-Object System.Data.DataSet
        [void]$adapter.Fill($data)
        if ($googleSearches)
        {
            $entries = $data.Tables[0]
            if ($lastHours -gt 0)
            {
                write-host 'searching google last hours'
                $entries = $entries | where {(convertWebKitToUTC($_.last_visit_time)) -GE ((Get-Date).AddHours(-$lastHours).ToUniversalTime())}
            }
            if (![string]::IsNullOrEmpty($endDate) -and ![string]::IsNullOrEmpty($startDate)) 
            {
                 write-host 'searching google last range'
                #$startDate = [DateTime]::Parse($dateRange.split('-')[0])
                #$endDate = [DateTime]::Parse($dateRange.split('-')[1])
                $entries = $entries | where {((convertWebKitToUTC($_.last_visit_time)) -GE $startDate) -and ((convertWebKitToUTC($_.last_visit_time)) -le $endDate)}
            }
            Add-Type -AssemblyName System.Web
            $redirects = $entries | where { $_.url -like '*google.com/url?sa*'}
            $searches = $entries | where { $_.url -like '*google.com/search*'}
            foreach ($row in $searches) 
            {
                $item = New-Object PSObject -Property @{
                Date = checkExists(convertWebKitToUTC($row.last_visit_time).ToString()) 
                Title = checkExists($row.title)
                Url = checkExists($row.url) 
                }
                $history += $item
            }
            Write-Host "`n`n===========The Following URLs are referred from Google==========="
            foreach ($row in $redirects)
            {
                $url = [System.Web.HttpUtility]::UrlDecode(($row.url -split '&url=')[1])
                $item = New-Object PSObject -Property @{
                Date = checkExists(convertWebKitToUTC($row.last_visit_time).ToString()) 
                Title = checkExists($row.title)
                Url = checkExists($url) 
                }
                $history += $item
            }
        }
        else 
        {
            $entries = $data.Tables[0]
            if ($lastHours -gt 0)
            {
                $entries = $entries | where {(convertWebKitToUTC($_.last_visit_time)) -GE ((Get-Date).AddHours(-$lastHours).ToUniversalTime())}
            }
            if (![string]::IsNullOrEmpty($endDate) -and ![string]::IsNullOrEmpty($startDate)) 
            {
                #$startDate = [DateTime]::Parse($dateRange.split('-')[0])
                #$endDate = [DateTime]::Parse($dateRange.split('-')[1])
                $entries = $entries | where {((convertWebKitToUTC($_.last_visit_time)) -GE $startDate) -and ((convertWebKitToUTC($_.last_visit_time)) -le $endDate)}
            }
            foreach ($row in $entries)
            {
                $item = New-Object PSObject -Property @{
                Date = checkExists(convertWebKitToUTC($row.last_visit_time).ToString()) 
                Title = checkExists($row.title)
                Url = checkExists($row.url) 
                }
                $history += $item
            }
        }
        $con.Close()
        $browser = $browser.split('\')[6]
        $path = "C:\Temp\browserhistory\"
        If(!(test-path $path))
        {
              New-Item -ItemType Directory -Force -Path $path
        }
        $history | Export-Csv -NoTypeInformation -Path "$path$($browser)_history.csv"
        $sql = $null
        $con = $null
        $adapter = $null
	    [System.GC]::Collect()
	    Remove-Item -Path $cpFile

    }
}

#legacy IE history
function getIEHistory()
{
    $history = @()
    $shell = New-Object -ComObject Shell.Application            
    $ie = $shell.NameSpace(“C:\Users\$username\AppData\Local\Microsoft\Windows\History”)            
    $folder = $ie.Self            
            
    $ie.Items() |             
    foreach {            
     if ($_.IsFolder) {            
       $siteFolder = $_.GetFolder            
       $siteFolder.Items() |             
       foreach {            
         $site = $_            
             
         if ($site.IsFolder) {            
            $pageFolder  = $site.GetFolder            
            $pageFolder.Items() |             
            foreach { 
               $visit = New-Object -TypeName PSObject -Property @{            
                   Site = $($site.Name)            
                   URL = $($pageFolder.GetDetailsOf($_,0))            
                   Date = $( $pageFolder.GetDetailsOf($_,2))
               }   
               if ($lastHours -gt 0)
               {
                if (([Datetime]$visit.Date).ToUniversalTime() -GE ((Get-Date).AddHours(-$lastHours).ToUniversalTime()))
                {
                    $history += $visit  
                }
               }
               elseif (![string]::IsNullOrEmpty($endDate) -and ![string]::IsNullOrEmpty($startDate))
               { 
                if ((([Datetime]$visit.Date).ToUniversalTime() -GE $startDate) -and (([Datetime]$visit.Date).ToUniversalTime() -le $endDate))
                {
                    $history += $visit  
                }   
               
               }
               else 
               {       
                    $history += $visit      
               }
            }            
         }            
       }            
     }            
}
    $history | Export-Csv -NoTypeInformation "C:\Temp\browserhistory\IE_history.csv"
}

try 
{
    if (![string]::IsNullOrEmpty($search) -and $googleSearches)
    {
        Throw "Please only choose either a custom search or google search at one time"
    }
    
    Switch ((Get-Culture).TextInfo.ToTitleCase($browser))
    {
        'All' 
        {
            loadDLL
            Write-Host "===========The Following URLs are from Chrome==========="
            getChromiumHistory($chrome)
            Write-Host "`n`n===========The Following URLs are from Edge==========="
            getChromiumHistory($edge)
            getIEHistory
        }
        'Chrome' 
        {
            loadDLL
            Write-Host "`n`n===========The Following URLs are from Chrome==========="
            getChromiumHistory($chrome)
        }
        'Edge' 
        {
            loadDLL
            Write-Host "`n`n===========The Following URLs are from Edge==========="
            getChromiumHistory($edge)
        }
        'IE' 
        {
            getIEHistory
        }

    }
    if (Test-Path "C:\Temp\browserhistory") 
    {
        if (Test-Path "c:\temp\browserhistory.zip")
        {
            Remove-Item "c:\temp\browserhistory.zip"
        }
        Compress-Archive "C:\Temp\browserhistory" -DestinationPath "C:\Temp\browserhistory.zip"
    }


}
catch
{
    New-Item -Path "c:\temp\browserhistory" -Name "getBrowserHistoryLog.txt" -ItemType "file"
    Add-Content -Path "c:\temp\browserhistory\getBrowserHistoryLog.txt" -Value "Error Message:  $($Error[0].Exception.Message)"
    Add-Content -Path "c:\temp\browserhistory\getBrowserHistoryLog.txt" -Value "Error in Line: $($Error[0].InvocationInfo.Line)"
    Add-Content -Path "c:\temp\browserhistory\getBrowserHistoryLog.txt" -Value "Error in Line Number: $($Error[0].InvocationInfo.ScriptLineNumber)"
    Add-Content -Path "c:\temp\browserhistory\getBrowserHistoryLog.txt" -Value "Error Item Name: $($Error[0].Exception.ItemName)"
}
