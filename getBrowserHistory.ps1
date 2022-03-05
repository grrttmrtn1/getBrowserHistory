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
    [Parameter(HelpMessage="Provide a date range for time you'd like to search. ex. 2/2/2022 3:15-2/2/2022 13:15. Please use UTC.")]
    [datetime]$dateRange,
    [Parameter(HelpMessage="Provide integer parameter for the number of hours from current time you'd like to search in history")]
    [int]$lastHours
)


$chrome = "C:\Users\$username\AppData\Local\Google\Chrome\User Data\Default\History"
$edge ="C:\Users\$username\AppData\Local\Microsoft\Edge\User Data\Default\History"


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

#utilized for both Chrome/Edge just need a different db
function getChromiumHistory($browser)
{
    $cpFile = "$browser-cp"
    Copy-Item -Path $browser -Destination $cpFile
    if (Test-Path $cpFile)
    {

        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=$cpFile"
        $con.Open()
        $sql = $con.CreateCommand()
        if ([string]::IsNullOrEmpty($search))
        {
            $sql.CommandText = "select * from urls;"
        }
        else 
        {
            if ($googleSearches) 
            {
                $sql.CommandText = "select * from urls where url like '%google%';"
            }
            else 
            {
                $sql.CommandText = "select * from urls where url like '%$search%';"
            }
        }
        $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
        $data = New-Object System.Data.DataSet
        [void]$adapter.Fill($data)
        if ($googleSearches)
        {
            Add-Type -AssemblyName System.Web
            $redirects = $data.Tables[0] | where { $_.url -like '*google.com/url?sa*'}
            $searches = $data.Tables[0] | where { $_.url -like '*google.com/search*'}
            foreach ($item in $searches) 
            {
                $dateTime = (([System.DateTimeOffset]::FromUnixTimeSeconds(($item.last_visit_time/1000000 - 11644473600))).DateTime).ToString()
                Write-Host "$($item.title) - $($item.url) - $dateTime"
            }
            Write-Host "`n`n===========The Following URLs are referred from Google==========="
            foreach ($item in $redirects)
            {
                $dateTime = (([System.DateTimeOffset]::FromUnixTimeSeconds(($item.last_visit_time/1000000 - 11644473600))).DateTime).ToString()
                $url = [System.Web.HttpUtility]::UrlDecode(($item.url -split '&url=')[1])
                Write-Host "$($item.title) - $url - $dateTime"
            }
        }
        else 
        {
            foreach ($row in $data.Tables[0])
            {
                $dateTime = (([System.DateTimeOffset]::FromUnixTimeSeconds(($row.last_visit_time/1000000 - 11644473600))).DateTime).ToString()
                write-host "$($row.title) - $($row.url)  - $dateTime"
            }
        }
        $con.Close()
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
    $shell = New-Object -ComObject Shell.Application            
    $hist = $shell.NameSpace(“C:\Users\$username\AppData\Local\Microsoft\Windows\History”)            
    $folder = $hist.Self            
            
    $hist.Items() |             
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
               write-host $visit            
            }            
         }            
       }            
     }            
}
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


}
catch
{
    write-host "Error Message:  $($Error[0].Exception.Message)"
    write-host "Error in Line: $($Error[0].InvocationInfo.Line)"
    write-host "Error in Line Number: $($Error[0].InvocationInfo.ScriptLineNumber)"
    write-host "Error Item Name: $($Error[0].Exception.ItemName)"
}
