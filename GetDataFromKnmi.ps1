$outputPath = "E:\Data\";
$urlDay  = "http://projects.knmi.nl/klimatologie/daggegevens/getdata_dag.cgi";
$urlHour = "http://projects.knmi.nl/klimatologie/uurgegevens/getdata_uur.cgi";
$startYear = 2005;

###################################################################################

# Get the KNMI data using a webrequest post
function DoWebRequest($url, $params, $outputFile)
{
    [System.Net.HttpWebRequest] $request = [System.Net.WebRequest]::Create($url);
    $request.ContentType = "application/x-www-form-urlencoded";
    $request.Method="POST";
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($params);
    $request.ContentLength = $buffer.Length;

    $requestStream = $request.GetRequestStream();
    $requestStream.Write($buffer, 0, $buffer.Length);
    $requestStream.Flush();
    $requestStream.Close();

    [System.Net.HttpWebResponse] $response = $request.GetResponse();
    $responseStream = $response.GetResponseStream();
    $reader = New-Object System.IO.StreamReader($responseStream);
    $result = $reader.ReadToEnd();
    $utf8 = New-Object System.Text.UTF8Encoding($false);
    [System.IO.File]::WriteAllLines($outputFile, $result, $utf8);

    $reader.Close();
    $reader.Dispose();
    
    write-host "Finished export file: $outputFile";
}

$curYear = (Get-Date).Year;
$prevYear = $null;
if ((Get-Date).Month -eq 1 -and (Get-Date).Day -le 3) { $prevYear = $curYear-1; } # On first 3 days of year also get data from last year

for ($year=$startYear;$year -le $curYear;$year++)
{
    $gotDay = $gotHour = $false;
    $fileDay = [System.IO.Path]::Combine($outputPath, "knmi-dag_$year.txt");
    $fileHour = [System.IO.Path]::Combine($outputPath, "knmi-uur_$year.txt");

    # The following parameter lines could be edited to get different data
    # Check the following URL's for more information and an interactive selection option:
    # http://projects.knmi.nl/klimatologie/daggegevens/selectie.cgi
    # http://projects.knmi.nl/klimatologie/uurgegevens/selectie.cgi
    $paramsDay = "stns=ALL&vars=ALL&start=$year 0101&end=$year 3112" -Replace " ", "";
    $paramsHour = "stns=ALL&vars=TEMP:SUNR:PRCP:VICL&start=$year 010101&end= $year 311224" -Replace " ","";
    
    # If files do not exist for the year, create the files
    if (-not (Test-Path $fileDay))
    {
        DoWebRequest $urlDay $paramsDay $fileDay;
        $gotDay = $true;
    }
    if (-not (Test-Path $fileHour))
    {
        DoWebRequest $urlHour $paramsHour $fileHour;
        $gotHour = $true;
    }
    
    # Always refresh the current year file and optionaly the file from last year
    if ($year -eq $curYear -or $year -eq $prevYear)
    {
        if (-not ($gotDay)) { DoWebRequest $urlDay $paramsDay $fileDay; }
        if (-not ($gotHour)) { DoWebRequest $urlHour $paramsHour $fileHour; }
    }
}


