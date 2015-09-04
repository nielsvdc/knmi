###################################################################################
# NAME:   GetDataFromKnmi.ps1
# AUTHOR: Niels van de Coevering
# DATE:   05 September 2015
#
# COMMENTS: This script will use a webrequest post action to retrieve Dutch weather
# data from the Royal National Weather Institute KNMI website and use the getdata
# service for daily and hourly statistics per weather station.
#
# This script is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
###################################################################################

$outputPath = "E:\Data\";
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

$urlDay  = "http://projects.knmi.nl/klimatologie/daggegevens/getdata_dag.cgi";
$urlHour = "http://projects.knmi.nl/klimatologie/uurgegevens/getdata_uur.cgi";
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
    $paramsHour = "stns=ALL&vars=ALL&start=$year 010101&end= $year 311224" -Replace " ","";
    
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


