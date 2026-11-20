<#
dax_query.ps1 - 열려 있는 Power BI Desktop 모델에 DAX(또는 DMV) 질의를 보내 결과를 표로 출력한다.

    powershell -File tools\dax_query.ps1 -QueryFile q.dax          # DAX 를 파일에 적어 넘긴다
    powershell -File tools\dax_query.ps1 -Query "SELECT [Name] FROM `$SYSTEM.TMSCHEMA_MEASURES"

교안의 측정값 결과가 맞는지 확인할 때만 쓰는 도구다 (학생은 실행하지 않아도 된다).
  - Power BI Desktop 은 파일을 열면 내부에 Analysis Services 엔진(msmdsrv.exe)을 띄우고,
    그 포트 번호를 AnalysisServicesWorkspaces\...\Data\msmdsrv.port.txt 에 적어 둔다.
  - ADOMD 클라이언트(NuGet: Microsoft.AnalysisServices.AdomdClient.retail.amd64)로 그 포트에 접속한다.
    -AdomdDll 로 DLL 위치를 넘긴다.
#>
param(
    [string]$Query = "",
    [string]$QueryFile = "",
    [string]$AdomdDll = "$env:TEMP\adomd\Microsoft.AnalysisServices.AdomdClient.dll",
    [string]$OutCsv = ""
)
Add-Type -Path $AdomdDll
$ws = "$env:USERPROFILE\Microsoft\Power BI Desktop Store App\AnalysisServicesWorkspaces"
if (-not (Test-Path $ws)) { $ws = "$env:LOCALAPPDATA\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces" }
# 가장 최근에 열린 작업 공간의 포트
$portFile = Get-ChildItem $ws -Recurse -Filter msmdsrv.port.txt | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$port = ([IO.File]::ReadAllText($portFile.FullName, [Text.Encoding]::Unicode)).Trim()
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source=localhost:$port")
$conn.Open()
# Desktop 의 작업 공간에는 모델이 하나뿐이라 기본 데이터베이스가 곧 그 모델이다
if ($QueryFile) { $Query = [IO.File]::ReadAllText($QueryFile, [Text.Encoding]::UTF8) }  # 따옴표가 섞인 긴 DAX 는 파일로
$cmd = $conn.CreateCommand()
$cmd.CommandText = $Query
$table = New-Object System.Data.DataTable
$adapter = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter($cmd)
[void]$adapter.Fill($table)
$conn.Close()
if ($OutCsv) { $table | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8 }
$table | Format-Table -AutoSize | Out-String -Width 400
