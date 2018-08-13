## - Name GetsendportURI.ps1
## - Eksample: ./getsendporturi.ps1 -port FILE 
Param(
[string]$SQLSERVER = "$SQLSERVERremote",
[string]$port = "*"
)
## - Get the name of the SQL server with the management database.
$SQLSERVERremote = Get-WmiObject -class MSBTS_GroupSetting -namespace root\MicrosoftBizTalkServer | Format-table -Property MgmtDbServerName -hidetableheaders| Out-String

## - Load the following Assembly:
[System.Reflection.Assembly]::loadwithPartialName("Microsoft.BizTalk.ExplorerOM");
 
## - Create a new empty variable to stored the .NET information:
$BTSexp = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer;
 
## - Now, this line will connect to your Biztalk Local instance:
$BTSexp.ConnectionString = `
  "Server=$SQLSERVERremote;Initial Catalog=BizTalkMgmtDb;Integrated Security=SSPI;";
 
## - Initialize variables:
$y = 0; [Array] $MyNewObject = $null;
 
## - Initialize variables:
$y = 0; [Array] $MyNewObject = $null;
 
#Add each result into our New object:
$MyNewObject = foreach($item in ($BTSexp.SendPorts))
{
 #Building your PowerShell PSObject item:
    $newPSitem = New-Object PSObject -Property @{
  seq = $y.ToString("000");
  Application = $item.Application.name;
  PortType = $item.PrimaryTransport.TransportType.Name;
  URI = $item.primarytransport.Address;
  Port = $item.Name;
 };
 #To display PSObject item values while processing:
 $newPSitem;
 $y++;
};
$MyNewObject | Where-Object {$_.PortType -like $Port}| Select seq,Application,PortType,URI | ft -auto;