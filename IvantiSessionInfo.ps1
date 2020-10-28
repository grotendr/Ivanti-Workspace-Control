#Array With collected Info
$AllData= @()
$i = $null

#Server, Database, User, Password
$SRV= Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Wow6432Node\RES\Workspace Manager' -Name 'DBServer'
$DB = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Wow6432Node\RES\Workspace Manager' -Name 'DBName'
#$DB="WorkspaceLogging"
$uid = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Wow6432Node\RES\Workspace Manager' -Name 'DBUser'
$pwd = Read-Host "Enter password for $uid "



#Create Connection
$sqlConn = New-Object System.Data.SqlClient.SqlConnection
$sqlConn.ConnectionString = “Server=$SRV; User ID = $uid; Password = $pwd; Initial Catalog=$DB”
$sqlConn.Open()
$pwd = $null

#Create Command
$sqlcmd = New-Object System.Data.SqlClient.SqlCommand
$sqlcmd.Connection = $sqlConn

#Query to Run
#$query = "select Convert(NVarchar(max),Convert(Varbinary(max),imgInfo)) from tbllogs where lngClassID = 47"
$query = “SELECT lngConnectionState, strUser,strComputerName From tblLicenses"
$sqlcmd.CommandText = $query

#Data Adapter
$adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcmd

#Create Dataset and fill
$data = New-Object System.Data.DataSet
$adp.Fill($data) | Out-Null

#Show Data
#$data.Tables


#Close Connection
$sqlConn.Close()

ForEach ($tblitem in $data.Tables[0]){
    $ADUser=($tblitem.strUser).Replace('DEVENTER\','')
    
    
    #Count Disconnected Sessions
    If ($tblitem.lngConnectionState -eq 0){
        $i = $i + 1
        }

    #Get Info from Active Directory
    [string]$Scope = "Subtree" 
    $Filter = "((samaccountname=$ADUser))"
    $RootOU = "OU=Medewerkers,OU=Organisatie,DC=deventer,DC=intern"

    $Searcher = New-Object DirectoryServices.DirectorySearcher
    $Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
    $Searcher.Filter = $Filter
    $Searcher.SearchScope = $Scope # Either: "Base", "OneLevel" or "Subtree"
    $Searcher.FindOne()

    #Fill The Array with Data          
            $item = New-Object PSObject
            $item | Add-Member -MemberType NoteProperty -Name 'Active' -Value $tblitem.lngConnectionState
            $item | Add-Member -MemberType NoteProperty -Name 'LogonName' -Value $ADUser
            $item | Add-Member -MemberType NoteProperty -Name 'FullName' -Value $((($Searcher.FindOne()).Properties).displayname)
            $item | Add-Member -MemberType NoteProperty -Name 'Mailadres' -Value $((($Searcher.FindOne()).Properties).mail)
            $item | Add-Member -MemberType NoteProperty -Name 'Afdeling' -Value $((($Searcher.FindOne()).Properties).department)
            $item | Add-Member -MemberType NoteProperty -Name 'LogonCount' -Value $((($Searcher.FindOne()).Properties).logoncount)

            $AllData += $item
    
    }

    $AllData | Out-GridView -Title "$(Get-Date -Format "dddd dd/MM/yyyy HH:mm") --- Ivanti Sessions: Disconnected Total:$i  |  Connected/Active Total:$($AllData.Count - $i) " -OutputMode Multiple


