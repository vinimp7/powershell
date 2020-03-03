<# 
Author:     Vinícius Santos - Sempre IT
Date:       20/02/2020
Script:     Set-MSTeamsGroups.ps1 
Version:    0.0.2
#>


#Log in MSTeams
$stname = 'StorageAccountName'
$stkey  = 'StorageAccountKey'
$container = 'container'
$blob = 'users.csv'
$userName = "Main Account that will create everything"
$securePassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($userName, $securePassword)
Connect-MicrosoftTeams -Credential $cred


#download the CSV file from the blob
$path = $env:TEMP + "\users.csv"
$path 
$context = New-AzureStorageContext -StorageAccountName $stname -StorageAccountKey $stkey
Get-AzureStorageBlobContent -Container $container -Blob $blob -Destination $path -Context $context -Force


#Inicio do Script
$groupID = "The main Team ID from AAD"
$usuariosBIG    = Import-csv -Path $path -Delimiter "," -Header nome , email , cod_projeto , role_usuario -Encoding UTF8 
$usuariosTeams  = Get-TeamUser -GroupId $groupID
$canaisTeams    = Get-TeamChannel -GroupId $groupID
$usuariosBIG


#Check if the channels already exist in MSTeams
$UsuariosDiferenca = Compare-Object -ReferenceObject $usuariosTeams.user -DifferenceObject $usuariosBIG.email -IncludeEqual
$CanaisDiferenca = Compare-Object -ReferenceObject $canaisTeams.DisplayName -DifferenceObject $usuariosBIG.cod_projeto -IncludeEqual

$Canais = @()
#Verificar se os canais já existem
foreach($CD in $CanaisDiferenca){
    if(($CD.SideIndicator -ne "=>") -or ($CD.InputObject -eq "General")){
        Write-Output "."
    }else{
        $Canais += $CD.InputObject 
    }
    
}

#Create the new channels
$CanaisTeste = @()
$CanaisN = $Canais | select -Unique
$CanaisNovos = Compare-Object -ReferenceObject $canaisTeams.DisplayName -DifferenceObject $CanaisN -IncludeEqual
foreach($CN in $CanaisNovos){
    if($CN.SideIndicator -ne "=>"){
        Write-Output "."
    }else{
        New-TeamChannel -GroupId $groupID -DisplayName $CN.InputObject -MembershipType Private
        Write-Output "Canal" $CN.InputObject "adicionado"
        $CanaisTeste += $CN.InputObject
    }

}

Compare-Object -ReferenceObject $canaisTeams.DisplayName -DifferenceObject $canaisTeste -IncludeEqual

#Exclude ADM from the list
foreach($u in $usuariosTeams){
    if($u.Role -eq "owner"){
        $owner = $u.User
    }
}

$UsuariosN = @()
#Check if the users already exist in MSTeams
foreach($usuarios in $UsuariosDiferenca){
    if(($usuarios.SideIndicator -ne "=>") -or ($usuarios.InputObject -eq $owner)){
        Write-Output "."
    }else{
        $UsuariosN += $usuarios.InputObject
    }
}

#Adding the new users
$UsuariosNov = $UsuariosN | select -Unique
$UsuariosNovos = Compare-Object -ReferenceObject $usuariosTeams.user -DifferenceObject $UsuariosNov -IncludeEqual

foreach($UN in $UsuariosNovos){
    if($UN.SideIndicator -ne "=>"){
        Write-Output "."
    }else{
        Add-TeamUser -GroupId $groupID -User $UN.InputObject
        Write-Output $UN.InputObject "adicionado"
    }
}

#Allocate the users on their specific channels
$canaisTeamsAtual = Get-TeamChannel -GroupId $groupID

foreach($Can in $canaisTeamsAtual){
    $usuariosTeamsAtual = Get-TeamChannelUser -GroupId $groupID -DisplayName $Can.DisplayName
	foreach($Usuar in $usuariosBIG){
        if($usuariosTeamsAtual.User -contains $Usuar.email){
            Write-Output "." 
        }elseif(($Usuar.cod_projeto -eq $Can.DisplayName) -and ($Usuar.email -ne $owner)){
            Add-TeamChannelUser -GroupId $groupID -DisplayName $Can.DisplayName -User $Usuar.email
			Write-Output $Usuar.email "foi adicionado ao projeto:" $Can.DisplayName
            if($Usuar.role_usuario -eq "Propietário"){
                Add-TeamChannelUser -GroupId $groupID -DisplayName $Can.DisplayName -User $Usuar.email -Role Owner
                Write-Output $Usuar.email "agora é propietário do projeto:" $Can.DisplayName
            }
        }
	} 
}