#David L. 05/11/2020, v0.0.1
#Script Pulls list of computers from specified OU, Tests if they are online with a ping command, searches for TPM values of online systems, exports data into 3 CSVs (

#Set path var to where ever the script is currently running
$path = Split-Path -parent $MyInvocation.MyCommand.Definition

#CSV Data Exports
$OnlineVisibleTPM = $path + "\OnlineVisibleTPM.csv"
$OnlineNoTPM = $path + "\OnlineNoTPM.csv"
$Offline = $path + "\Offline.csv"

#OU to pull systems from
$OUpath = 'OU=Win10,OU=Computers,OU=AD,DC=pop,DC=portptld,DC=com'

#Array of computers from above OU
$Win10Computers = Get-ADComputer -Filter * -SearchBase $OUpath | Sort-Object


foreach($computer in $Win10Computers){
    if(Test-Connection $computer.Name -Quiet -Count 2){
        #$computer | Select-Object -Property Name | Export-CSV $OnlineCSVExport -Force -NoTypeInformation -Append 
        $tpm = Get-WmiObject -class Win32_Tpm -namespace root\CIMV2\Security\MicrosoftTpm -ComputerName $computer.Name -Authentication PacketPrivacy    
        if($tpm -eq $null){
            $computer | Select-Object Name | Export-CSV $OnlineNoTPM -Append
        }
        else{
            $tpm | Select-Object PSComputerName,IsActivated_InitialValue,IsEnabled_InitialValue,IsOwned_InitialValue,ManufacturerVersionFull20,SpecVersion | Export-CSV $OnlineVisibleTPM -Append
        }
    }
    else{ 
        $computer | Select-Object -Property Name | Export-CSV $Offline -Force -NoTypeInformation -Append
    }
}