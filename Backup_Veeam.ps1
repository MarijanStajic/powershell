############################################################
#                                                          #
#            Projet : Backup Veeam                         #
#            Auteur : Marijan Stajic                       #
#            Version : 1                                   #
#            Dernière modification : 09.06.2023            #
#            Date de création : 09.06.2023                 #
#                                                          #
############################################################


Clear-Host

# Connexion au serveur Veeam

Write-Host "`nConnexion au serveur Veeam sur win-serv-02 ..." -ForegroundColor Yellow

# Connect-VBRServer -User "WIN-SERV-02\Administrator" -Password "vTx5upp0rT" -Server 127.0.0.1 -Port 9392

Write-Host "Connexion au serveur réussi `n" -ForegroundColor Green

Start-Sleep 2

# Initialisation des variables de vérification

Clear-Host

$vms_exist = Find-VBRViEntity | Select-Object -ExpandProperty Name
$job_objects = Get-VBRJobObject -Job "Backup VMs - esxi-server-01" | Where-Object { $_.Type -eq "Include" } | Select-Object -ExpandProperty Name
$check_all_exist = $false

# Vérification de la disponibilité des machines virtuelles

while (-not $check_all_exist) {

    Write-Host "Vérification des machines virtuelles ...`n" -ForegroundColor Yellow

    foreach ($job_object in $job_objects) {

        $vm_found = $false
    
        foreach ($vm in $vms_exist) {

            if ($vm -eq $job_object) {

                $vm_found = $true
            }
        }
    
        # Machine virtuelle toujours existantes

        if ($vm_found) {
        
            Write-Host "La machine virtuelle $job_object existe toujours" -ForegroundColor Green
        }

        # Machine virtuelle inexistante 

        elseif (-not $vm_found) {

            Write-Host "La machine virtuelle $job_object a été supprimée de l'ESXi. Sa stratégie de sauvegarde et ses données ont également été supprimées pour libérer de l'espace disque et éviter les erreurs." -ForegroundColor Red

            # Suppression de la machine virtuelle de la liste à backup

            Get-VBRJobObject -Job "Backup VMs - esxi-server-01" -Name $job_object | Remove-VBRJobObject

            # Suppression du disque de la machine virtuelle

            Get-VBRRestorePoint -Name $job_object | Remove-VBRRestorePoint -Confirm:$false
        }
    }

    # Vérification complète

    Start-Sleep 2

    $check_all_exist = $true 
}

# Lancement du backup simple

Clear-Host

Write-Host "Lancement du backup des machines virtuelles ...`n" -ForegroundColor Yellow

Start-VBRJob -Job "Backup VMs - esxi-server-01"

Write-Host "`nBackup terminée avec succès" -ForegroundColor Green

Start-Sleep 3