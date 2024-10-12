############################################################
#                                                          #
#            Projet : Création de VM sur l'ESXi            #
#            Auteur : Marijan Stajic                       #
#            Version : 1                                   #
#            Dernière modification : 07.06.2023            #
#            Date de création : 05.06.2023                 #
#                                                          #
############################################################


Clear-Host

# Vérification si PowerCLI est déjà installé ou non, si ce n'est pas le cas, installer

Write-Host "Vérification du module VMware.PowerCLI ..." -ForegroundColor Yellow

$module_powercli = Get-Module -ListAvailable -Name "*VMware.PowerCLI*"

if ($module_powercli) {

    Write-Host "PowerCLI est déjà installé sur le serveur" -ForegroundColor Green

}

else {
 
    Install-Module VMware.PowerCLI

    Write-Host "Installation de PowerCLI ...`n" -ForegroundColor Red
}

Start-Sleep 2

# Identification pour la connexion à l'ESXi

Clear-Host

Write-Host "Veuillez saisir les informations de connexion pour le serveur esxi-server-01`n" -ForegroundColor Yellow

$server = "172.16.0.6"
$username = Read-Host "Entrez le nom de l'utilisateur"
$password = Read-Host "Entrez le mot de passe"

# Connexion à l'ESXi

Write-Host "`nConnexion au serveur esxi-server-01 ..." -ForegroundColor Yellow

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null
$connect = Connect-VIServer -Server $server -User $username -Password $password

if ($connect) {

    Write-Host "Connexion au serveur réussi`n" -ForegroundColor Green

    Start-Sleep 2

    Clear-Host

    # Création d'une machine virtuelle

    # Initialisation des variables

    $check_vm_name = $false
    $check_vm_os = $false
    $check_vm_cpu = $false
    $check_vm_memory_go = $false
    $check_vm_disk_go = $false
    $check_vm_datastore = $false
    $check_vm_creation = $false
    $check_add_vm_veeam = $false

    # Définition du Nom

    while (-not $check_vm_name) {  
                
        $vm_name = Read-Host "Veuillez saisir le nom de votre machine virtuelle"

        $vm_get_all = Get-VM | Select-Object -ExpandProperty Name

        if ($vm_get_all -contains $vm_name) {

            Write-Host "Le nom $vm_name est déjà utilisé par une autre machine virtuelle" -ForegroundColor Red
        }

        else {

            if ($vm_name.Length -ge 2) {
                    
                $check_vm_name = $true
            }

            else {

                Write-Host "Le nom de la machine virtuelle doit comporter au moins 2 caractères" -ForegroundColor Red
            }
        }
    }

    # Sélection de l'OS
    
    while (-not $check_vm_os) {
        
        Write-Host "`n1. Linux - Debian`n2. Linux - Ubuntu`n" -ForegroundColor Yellow

        $vm_os = Read-Host "Veuillez indiquer quel système d'exploitation souhaitez-vous utiliser (1 à 2)"

        if ($vm_os -lt 1 -or $vm_os -gt 2) {

            Write-Host "La valeur doit être comprise entre 1 (minimum) et 1 (maximum)" -ForegroundColor Red

        }

        else {

            if ($vm_os -eq 1) {
                
                $vm_os = "debian10_64Guest"
                $vm_path = "[datastore1] ISO/debian-11.1.0-amd64-netinst.iso"
            }

            elseif ($vm_os -eq 2) {
                
                $vm_os = "ubuntu64Guest"
                $vm_path = "[datastore1] ISO/ubuntu-22.04.1-desktop-amd64.iso"
            }

            $check_vm_os = $true
        }
    }

    # Sélection du CPU
    
    while (-not $check_vm_cpu) {

        [int] $vm_cpu = Read-Host "Veuillez indiquer la quantité de CPU pour votre machine virtuelle (1 à 12)"

        if ($vm_cpu -lt 1 -or $vm_cpu -gt 12) {

            Write-Host "La valeur doit être comprise entre 1 CPU (minimum) et 12 CPU (maximum)" -ForegroundColor Red
        }

        else {

            $check_vm_cpu = $true
        }
    }

    # Sélection de la mémoire

    while (-not $check_vm_memory_go) {

        [int] $vm_memory_go = Read-Host "Veuillez indiquer la quantité de mémoire pour votre machine virtuelle (en Go, 1 à 8)"

        if ($vm_memory_go -lt 1 -or $vm_memory_go -gt 8) {

            Write-Host "La valeur doit être comprise entre 1 Go (minimum) et 8 Go (maximum)" -ForegroundColor Red
        }

        else {

            $check_vm_memory_go = $true
        }
    }

    # Sélection du disque

    while (-not $check_vm_disk_go) {

        [int] $vm_disk_go = Read-Host "Veuillez indiquer la quantité de stockage pour votre machine virtuelle (en Go, 20 à 50)"

        if ($vm_disk_go -lt 20 -or $vm_disk_go -gt 50) {

            Write-Host "La valeur doit être comprise entre 20 Go (minimum) et 50 Go (maximum)" -ForegroundColor Red
        }

        else {

            $check_vm_disk_go = $true
        }
    }

    # Sélection du datastore

    Get-Datastore | Select-Object Name, FreeSpaceGB | Out-Host

    while (-not $check_vm_datastore) {

        $vm_datastore = Read-Host "Veuillez saisir le nom du datastore que vous souhaitez utiliser pour stocker votre machine virtuelle (datastore1 ou datastore2)"

        if ($vm_datastore -eq "datastore1" -or $vm_datastore -eq "datastore2") {
            
            $selected_datastore = Get-Datastore -Name $vm_datastore

            if ($selected_datastore.FreeSpaceGB -ge ($vm_disk_go + 5)) {

                $check_vm_datastore = $true
            }

            else {
                
                Write-Host "La $vm_datastore sélectionnée ne dispose pas d'un espace disque suffisant pour stocker votre machine virtuelle qui requiert $vm_disk_go Go" -ForegroundColor Red
                $vm_datastore = ""
            }

        }

        else {
           
            Write-Host "Veuillez saisir les noms des datastores disponibles, datastore1 et datastore2, en fonction de leur espace de stockage disponible" -ForegroundColor Red
        }
    }

    Clear-Host

    # Résumé de la configuration de la machine virtuelle

    Write-Host "Voici un résumé de la configuration de votre machine virtuelle`n"

    Write-Host "Nom: $vm_name`nSystème d'exploitation: $vm_os`nCPU: $vm_cpu`nMémoire: $vm_memory_go Go`nEspace disque: $vm_disk_go Go`nDatastore: $vm_datastore`nRéseau: VM Network`n" -ForegroundColor Yellow

    # Validation de la création de la machine virtuelle

    while (-not $check_vm_creation) {

        $creation =  Read-Host "Veuillez appuyer sur [Y] pour confirmer la création ou sur [N] pour annuler"

        # Création de la machine virtuelle

        if ($creation -eq "Y") {

            New-VM -Name $vm_name -GuestID $vm_os -NumCPU $vm_cpu -MemoryGB $vm_memory_go -DiskGB $vm_disk_go -Datastore $vm_datastore -NetworkName "VM Network" | Out-Null
            
            New-CDDrive -VM $vm_name -IsoPath $vm_path -StartConnected | Out-Null

            Write-Host "`nLa machine virtuelle $vm_name à été créer avec succès`n" -ForegroundColor Green

            # Ajouter à Veeam la machine virtuelle

            while (-not $check_add_vm_veeam) {
                
                Clear-Host

                $add_vm_veeam = Read-Host "Souhaitez-vous ajouter $vm_name au stockage de Veeam ? Appuyez sur [Y] pour confirmer l'ajout ou sur [N] pour annuler"

                if ($add_vm_veeam -eq "Y") {

                    # Connexion au serveur Veeam

                    Write-Host "`nConnexion au serveur Veeam de win-server-02 ..." -ForegroundColor Yellow

                    Connect-VBRServer -User "WIN-SERV-02\Administrator" -Password "vTx5upp0rT" -Server 127.0.0.1 -Port 9392

                    Write-Host "Connexion au serveur win-server-02 réussi`n" -ForegroundColor Green

                    Start-Sleep 2

                    # Ajout de la VM au Job Backup ESXi

                    $vm_name_veeam = Find-VBRViEntity -Name $vm_name

                    $job_veeam = Get-VBRJob -Name "Backup VMs - esxi-server-01"

                    Add-VBRViJobObject -Job $job_veeam -Entities $vm_name_veeam | Out-Null

                    Write-Host "L'ajout de la machine virtuelle $vm_name à Veeam été faites avec succès" -ForegroundColor Green

                    Start-Sleep 2

                    $check_add_vm_veeam = $true

                    $check_vm_creation = $true

                    # Déconnexion du serveur Veeam

                    Disconnect-VBRServer
                }

                # Ne pas ajouter la machine virtuelle à Veeam

                elseif ($add_vm_veeam -eq "N") {

                    Write-Host "`nL'ajout de la machine virtuelle $vm_name à Veeam été annulée" -ForegroundColor Red

                    Start-Sleep 2

                    $check_add_vm_veeam = $true

                    $check_vm_creation = $true
                }

                else {

                    Write-Host "Veuillez saisir [Y] pour confirmer l'ajout ou [N] pour annuler" -ForegroundColor Red
                }

            }

            $check_vm_creation = $true

        # Annuler la création de la machine

        }

        elseif ($creation -eq "N") {

            Write-Host "`nLa création de la machine virtuelle à été annulée" -ForegroundColor Red

            Start-Sleep 2

            $check_vm_creation = $true
        }

        else {
                
            Write-Host "Veuillez saisir [Y] pour confirmer la création ou [N] pour annuler" -ForegroundColor Red
        }
    }
}

else {

    # Echec de la connexion

    Write-Host "Connexion au serveur échoué" -ForegroundColor Red

    Start-Sleep 2
}