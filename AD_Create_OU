#####################################################################
#                                                                   #
#      Auteur : Marijan Stajic                                      #
#      Date de création : 11.01.2023                                #
#                                                                   #
#####################################################################
#                                                                   #
#      Ce script permet de créer des OUs à partir d'un fichier      #
#      CSV ou manuelllement.                                        #
#                                                                   #
#####################################################################

# Initilisation des variables

$OU_Manual_or_CSV = ""

$OU_Manual_Create_New = ""

$OU_After_CSV = ""

$OU_Protected = ""

$OU_Manual_Create_New_Recreate = ""

# Début de la boucle pour la création

while ($OU_Manual_or_CSV -ne 1 -and $OU_Manual_or_CSV -ne 2) {
    
    # Sélection entre l'importation du CSV ou la création à la main

    $OU_Manual_or_CSV = Read-Host "Voulez-vous importer un fichier CSV [1] ou entrer les OUs à la main [2] ?"

    if ($OU_Manual_or_CSV -ne 1 -and $OU_Manual_or_CSV -ne 2) {
        Write-Host "Erreur ! Veuillez faire [1] pour importer un fichier CSV ou [2] entrer les OUs à la main."
    }

    # Si l'utilisateur choisi [1] donc création à partir d'un fichier CSV

    if ($OU_Manual_or_CSV -eq 1) {
        
        # Message pour vérifier si la colonne existe dans le fichier CSV

        Write-Host "Le message box pour confirmer si votre CSV contient bien certaines colonnes peut s'ouvrir en arrière-plan. Veuillez vérifier si vous ne le voyez pas."

        $OU_CSV_YesNo = [System.Windows.MessageBox]::Show("Il faut que votre fichier contienne une colonne 'Departement' si vous souhaitez pouvoir créer des OUs avec ce script. Est-ce votre fichier contient cette colonne ?", "Création de d'OU à partir d'un fichier CSV", [System.Windows.MessageBoxButton]::YesNo)

        # Ouverture de l'explorateur Windows et sélection du fichier CSV

        if ($OU_CSV_YesNo -eq [System.Windows.MessageBoxResult]::Yes) {
            
            $File_Explorer = New-Object System.Windows.Forms.OpenFileDialog 
            $File_Explorer.InitialDirectory = $InitialDirectory
            $File_Explorer.Filter = 'Fichier .CSV (*.csv)|*.csv'
            $File_Explorer.ShowDialog() | Out-Null
           
            $CSV_Path = $File_Explorer.FileName

            # Si on décide d'annuler l'importation

            if ($CSV_Path -eq "") {
                Write-Host "Vous avez décidé d'annuler l'importation du CSV."
            }
                        
            # Importation du fichier sélectionner
                        
            else {
                
                # Permet de vérifier si la colonne "Department" existe vraiment dans le CSV. Sinon, il retourne un message disant que non et passe à l'étape suivante.

                $CSV_File_Check = Import-Csv -Path $CSV_Path
                
                $CSV_File_Check[0].psobject.properties.name | Out-Null

                if ($CSV_File_Check -match "Department") {

                    $CSV_File = Import-Csv -Path $CSV_Path -Delimiter ";" -Encoding Default

                    # Début du foreach, donc toutes les $OU.Colonne seront basé sur le CSV_File choisi

                    foreach ($OU in $CSV_File) {
                        
                        $OU_Departement = $OU.Department

                        # Vérification si le contenu dans la colonne Department est vide

                        if ($OU_Departement -eq "") {

                            Write-Host  "Il est impossible de créer l'OU car le contenu dans la colonne Department est incomplète."
                        }

                        else {
                            
                            # Variable permettant de vérifier si l'OU n'est pas déjà existante dans l'AD

                            $OU_Exist = Get-ADOrganizationalUnit -Filter "Name -eq '$OU_Departement'"
                    
                            # Boucle pour l'activation ou non de la sécurité de l'objet dans l'AD

                            while ($OU_Protected -ne "O" -and $OU_Protected -ne "N") {
                                $OU_Protected = Read-Host "Voulez-vous activer la protection de suppresion de vos OUs ? [O/N]"
                
                                if ($OU_Protected -eq "O") {
                                    $OU_Protected_SOU = 1
                                }
                
                                if ($OU_Protected -eq "N") {
                                    $OU_Protected_SOU = 0
                                }

                                if ($OU_Protected -ne "O" -and $OU_Protected -ne "N") {
                                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez protéger vos OUs de la suppression ou [N] si vous ne voulez pas les protégers."
                                }
                            }    

                            # Permet d'indiquer que si la valeur de $OU_Last_Departement est différent à celle de $OU_Departement il peut lire la suite, sinon il recommence, cela permet d'éviter d'avoir plusieurs un message "Déjà existant" si le nom du departement apparait plusieurs fois dans la colonne

                            if ($OU_Last_Departement -ne $OU_Departement) {
                        
                                # Vérification si l'OU qui est tenté d'être créer est déjà existante

                                if ($OU_Exist) {
                                    Write-Host "$OU_Departement que vous tentez de créer est déjà existant dans Active Directory."
                                }

                                # Création de l'OU dans l'AD

                                else {
                                    New-ADOrganizationalUnit -Name $OU_Departement -Path "DC=GOTHAM,DC=CH" -ProtectedFromAccidentalDeletion $OU_Protected_SOU
                                    Write-Host "$OU_Departement à été créer avec succès."
                                }

                                # Permet d'indiquer que l'OU qui vient d'être créer soit égal à la variable OU_Last_Departement afin d'éviter d'avoir plusieurs un message "Déjà existant" si le nom du departement apparait plusieurs fois dans la colonne

                                $OU_Last_Departement = $OU_Departement
                            }
                        }
                    }
                }

                else {
                    Write-Host "Votre fichier CSV ne contient pas la colonne Department."
                }
            }

            # Boucle pour créer une OU à la main si on le souhaite après l'importation du CSV

            while ($OU_After_CSV -ne "N" -and $OU_After_CSV -ne "O") {
                $OU_After_CSV = Read-Host "Voulez-vous créer d'autres OUs à la main cette fois-ci ? [O/N]"

                if ($OU_After_CSV -eq "O") {
                    $OU_Manual_or_CSV = 2
                }
            
                if ($OU_After_CSV -ne "O" -and $OU_After_CSV -ne "N") {
                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez créer une autre OU à la main ou [N] pour stopper la création."
                }
            }
        }

        # Boucle pour créer une OU à la main si on le souhaite après avoir mis "Non" au message si la colonne Departement existe bien

        else {
            while ($OU_After_CSV -ne "N" -and $OU_After_CSV -ne "O") {
                $OU_After_CSV = Read-Host "Voulez-vous tout de même créer des OUs à la main ? [O/N]"

                if ($OU_After_CSV -eq "O") {
                    $OU_Manual_or_CSV = 2
                }
            
                if ($OU_After_CSV -ne "O" -and $OU_After_CSV -ne "N") {
                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez créer une autre OU à la main ou [N] pour stopper la création."
                }
            }
        }
    }
    
    # Si l'utilisateur choisi [2] donc création manuel
        
    if ($OU_Manual_or_CSV -eq 2) {

        # Boucle pour pouvoir créer des utilisateurs à la main jusqu'à avoir fais "N" pour sortir du script

        while ($OU_Manual_Create_New -ne "N") {
            
            # Initilisation des variables

            $OU_Protected = ""
            $OU_Manual_Create_New = ""
            $OU_Manual_Create_New_Recreate = 0
            $Check_OU_Manual_Create = 0

            # Boucle pour rentrer le nom de l'OU et vérifier si il fait bien au minimum un caractère

            while ($Check_OU_Manual_Create -ne 1) {
                
                $OU_Manual_Create = Read-Host "Quel nom souhaitez-vous mettre à votre OU ?"

                if ($OU_Manual_Create.Length -ge 1) {
                    
                    $Check_OU_Manual_Create++
                }

                else {

                    Write-Host "Le nom de l'OU doit faire au minimum 1 caractères."
                }
            }

            # Boucle pour l'activation ou non de la sécurité de l'objet dans l'AD

            while ($OU_Protected -ne "O" -and $OU_Protected -ne "N") {
                $OU_Protected = Read-Host "Voulez-vous activer la protection de suppresion de vos OUs ? [O/N]"
                
                if ($OU_Protected -eq "O") {
                    $OU_Protected_SOU = 1
                }
                
                if ($OU_Protected -eq "N") {
                    $OU_Protected_SOU = 0
                }

                if ($OU_Protected -ne "O" -and $OU_Protected -ne "N") {
                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez protéger votre OU de la suppression ou [N] si vous ne voulez pas la protéger."
                }
            }

            # Vérification si l'OU qui est tenté d'être créer est déjà existante

            $OU_Manual_Exist = Get-ADOrganizationalUnit -Filter "Name -eq '$OU_Manual_Create'"

            if ($OU_Manual_Exist) {
                Write-Host "$OU_Manual_Create que vous tentez de créer est déjà existant dans Active Directory."
            }
            
            # Création de l'OU dans l'AD

            else {
                New-ADOrganizationalUnit -Name $OU_Manual_Create -Path "DC=GOTHAM,DC=CH" -ProtectedFromAccidentalDeletion $OU_Protected_SOU
                Write-Host "$OU_Manual_Create à été créer avec succès."
            }
            
            # Boucle pour créer d'autre OU à la main après la première création

            while ($OU_Manual_Create_New_Recreate -eq 0) {
                $OU_Manual_Create_New = Read-Host "Voulez-vous en créer d'autres OUs ? [O/N]"
                
                if ($OU_Manual_Create_New -ne "O" -and $OU_Manual_Create_New -ne "N") {
                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez créer une autre OU ou [N] pour stopper la création."
                }

                if ($OU_Manual_Create_New -eq "N") {
                    $OU_Manual_Create_New_Recreate++
                }

                if ($OU_Manual_Create_New -eq "O") {
                    $OU_Manual_Create_New_Recreate++
                }
            }   
        }
    }
}
