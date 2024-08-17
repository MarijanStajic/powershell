#####################################################################
#                                                                   #
#      Auteur : Marijan Stajic                                      #
#      Date de création : 15.01.2023                                #
#                                                                   #
#####################################################################
#                                                                   #
#      Ce script permet de créer des Utilisateurs à partir d'un     #
#      fichier CSV ou manuellement.                                 #
#                                                                   #
#####################################################################

# Initilisation des variables

$US_Manual_or_CSV = ""

$US_After_CSV = ""

$US_Manual_Create_New = ""

$US_Manual_Create_New_Recreate = ""

# Début de la boucle pour la création

while ($US_Manual_or_CSV -ne 1 -and $US_Manual_or_CSV -ne 2) {

    # Sélection entre l'importation du CSV ou la création à la main

    $US_Manual_or_CSV = Read-Host "Voulez-vous importer les utilisateur à l'aide d'un fichier CSV [1] ou entrer les utilisateurs à la main [2] ?"
    
    if ($US_Manual_or_CSV -ne 1 -and $US_Manual_or_CSV -ne 2) {

        Write-Host "Erreur ! Veuillez faire [1] pour importer les utilisateur à l'aide d'un fichier CSV ou [2] entrer les utilisateurs à la main."
    }

    # Si l'utilisateur choisi [1] donc création à partir d'un fichier CSV

    if ($US_Manual_or_CSV -eq 1) {
        
        # Message pour vérifier si la colonne existe dans le fichier CSV

        Write-Host "Le message box pour confirmer si votre CSV contient bien certaines colonnes peut s'ouvrir en arrière-plan. Veuillez vérifier si vous ne le voyez pas."

        $US_CSV_YesNo = [System.Windows.MessageBox]::Show("Il faut que votre fichier contienne ces colonnes 'FirstName, Lastname, Username, Initials, JobTitle, Departement, Company et Password' si vous souhaitez pouvoir créer des utilisateurs avec ce script. Est-ce votre fichier contient ces colonnes ?", "Création de d'OU à partir d'un fichier CSV", [System.Windows.MessageBoxButton]::YesNo)

        # Ouverture de l'explorateur Windows et sélection du fichier CSV

        if ($US_CSV_YesNo -eq [System.Windows.MessageBoxResult]::Yes) {

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
                
                # Permet de vérifier si la colonne "FirstName, Lastname, Username, Initials, JobTitle, Departement, Company et Password" existe vraiment dans le CSV. Sinon, il retourne un message disant que non et passe à l'étape suivante.

                $CSV_File_Check = Import-Csv -Path $CSV_Path
                
                $CSV_File_Check[0].psobject.properties.name | Out-Null

                if ($CSV_File_Check -match "FirstName" -and $CSV_File_Check -match "Initials" -and $CSV_File_Check -match "Lastname" -and $CSV_File_Check -match "Username" -and $CSV_File_Check -match "Department" -and $CSV_File_Check -match "Password" -and $CSV_File_Check -match "JobTitle" -and $CSV_File_Check -match "Company") {

                    $CSV_File = Import-Csv -Path $CSV_Path -Delimiter ";" -Encoding Default
                    
                    # Début du foreach, donc toutes les $US.Colonne seront basé sur le CSV_File choisi
                                        
                    foreach($US in $CSV_File) {

                        $Firstname = $US.FirstName
                        $Lastname = $US.Lastname
                        $Username = $US.Username
                        $Initials = $US.Initials
                        $Mail = $US.Email
                        $Jobtitle = $US.JobTitle
                        $Departement = $US.Department
                        $Password = $US.Password
                        $Company = $US.Company

                        # Ajout de "null" si la colonne departement est vide afin d'éviter une erreur avec la variable OU_Exist

                        if ($Departement -eq "") {

                            $Departement = "null"
                        }

                        $OU_Exist = Get-ADOrganizationalUnit -Filter "Name -eq '$Departement'"

                        # Vérification si la colonne Firstname, Lastname, Username ou Departement est vide ou égal à "null"

                        if ($Firstname -eq "" -or $Lastname -eq "" -or $Username -eq "" -or $Password -eq "" -or $Departement -eq "null") {

                            Write-Host "Il est impossible de créer cette utilisateur car la colonne Firstname, Lastname, Username, Password ou Department est incomplète."
                        }

                        # Vérification du SamAccountName

                        elseif (Get-ADUser -Filter {SamAccountName -eq $Username}) {

                            Write-Host "L'identifiant $username existe déjà dans l'Active Directory."
                        }

                        # Vérification si l'OU existe

                        elseif (-not $OU_Exist) {

                            Write-Host "Le département $Departement (OU) n'existe pas dans l'Active Directory. L'utilisateur avec l'identifiant $username n'a donc pas pu être créer."
                        }
                            
                        else {
                            
                            # Passage des lettres en majusucles en miniscules pour l'adresse mail

                            $Firstname_Lower = $Firstname.ToLower() -replace "[éèêë]","E" -replace "[àâä]","A" -replace "[ôöòóœ]","O"
                            $Lastname_Lower = $Lastname.ToLower() -replace "[éèêë]","E" -replace "[àâä]","A" -replace "[ôöòóœ]","O"
                            $Initials -replace "[éèêë]","E" -replace "[àâä]","A" -replace "[ôöòóœ]","O" | Out-Null

                            $Email = "$Firstname_Lower.$Lastname_Lower@opei3a.ch"
                            
                            # Vérification si le nom, prénom est déjà utiliser pour ajouter un numéro à l'adresse mail pour éviter les doublons
                                                        
                            if (Get-ADUser -Filter {(givenName -eq $Firstname) -and (surName -eq $Lastname)}) {
                                
                                For ($Mail_Dispo = 1;Get-ADUser -Filter {UserPrincipalName -eq $Email};$Mail_Dispo++) {
    
                                    $Email = "$Firstname_Lower.$Lastname_Lower$Mail_Dispo@opei3a.ch"
                                }

                                Write-Host "L'utilisateur $Firstname $Lastname est déjà utilisé. Afin de différencier l'adresse mail, un numéro à été ajouté à la fin du nom de famille : $Email"
                            }
                            
                            # Vérification si le nom, prénom est déjà présent dans la même OU pour ajouter un numéro et éviter les erreurs de doublon

                            $Lastname_Check = $Lastname

                            if (Get-ADUser -Filter {(givenName -eq $Firstname)  -and (surName -eq $Lastname_Check)} -SearchBase "OU=$Departement,DC=GOTHAM,DC=CH") {
                            
                                For ($Lastname_Dispo = 1;Get-ADUser -Filter {(givenName -eq $Firstname) -and (surName -eq $Lastname_Check)} -SearchBase "OU=$Departement,DC=GOTHAM,DC=CH";$Lastname_Dispo++) {

                                    $Lastname_Check = $Lastname + $Lastname_Dispo
                                }
                            
                                Write-Host "Le nom et prénom $Firstname $Lastname est déjà utilisé dans l'Active Directory, dans le même département $Departement (OU). Afin de différencier, un numéro à été ajouté à la fin du nom de famille : $Lastname_Check"
                            }
                            
                            # Création de l'utilisateur dans l'AD

                            New-ADUser  -Name "$Lastname_Check $Firstname" -GivenName $Firstname -Surname $Lastname_Check -Initials $Initials -Department $Departement -Company $Company -SamAccountName $Username -UserPrincipalName $Email -Title $Jobtitle -Enabled $true -Path "OU=$Departement,DC=GOTHAM,DC=CH" -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force)
   
                            Write-Host "Création de l'utilisateur : $Username ($Firstname $Lastname)"
                        }
                    }
                }

                # Si les colonnes Firstname, Lastname, etc. ne sont pas présente

                else {

                    Write-Host "Votre fichier CSV ne contient pas les colonnes requises."
                }
            }

            # Boucle pour créer un utilisateur à la main si on le souhaite après l'erreur de colonne

            while ($US_After_CSV -ne "N" -and $US_After_CSV -ne "O") {

                $US_After_CSV = Read-Host "Voulez-vous créer des utilisateurs à la main cette fois-ci ? [O/N]"

                if ($US_After_CSV -eq "O") {

                    $US_Manual_or_CSV = 2
                }
            
                if ($US_After_CSV -ne "O" -and $US_After_CSV -ne "N") {

                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez créer une autre utilisateur à la main ou [N] pour stopper la création."  
                }
            }
        }

        # Boucle pour créer un utilisateur à la main si on le souhaite après l'importation du CSV

        else {

            while ($US_After_CSV -ne "N" -and $US_After_CSV -ne "O") {

                $US_After_CSV = Read-Host "Voulez-vous créer des utilisateurs à la main cette fois-ci ? [O/N]"

                if ($US_After_CSV -eq "O") {

                    $US_Manual_or_CSV = 2
                }
            
                if ($US_After_CSV -ne "O" -and $US_After_CSV -ne "N") {

                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez créer un autre utilisateur à la main ou [N] pour stopper la création."  
                }
            }
        }
    }
    
    # Si l'utilisateur choisi [2] donc création manuel

    if ($US_Manual_or_CSV -eq 2) {
        
        # Boucle pour pouvoir créer des utilisateurs à la main jusqu'à avoir fais "N" pour sortir du script

        while ($US_Manual_Create_New -ne "N") {
            
            # Initilisation des variables

            $Check_Departement = 0
            $Check_Firstname = 0
            $Check_Lastname = 0
            $Username = "opei3a-01"
            $US_Manual_Create_New = ""
            $US_Manual_Create_New_Recreate = 0

            # Boucle pour rentrer le prénom de l'utilisateur et vérifier si il fait bien au minimum un caractère
                      
            while ($Check_Firstname -ne 1) {
                
                $Firstname = Read-Host "Entrez le prénom de l'utilisateur"

                if ($Firstname.Length  -ge 1) {
                    
                    $Check_Firstname++
                }

                else {

                    Write-Host "Le prénom de l'utilisateur doit faire au minimum 1 caractère."
                }
            }

            # Boucle pour rentrer le nom  de l'utilisateur et vérifier si il fait bien au minimum deux caractères

            while ($Check_Lastname -ne 1) {
                
                $Lastname = Read-Host "Entrez le nom de l'utilisateur"

                if ($Lastname.Length -ge 2) {
                    
                    $Check_Lastname++
                }

                else {

                    Write-Host "Le nom de l'utilisateur doit faire au minimum 2 caractères."
                }
            }

            # La raison de pourquoi il faut vérifier si le prénom fait minimum un caractère et le nom deux est parce que les initiales automatique ce base dessus.

            # Initialisation de la variable pour les initiales automatique 

            $Initials = $Firstname.ToUpper().Substring(0,1)+$Lastname.ToUpper().Substring(0,2) -replace "[éèêë]","E" -replace "[àâä]","A" -replace "[ôöòóœ]","O"
            
            # Vérification si l'OU existe

            while ($Check_Departement -ne 1) {

                $Departement = Read-Host "Entrez le nom du département"

                $OU_Manual_Exist = Get-ADOrganizationalUnit -Filter {Name -eq $Departement}

                if ($OU_Manual_Exist) {

                    $Check_Departement++
                }

                else {

                    Write-Host "Le departement $Departement (OU) n'existe pas dans l'Active Directory."                
                }
            }

            # Champs pour rentrer la fonction de l'utilisateur

            $Jobtitle = Read-Host "Entrez le nom de la fonction de l'utilisateur"
            
            # Passage des lettres en majusucles en miniscules pour l'adresse mail 

            $Firstname_Lower = $firstname.ToLower() -replace "[éèêë]","e" -replace "[àâä]","a" -replace "[ôöòóœ]","o"
            $Lastname_Lower = $lastname.ToLower() -replace "[éèêë]","e" -replace "[àâä]","a" -replace "[ôöòóœ]","o"
            $Email = "$Firstname_Lower.$Lastname_Lower@opei3a.ch"
            
            # Vérification du SamAccountName disponible afin de l'attribuer à l'utilisateur

            for ($Username_Dispo = 1;Get-ADUser -Filter {SamAccountName -eq $Username};$Username_Dispo++) {
                
                $Username_Check = $Username_Dispo

                if ($Username_Check -ge 1 -and $Username_Check -le 9) {

                $Username_Check = "0" + $Username_Check
                }

                $Username = "opei3a-$Username_Check"
            }

            # Vérification si le nom, prénom est déjà utiliser pour ajouter un numéro à l'adresse mail pour éviter les doublons
                                                   
            if (Get-ADUser -Filter {(givenName -eq $Firstname) -and (surName -eq $Lastname)}) {
                                
                for ($Mail_Dispo = 1;Get-ADUser -Filter {UserPrincipalName -eq $email};$Mail_Dispo++) {
    
                    $Email = "$Firstname_Lower.$Lastname_Lower$Mail_Dispo@opei3a.ch"
                }

                Write-Host "L'utilisateur $Firstname $Lastname est déjà utilisé. Afin de différencier l'adresse mail, un numéro à été ajouté à la fin du nom de famille : $Email"
            }

            # Vérification si le nom, prénom est déjà présent dans la même OU pour ajouter un numéro et éviter les erreurs de doublon

            $Lastname_Check = $Lastname

            if (Get-ADUser -Filter {(givenName -eq $Firstname) -and (surName -eq $Lastname_Check)} -SearchBase "OU=$Departement,DC=GOTHAM,DC=CH") {
                            
                For ($Lastname_Dispo = 1;Get-ADUser -Filter {(givenName -eq $Firstname) -and (surName -eq $Lastname_Check)} -SearchBase "OU=$Departement,DC=GOTHAM,DC=CH";$Lastname_Dispo++) {

                    $Lastname_Check = $Lastname + $Lastname_Dispo
                }
                            
                Write-Host "Le nom et prénom $Firstname $Lastname est déjà utilisé dans l'Active Directory, dans le même département $Departement (OU). Afin de différencier, un numéro à été ajouté à la fin du nom de famille : $Lastname_Check"
            }
            
            # Création de l'utilisateur dans l'AD

            New-ADUser  -Name "$Lastname_Check $Firstname" -GivenName $Firstname -Surname $Lastname_Check -Initials $Initials -Department $Departement -Company "opei3a SA" -SamAccountName $Username -UserPrincipalName $Email -Title $Jobtitle -Enabled $true -Path "OU=$Departement,DC=GOTHAM,DC=CH" -AccountPassword(Read-Host -AsSecureString "Input Password")
            Write-Host "Création de l'utilisateur : $Username ($Firstname $Lastname)"

            # Boucle pour créer d'autres utilisateurs à la main après la première création

            while ($US_Manual_Create_New_Recreate -eq 0) {

                $US_Manual_Create_New = Read-Host "Voulez-vous en créer d'autres utilisateurs ? [O/N]"
                
                if ($US_Manual_Create_New -ne "O" -and $US_Manual_Create_New -ne "N") {

                    Write-Host "Erreur ! Veuillez faire [O] si vous souhaitez créer une autre OU ou [N] pour stopper la création."
                }

                if ($US_Manual_Create_New -eq "N") {

                    $US_Manual_Create_New_Recreate++
                }

                if ($US_Manual_Create_New -eq "O") {

                    $US_Manual_Create_New_Recreate++
                }
            }
        }
    } 
}
