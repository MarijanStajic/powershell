#####################################################################
#                                                                   #
#      Auteur : Marijan Stajic                                      #
#      Date de création : 12.02.2023                                #
#                                                                   #
#####################################################################
#                                                                   #
#      Ce script permet d'attribuer des dates d'expiration          #
#      aléatoirement aux utilisateurs "opei3a-01 ==> -30".          #
#                                                                   #
#####################################################################

# Vérification si il existe 30 utilisateurs

$All_User = Get-ADUser -Filter {SamAccountName -like "opei3a-*"}

$All_User_Number = $All_User.Count

if ($All_User_Number -lt 30) {

    Write-Host "L'Active Directory contient actullement $All_User_Number utilisateurs. Il en faut minimum 30 pour démarrer le script."
    Return
}

# Création de l'OU "Retired"

$Retired_Exist = Get-ADOrganizationalUnit -Filter "Name -eq 'Retired'"

if ($Retired_Exist) {
    
    Write-Host "L'unité d'organisation Retired existe déjà."
}

else {
    
    New-ADOrganizationalUnit -Name "Retired" -Path "DC=GOTHAM,DC=CH" -ProtectedFromAccidentalDeletion $False
    Write-Host "Création de l'unité d'organisation Retired"
}

# Activation de la date d'expiration au 31/12/2023 alétoirement pour 10 utilisateurs

for ($Username_Dispo = 1;$Username_Dispo -le 10; $Username_Dispo++) {
    
    # Initilisation des variables

    $Username_Random_Dispo = 0

    $Identity_All_Validation = $true

    # Vérification si les 30 utilisateurs n'ont pas déjà une date d'expiration

    for ($Identity_Check = 1; $Identity_Check -le 30; $Identity_Check++) {
        
        # Initialisation de la variable afin d'avoir le SamAccountName correct

        $Identity_All = $Identity_Check

        if ($Identity_All -ge 1 -and $Identity_All -le 9) {

            $Identity_All = "0" + $Identity_All
        }

        $Identity_Username = "opei3a-$Identity_All"

        # Vérification si l'utilisateur n'a pas déjà de date d'expiration

        $Username_Check = Get-ADUser -Identity $Identity_Username -Properties AccountExpirationDate

        if (-not $Username_Check.AccountExpirationDate) {
                
            $Identity_All_Validation = $false
        }
    }

    # Si c'est le cas, passage de la variable d'incérmentation du For à 10 directement afin de couper la boucle

    if ($Identity_All_Validation) {

        $Username_Dispo = 10
    }

    else {

    # Si aucun utilisateur n'a de date d'expiration, lancement de l'attribution de la date sur un utilisateur aléatoire

        while ($Username_Random_Dispo -ne 1) {  
        
            # Attribution aléatoire et initialisation de la variable afin d'avoir le SamAccountName correct

            $Username_Random = Get-Random -Minimum 1 -Maximum 31 

            if ($Username_Random -ge 1 -and $Username_Random -le 9) {

                $Username_Random = "0" + $Username_Random
            }

            $Username = "opei3a-$Username_Random"

            # Vérification si l'utilisateur n'a pas déjà de date d'expiration

            $Username_Check = Get-ADUser -Identity $Username -Properties AccountExpirationDate
            
            if (-not $Username_Check.AccountExpirationDate) {

                $Username_Random_Dispo++
            }
        }

        # Attribution et message pour informer que la date d'expiration à été ajouté à un utilisateur aléatoire 

        Write-Host "La date d'expiration de l'utilisateur $Username à été fixé au 31/12/2023."

        Set-ADAccountExpiration -Identity $Username -DateTime "31/12/2023"
    }
}

# Activation de la date d'expiration au 30/06/2023 alétoirement pour 10 utilisateurs

for ($Username_Dispo = 1;$Username_Dispo -le 10; $Username_Dispo++) {
    
    # Initilisation des variables

    $Username_Random_Dispo = 0

    $Identity_All_Validation = $true

    # Vérification si les 30 utilisateurs n'ont pas déjà une date d'expiration

    for ($Identity_Check = 1; $Identity_Check -le 30; $Identity_Check++) {
        
        # Initialisation de la variable afin d'avoir le SamAccountName correct

        $Identity_All = $Identity_Check

        if ($Identity_All -ge 1 -and $Identity_All -le 9) {

            $Identity_All = "0" + $Identity_All
        }

        $Identity_Username = "opei3a-$Identity_All"

        # Vérification si l'utilisateur n'a pas déjà de date d'expiration

        $Username_Check = Get-ADUser -Identity $Identity_Username -Properties AccountExpirationDate

        if (-not $Username_Check.AccountExpirationDate) {
                
            $Identity_All_Validation = $false
        }
    }

    # Si c'est le cas, passage de la variable d'incérmentation du For à 10 directement afin de couper la boucle

    if ($Identity_All_Validation) {

        $Username_Dispo = 10
    }

    else {

    # Si aucun utilisateur n'a de date d'expiration, lancement de l'attribution de la date sur un utilisateur aléatoire

        while ($Username_Random_Dispo -ne 1) {  
        
            # Attribution aléatoire et initialisation de la variable afin d'avoir le SamAccountName correct

            $Username_Random = Get-Random -Minimum 1 -Maximum 31 

            if ($Username_Random -ge 1 -and $Username_Random -le 9) {

                $Username_Random = "0" + $Username_Random
            }

            $Username = "opei3a-$Username_Random"

            # Vérification si l'utilisateur n'a pas déjà de date d'expiration

            $Username_Check = Get-ADUser -Identity $Username -Properties AccountExpirationDate
            
            if (-not $Username_Check.AccountExpirationDate) {

                $Username_Random_Dispo++
            }
        }

        # Attribution et message pour informer que la date d'expiration à été ajouté à un utilisateur aléatoire 

        Write-Host "La date d'expiration de l'utilisateur $Username à été fixé au 30/06/2023."

        Set-ADAccountExpiration -Identity $Username -DateTime "30/06/2023"
    }
}

# Activation de la date d'expiration au 30/11/2022 alétoirement pour 5 utilisateurs

for ($Username_Dispo = 1;$Username_Dispo -le 5; $Username_Dispo++) {
    
    # Initilisation des variables

    $Username_Random_Dispo = 0

    $Identity_All_Validation = $true

    # Vérification si les 30 utilisateurs n'ont pas déjà une date d'expiration

    for ($Identity_Check = 1; $Identity_Check -le 30; $Identity_Check++) {
        
        # Initialisation de la variable afin d'avoir le SamAccountName correct

        $Identity_All = $Identity_Check

        if ($Identity_All -ge 1 -and $Identity_All -le 9) {

            $Identity_All = "0" + $Identity_All
        }

        $Identity_Username = "opei3a-$Identity_All"

        # Vérification si l'utilisateur n'a pas déjà de date d'expiration

        $Username_Check = Get-ADUser -Identity $Identity_Username -Properties AccountExpirationDate

        if (-not $Username_Check.AccountExpirationDate) {
                
            $Identity_All_Validation = $false
        }
    }

    # Si c'est le cas, passage de la variable d'incérmentation du For à 10 directement afin de couper la boucle

    if ($Identity_All_Validation) {

        $Username_Dispo = 10
    }

    else {

    # Si aucun utilisateur n'a de date d'expiration, lancement de l'attribution de la date sur un utilisateur aléatoire

        while ($Username_Random_Dispo -ne 1) {  
        
            # Attribution aléatoire et initialisation de la variable afin d'avoir le SamAccountName correct

            $Username_Random = Get-Random -Minimum 1 -Maximum 31 

            if ($Username_Random -ge 1 -and $Username_Random -le 9) {

                $Username_Random = "0" + $Username_Random
            }

            $Username = "opei3a-$Username_Random"

            # Vérification si l'utilisateur n'a pas déjà de date d'expiration

            $Username_Check = Get-ADUser -Identity $Username -Properties AccountExpirationDate
            
            if (-not $Username_Check.AccountExpirationDate) {

                $Username_Random_Dispo++
            }
        }

        # Attribution et message pour informer que la date d'expiration à été ajouté à un utilisateur aléatoire 

        Write-Host "La date d'expiration de l'utilisateur $Username à été fixé au 30/11/2022."

        Set-ADAccountExpiration -Identity $Username -DateTime "30/11/2022"
    }
}

# Activation de la date d'expiration au 31/07/2022 alétoirement pour 5 utilisateurs

for ($Username_Dispo = 1;$Username_Dispo -le 5; $Username_Dispo++) {
    
    # Initilisation des variables

    $Username_Random_Dispo = 0

    $Identity_All_Validation = $true

    # Vérification si les 30 utilisateurs n'ont pas déjà une date d'expiration

    for ($Identity_Check = 1; $Identity_Check -le 30; $Identity_Check++) {
        
        # Initialisation de la variable afin d'avoir le SamAccountName correct

        $Identity_All = $Identity_Check

        if ($Identity_All -ge 1 -and $Identity_All -le 9) {

            $Identity_All = "0" + $Identity_All
        }

        $Identity_Username = "opei3a-$Identity_All"

        # Vérification si l'utilisateur n'a pas déjà de date d'expiration

        $Username_Check = Get-ADUser -Identity $Identity_Username -Properties AccountExpirationDate

        if (-not $Username_Check.AccountExpirationDate) {
                
            $Identity_All_Validation = $false
        }
    }

    # Si c'est le cas, passage de la variable d'incérmentation du For à 10 directement afin de couper la boucle

    if ($Identity_All_Validation) {
            
        $Username_Dispo = 10
    }

    else {

    # Si aucun utilisateur n'a de date d'expiration, lancement de l'attribution de la date sur un utilisateur aléatoire

        while ($Username_Random_Dispo -ne 1) {  
        
            # Attribution aléatoire et initialisation de la variable afin d'avoir le SamAccountName correct

            $Username_Random = Get-Random -Minimum 1 -Maximum 31 

            if ($Username_Random -ge 1 -and $Username_Random -le 9) {

                $Username_Random = "0" + $Username_Random
            }

            $Username = "opei3a-$Username_Random"

            # Vérification si l'utilisateur n'a pas déjà de date d'expiration

            $Username_Check = Get-ADUser -Identity $Username -Properties AccountExpirationDate
            
            if (-not $Username_Check.AccountExpirationDate) {

                $Username_Random_Dispo++
            }
        }

        # Attribution et message pour informer que la date d'expiration à été ajouté à un utilisateur aléatoire 

        Write-Host "La date d'expiration de l'utilisateur $Username à été fixé au 31/07/2022."

        Set-ADAccountExpiration -Identity $Username -DateTime "31/07/2022"
    }
}

# Message si tout les utilisateurs ont déjà une date d'expiration

if ($Identity_All_Validation) {

    Write-Host "Tous les utilisateurs ont déjà une date d'expiration."
}

# Initlisation de la variable pour avoir la date d'aujourd'hui

$Date = Get-Date -Format "dd/MM/yyyy"

# Déplacement des utilisateurs inactifs par rapport à la date

Get-ADUser -Filter {AccountExpirationDate -lt $Date} | Move-ADObject -TargetPath "OU=Retired,DC=GOTHAM,DC=CH"

Write-Host "Les utilisateurs avec la date antérieure au $Date ont été déplacé dans l'OU Retired."
