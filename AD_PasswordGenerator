#####################################################################
#                                                                   #
#      Auteur : Marijan Stajic                                      #
#      Date de création : 17.02.2023                                #
#                                                                   #
#####################################################################
#                                                                   #
#      Ce script permet de générer des mots de passes contenant     #
#      14 caractères dont au minimum 1 majusucle, 1 miniscule,      #
#      1 caractère spécial et 1 chiffre.                            #
#                                                                   #
#####################################################################

# Message pour vérifier si la colonne existe dans le fichier CSV

Write-Host "Le message box pour confirmer si votre CSV contient bien certaines colonnes peut s'ouvrir en arrière-plan. Veuillez vérifier si vous ne le voyez pas."

$PW_CSV_YesNo = [System.Windows.MessageBox]::Show("Il faut que votre fichier contienne la colonne 'Password' si vous souhaitez pouvoir générer des mots de passe avec ce script. Est-ce votre fichier contient cette colonne ?", "Génération de mot de passe", [System.Windows.MessageBoxButton]::YesNo)

# Ouverture de l'explorateur Windows et sélection du fichier CSV

if ($PW_CSV_YesNo -eq [System.Windows.MessageBoxResult]::Yes) {

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
        
        # Permet de vérifier si la colonne Password existe vraiment dans le CSV.

        $CSV_File_Check = Import-Csv -Path $CSV_Path
                
        $CSV_File_Check[0].psobject.properties.name | Out-Null

        if ($CSV_File_Check -match "Password") {

            $CSV_File = Import-Csv -Path $CSV_Path -Delimiter ";" -Encoding Default
            
            # Début du foreach, donc toutes les $PW.Password seront basé sur le CSV_File choisi

            Foreach ($PW in $CSV_File) {
                
                $PW.Password = $Password

                # Initilisation des variables

                $Uppercase_Character = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                $Lowercase_Character = "abcdefghijklmnopqrstuvwxyz"
                $Special_Character = "!#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
                $Numbers = "0123456789"
                $Password_Check = 0
                $Password = ""
                $Password_SpecialC = $False
                $Password_UpperC = $False
                $Password_LowerC = $False
                $Password_Numbers = $False
                    
                    # Génération du mot de passe

                    while ($Password_Check -ne 1) {
                        
                        # Attribution d'une valeur aléatoire à la variable pour le Switch

                        $Password_Random = Get-Random -Minimum 1 -Maximum 5

                        # Par rapport à la valeur de la variable précédente, choix entre les 4 options.

                        switch ($Password_Random) {
                            
                            # Si la variable est égal à 1, la variable pour le mot de passe va prendre au hasard deux caractères spéciaux

                            1 { if ($Password_SpecialC -eq $False) { $Password += $Special_Character.Substring((Get-Random -Minimum 1 -Maximum ($Special_Character.Length - 2)), 2) } }

                            # Si la variable est égal à 2, la variable pour le mot de passe va prendre au hasard cinq caractères en majuscule

                            2 { if ($Password_UpperC -eq $False) { $Password += $Uppercase_Character.Substring((Get-Random -Minimum 1 -Maximum ($Uppercase_Character.Length - 5)), 5) } }

                            # Si la variable est égal à 3, la variable pour le mot de passe va prendre au hasard cinq caractères en miniscule

                            3 { if ($Password_LowerC -eq $False) { $Password += $Lowercase_Character.Substring((Get-Random -Minimum 1 -Maximum ($Lowercase_Character.Length - 5)), 5) } }

                            # Si la variable est égal à 2, la variable pour le mot de passe va prendre au hasard deux numéros

                            4 { if ($Password_Numbers -eq $False) { $Password += $Numbers.Substring((Get-Random -Minimum 1 -Maximum ($Numbers.Length - 2)), 2) } }
                        }

                        # Si le switch a sélectionné le numéro 1, il passe la variable concernant les caractères spéciaux en $True, ce qui rend inutilisable dans le switch

                        if ($Password_Random -eq 1) {

                            $Password_SpecialC = $True
                        }

                        # Si le switch a sélectionné le numéro 2, il passe la variable concernant les caractères en majuscule en $True, ce qui rend inutilisable dans le switch

                        if ($Password_Random -eq 2) {

                            $Password_UpperC = $True
                        }

                        # Si le switch a sélectionné le numéro 3, il passe la variable concernant les caractères en minuscule en $True, ce qui rend inutilisable dans le switch

                        if ($Password_Random -eq 3) {

                            $Password_LowerC = $True
                        }

                        # Si le switch a sélectionné le numéro 4, il passe la variable concernant les numéros en $True, ce qui rend inutilisable dans le switch

                        if ($Password_Random -eq 4) {

                            $Password_Numbers = $True
                        }

                        # Si toutes les valeurs (caractères en majuscule, minuscule, spéciaux et numéro) sont en $True, passe la variable concernant le while à 1 afin de sortir de la boucle

                        if ($Password_SpecialC -and $Password_UpperC -and $Password_LowerC -and $Password_Numbers) {
        
                            $Password_Check++
                        }
                    }            
                }
        
        # Exportation du fichier CSV

        $CSV_File | Export-CSV -Path $CSV_Path -NoTypeInformation -Delimiter ";"

        Write-Host "Des mots de passe alétoire ont été généré et attribuer aux utilisateurs sous la colonne Password du fichier CSV."

        }

        # Si la colonne Password n'est pas présente

        else {

            Write-Host "Votre fichier CSV ne contient pas la colonne Password."
        }
    }
}

# Si la colonne Password n'est pas présente

else {

    Write-Host "Votre fichier CSV ne contient pas la colonne Password."
}
