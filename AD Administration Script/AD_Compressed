#####################################################################
#                                                                   #
#      Auteur : Marijan Stajic                                      #
#      Date de création : 02.02.2023                                #
#                                                                   #
#####################################################################
#                                                                   #
#      Ce script permet dexporter lAD en format CSV et de le mettre #
#      dans un fichier ZIP                                          #
#                                                                   #
#####################################################################

# Initalisation des variables pour les chemins

$Cheminc = "C:\\"
$Cheminworkspace = "C:\Workspace"
$Chemincompressed = "C:\Workspace\Compressed_Files\"
$Cheminusers = "C:\Workspace\Compressed_Files\Users_exported.csv"

# Variable contenant les différentes OUs sous forme de tableau

$OUs= @("OU=Finance,DC=gotham,DC=CH", "OU=IT,DC=gotham,DC=CH", "OU=RH,DC=gotham,DC=CH", "OU=Marketing,DC=gotham,DC=CH")

# Création d'une variable pour un tableau vide

$Alladusers = @()

# Vérification si le dossier Workplace existe

if (Test-Path "C:\Workspace\Compressed_Files\") {

    # Vérifie si le fichier ZIP existe déjà, si c'est le cas, il le supprime

    if (Test-Path "C:\Workspace\Compressed_Files\*") {

        Remove-Item "C:\Workspace\Compressed_Files\*" 
        Write-Host "Suppresion du contenu qui ce trouvait dans le dossier Workplace\Compressed_Files."
    }
}

# Création du dossier workplace si il n'existe pas de base 

else {

    New-Item -Path $Cheminc -Name "Workspace" -ItemType Directory
    New-Item -Path $Cheminworkspace  -Name "Compressed_Files" -ItemType Directory 
    Write-Host "Le repértoire Workplace\Compressed_Files à été créer"
}

# Création du fichier CSV qui sera utilisé pour l'exportation 
    
New-Item -Path $Chemincompressed -Name "Users_exported.csv"

Write-Host "Création du fichier Users_exported.csv"

# Boucle Foreach pour pouvoir exporter

foreach ($OU in $OUs) {
    $Users = Get-ADUser -Filter "*" -Properties "*" -SearchBase $OU 

    # Variable qui permets de mettre les donnés de la variable $users avec celle de $allADUsers

    $Alladusers += $Users
}

# Listes des informations demande pour l'export 
 
$Alladusers | Select-Object `
@{Label = "FirstName"; Expression = { $_.GivenName } },
@{Label = "Lastname"; Expression = { $_.Surname } },
@{Label = "Username"; Expression = { $_.SamAccountName } },
@{Label = "Email"; Expression = { $_.UserPrincipalName } },
@{Label = "Initials"; Expression = {$_.Initials}},
@{Label = "JobTitle"; Expression = { $_.Title } },
@{Label = "Department"; Expression = { $_.Department } },
@{Label = "Company"; Expression = { $_.Company } } |

# Export des données recupérer vers le CSV

Export-CSV -Path $Cheminusers -Delimiter ";" -NoTypeInformation

Write-Host "Exportation des données dans le fichier CSV"

# Compression du fichier CSV 

Compress-Archive -Path $Cheminusers -DestinationPath "C:\Workspace\Compressed_Files\Users_exported.zip"

Write-Host "Compression du fichier CSV"

# Suppression du fichier CSV - car deja dans le dossier ZIP

Remove-Item "C:\Workspace\Compressed_Files\Users_exported.csv"
