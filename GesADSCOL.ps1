#
# Ensemble d'outil pour la gestion d'un annuaire Active Directory (type scolaire ou petite organisation)
# 
# Fonctions principales du script (en tout cas les plus importantes selon moi) : 
#
#  • **Import CSV** : création de plusieurs comptes à partir d’un fichier
#    CSV.
#
#  • **Ajout interactif** : création d’un élève via des invites `Read‑Host`.
#    Le cmdlet `Read‑Host` est utilisé avec les options `-Prompt` et
#    `-AsSecureString` pour collecter des informations sensibles comme
#    le mot de passe. Les identifiants sont
#    générés automatiquement (SamAccountName et UPN) en suivant les
#    recommandations du forum LazyAdmin sur le lien en README.
#
#
# Auteur : S. COSTENOBLE, P. PERROT et toutes les autres sources cités en README
# Version : 0.9.1b
# Date de mise à jour en 0.9.1b : 15 septembre 2025
# Date de création initiale : 19 août 2025 (correspond à la date de la première version "dev")
# Description des fonctions et réorganisation du code par IA

<#
    UTILISATION
    ==========
    • Exécutez ce script avec des privilèges d’administrateur sur un serveur
      membre du domaine disposant du module ActiveDirectory ou des outils RSAT.
    • Choisissez l’ajout interactif pour créer un élève à la fois, ou
      l’import CSV pour créer plusieurs comptes.
    • Adaptez les constantes `$Domaine`, la structure des OU et les groupes
      selon votre environnement scolaire.
    • Le fichier CSV doit contenir au minimum les colonnes `GivenName` (prénom),
      `Surname` (nom) et `CodeClasse` (code de la classe). D’autres colonnes
      facultatives sont prises en charge (SamAccountName, Password, Path,
      Department, Description, Groups).
    • Exemple de ligne CSV :
        GivenName;Surname;CodeClasse;Password;Department;Groups
        Marie;Dupont;2A;Bienvenue2025!;Sciences;"2NDE,ClubTheatre"
#>

param()

# -----------------------------------------------------------------------------
# Variables globales et constantes
# -----------------------------------------------------------------------------

## Domaine DNS utilisé pour l’UPN des comptes élèves
$Domaine = 'exemple.lan'

## Année scolaire (utilisée dans la description des comptes)
$AnneeScolaire = '2025-2026'

## Fonction pour calculer l'année scolaire suivante à partir de $AnneeScolaire.
function Get-NextSchoolYear {
    <#
        .SYNOPSIS
            Calcule l'année scolaire suivante (par ex. 2024‑2025 → 2025‑2026).

        .DESCRIPTION
            Sépare la chaîne sur le tiret, convertit les deux parties en entiers et
            incrémente chacune d’une année. Si le format ne correspond pas à
            'YYYY‑YYYY', la même valeur est retournée.

        .PARAMETER CurrentYear
            L’année scolaire actuelle.

        .OUTPUTS
            Chaîne représentant l’année scolaire suivante.
    #>
    param([string]$CurrentYear)
    $parts = $CurrentYear -split '-'
    if ($parts.Count -eq 2 -and ($parts[0] -as [int]) -and ($parts[1] -as [int])) {
        $start = [int]$parts[0] + 1
        $end   = [int]$parts[1] + 1
        return "$start-$end"
    } else {
        return $CurrentYear
    }
}

## Liste de codes classes BTS pour aiguiller l’orientation de l’OU
## A ADAPTER SELON VOTRE SITUATION OU A LAISSER EN L'ETAT SI VOUS N'AVEZ PAS DE CLASSES DE BTS
$BTSClasses = @('BTS1M','BTS2M','BTS1G','BTS2G','BTS1CCST','BTS2CCST')

## OU racine des élèves utilisée pour les recherches lors des promotions et réorganisations
## DOSSIER PAR DÉFAUT POUR LES ÉLÈVES EN GROS
$ElevesOU = 'OU=Eleves,OU=Utilisateurs,DC=estic,DC=peda'

# -----------------------------------------------------------------------------
# Variables et fonctions supplémentaires pour la gestion du publipostage et la
# normalisation des chaînes. Ces éléments sont utilisés pour générer des
# rapports CSV des comptes créés et pour supprimer les accents tout en
# convertissant les prénoms et noms en majuscules.
# -----------------------------------------------------------------------------

# Liste globale pour stocker les comptes créés via l'import CSV
$global:UtilisateursCreesImport = @()
# Liste globale pour stocker les comptes créés via la saisie interactive
$global:UtilisateursCreesInteractive = @()
# Chemin du fichier de publipostage pour l'import CSV
$global:PublipostageImportPath = ''
# Chemin du fichier de publipostage pour la saisie interactive
$global:PublipostageInteractivePath = ''

# -----------------------------------------------------------------------------
# Variables et listes pour la gestion du personnel (profs, AESH, surveillants…)
# -----------------------------------------------------------------------------

# Liste globale pour stocker les comptes adultes créés via l'import CSV
$global:StaffCreesImport = @()
# Liste globale pour stocker les comptes adultes créés via la saisie interactive
$global:StaffCreesInteractive = @()
# Chemin du fichier de publipostage pour l'import CSV des personnels
$global:PublipostageStaffImportPath = ''
# Chemin du fichier de publipostage pour la saisie interactive des personnels
$global:PublipostageStaffInteractivePath = ''

function ConvertTo-PlainText {
    <#
        .SYNOPSIS
            Convertit un objet SecureString en texte clair.

        .DESCRIPTION
            Utilise la classe NetworkCredential pour extraire la chaîne de caractères
            lisible d'un SecureString afin de pouvoir l'exporter dans un fichier CSV.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$SecureString
    )
    # Ne pas écrire le mot de passe en clair dans la console, seulement le retourner
    return ([System.Net.NetworkCredential]::new('', $SecureString)).Password
}

function Normalize-String {
    <#
        .SYNOPSIS
            Supprime les accents d'une chaîne et la convertit en majuscules.

        .DESCRIPTION
            Les comptes AD nécessitent souvent des identifiants sans accents pour éviter
            des incohérences de saisie. Cette fonction décompose d'abord la chaîne en
            FormD (caractères de base + signes diacritiques), filtre les signes
            diacritiques (NonSpacingMark) puis recompose la chaîne en FormC en
            majuscules.

        .PARAMETER InputString
            Chaîne à normaliser.

        .OUTPUTS
            Chaîne normalisée en majuscules sans accents.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputString
    )
    if ([string]::IsNullOrWhiteSpace($InputString)) {
        return $InputString
    }
    # Convertir en majuscules invariantes et décomposer
    $normalized = $InputString.ToUpperInvariant().Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object -TypeName System.Text.StringBuilder
    foreach ($c in $normalized.ToCharArray()) {
        # Exclure les marques diacritiques (NonSpacingMark)
        $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

# -----------------------------------------------------------------------------
# Fonctions utilitaires
# -----------------------------------------------------------------------------

function Import-ADModule {
    <#
        .SYNOPSIS
            Importe le module ActiveDirectory ou l’installe si nécessaire.

        .DESCRIPTION
            Vérifie si le module `ActiveDirectory` est chargé. S’il est absent,
            tente de l’importer depuis les modules disponibles ou de l’installer
            via la fonctionnalité RSAT (`Add-WindowsFeature`). N'est pas nécessaire si exécuter sur un AD par exemple.
    #>
    if (-not (Get-Module -Name ActiveDirectory)) {
        if (Get-Module -ListAvailable | Where-Object { $_.Name -eq 'ActiveDirectory' }) {
            Import-Module -Name ActiveDirectory -ErrorAction Stop
        } else {
            Write-Warning "Module ActiveDirectory non trouvé. Tentative d’installation via RSAT…"
            if (Get-Command -Name Add-WindowsFeature -ErrorAction SilentlyContinue) {
                try {
                    Add-WindowsFeature -Name 'RSAT-AD-PowerShell' -IncludeAllSubFeature | Out-Null
                    Import-Module -Name ActiveDirectory -ErrorAction Stop
                } catch {
                    throw "Impossible d’installer le module ActiveDirectory : $($_.Exception.Message)"
                }
            } else {
                throw "La cmdlet Add-WindowsFeature n’est pas disponible. Installez les outils RSAT manuellement."
            }
        }
    }
}

function New-RandomPassword {
    <#
        .SYNOPSIS
            Génère un mot de passe aléatoire.

        .PARAMETER Length
            Longueur totale du mot de passe.
        .PARAMETER Uppercase
            Nombre minimum de lettres majuscules.
        .PARAMETER Digits
            Nombre minimum de chiffres.
        .PARAMETER SpecialCharacters
            Nombre minimum de caractères spéciaux.

        .OUTPUTS
            Chaîne contenant le mot de passe généré.

        .DESCRIPTION
        Vous pouvez également personnaliser les caractères qui peuvent être sélectionner afin de ne pas avoir de caractères proches (0 ou O par exemple).
    #>
    param(
        [int]$Length = 5,
        [int]$Uppercase = 1,
        [int]$Digits = 1,
        [int]$SpecialCharacters = 1
    )
    $Lowercase = $Length - $SpecialCharacters - $Uppercase - $Digits
    if ($Lowercase -lt 1) { throw "Longueur insuffisante pour les contraintes spécifiées." }
    $ArrayLower = @('a','b','c','d','e','f','g','h','j','k','l','m','n','p','q','r','s','t','u','v','w','x','y','z')
    $ArrayUpper = @('A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z')
    $ArraySpecial = @('*','$','%','?','!','@','#')

    # Compose chaque partie
    $pwd  = -join ($ArrayLower | Get-Random -Count $Lowercase)
    $pwd += -join ((0..9) | Get-Random -Count $Digits)
    $pwd += -join ($ArrayUpper | Get-Random -Count $Uppercase)
    $pwd += -join ($ArraySpecial | Get-Random -Count $SpecialCharacters)
    # Mélange aléatoire
    $chars = $pwd.ToCharArray()
    $final = $chars | Get-Random -Count $chars.Length
    return -join $final
}

function New-SamAccountName {
    <#
        .SYNOPSIS
            Génère un SamAccountName à partir du prénom et du nom.

        .DESCRIPTION
            Concatène la première lettre du prénom et le nom en supprimant les
            espaces, tirets et apostrophes. Si le résultat dépasse 20 caractères,
            il est tronqué. Technique décrite par LazyAdmin【949893158897911†L162-L173】.

        .PARAMETER GivenName
            Prénom de l’utilisateur.
        .PARAMETER Surname
            Nom de l’utilisateur.
        .PARAMETER MaxLength
            Longueur maximale autorisée (par défaut 20, sinon je pense que la personne se suicide rien qu'en tapant son nom d'utilisateur tout les 4 matins).
    #>
    param(
        [string]$GivenName,
        [string]$Surname,
        [int]$MaxLength = 20
    )
    $base = ($GivenName.Substring(0,1) + $Surname).ToLower() -replace "[-' ]"
    if ($base.Length -gt $MaxLength) {
        return $base.Substring(0,$MaxLength)
    }
    return $base
}

function New-UserPrincipalName {
    <#
        .SYNOPSIS
            Génère l’UPN à partir du SamAccountName et du domaine.
    #>
    param(
        [string]$SamAccountName,
        [string]$Domain
    )
    return "$SamAccountName@$Domain"
}

function Get-OU {
    <#
        .SYNOPSIS
            Retourne l’OU appropriée en fonction du code classe.

        .DESCRIPTION
            Utilise des expressions régulières pour déterminer l’emplacement
            d’un élève dans l’arborescence Active Directory. Si aucune
            correspondance n’est trouvée, l’OU par défaut des élèves est
            retournée.
    #>
    param([string]$CodeClasse)
    switch -Regex ($CodeClasse) {
        '^BTS' { return 'OU=BTS,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^6'  { return 'OU=College,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^5'  { return 'OU=College,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^4'  { return 'OU=College,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^3PM' { return 'OU=Lycee,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^3'  { return 'OU=College,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^2'  { return 'OU=Lycee,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^1'  { return 'OU=Lycee,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        '^T'  { return 'OU=Lycee,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        default { return 'OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' } #<-# Pour la prochaine version mettre la variable global par défaut de l'OU élève. Et pour éviter de tout retaper mettre le nom de domaine en variable global aussi (mon esprit flemmard me remerciera).
    }
}

function Get-GeneralGroup {
    <#
        .SYNOPSIS
            Retourne le nom du groupe général à partir du code classe.
            Dans le cas présent j'ai simplement nommé les groupes comme les codes classes mais on peut facilement faire autrement selon sa config via cet encadré.
    #>
    param([string]$CodeClasse)
    switch -Regex ($CodeClasse) {
        '^6'  { return '6EME' }
        '^5'  { return '5EME' }
        '^4'  { return '4EME' }
        '^3'  { return '3EME' }
        '^2'  { return '2NDE' }
        '^1'  { return '1ERE' }
        '^T'  { return 'TERMINALE' }
        '^BTS1M' { return 'BTS1M' }
        '^BTS2M' { return 'BTS2M' }
        '^BTS1G' { return 'BTS1G' }
        '^BTS2G' { return 'BTS2G' }
        '^BTS1CCST' { return 'BTS1CCST' }
        '^BTS2CCST' { return 'BTS2CCST' }
        default { return 'Inconnu' }
    }
}

function Get-StaffOU {
    <#
        .SYNOPSIS
            Retourne l’OU appropriée pour un membre du personnel selon son rôle.

        .DESCRIPTION
            Les adultes (personnels) sont organisés dans différentes OUs en
            fonction de leur statut : Extérieur, Personnel, Prof, Surveillant
            ou AESH. Cette fonction convertit le rôle en majuscules et
            sélectionne l’OU correcte dans l’arborescence Active Directory.

        .PARAMETER Role
            Catégorie du personnel (Extérieur, Personnel, Prof, Surveillant, AESH).

        .OUTPUTS
            Chaîne représentant l’OU cible.
    #>
    param([string]$Role)
    $r = $Role.ToUpperInvariant()
    switch ($r) {
        'EXTERIEUR' { return 'OU=Ext,OU=Profs,OU=Utilisateurs,DC=exemple,DC=com' }
        'PERSONNEL' { return 'OU=Personnels,OU=Profs,OU=Utilisateurs,DC=exemple,DC=com' }
        'PROF'      { return 'OU=Profs,OU=Utilisateurs,DC=exemple,DC=com' }
        'PROFS'     { return 'OU=Profs,OU=Utilisateurs,DC=exemple,DC=com' }
        'SURVEILLANT' { return 'OU=Surveillants,OU=Profs,OU=Utilisateurs,DC=exemple,DC=com' }
        'AESH'      { return 'OU=AESH,OU=Eleves,OU=Utilisateurs,DC=exemple,DC=com' }
        default     { return 'OU=Utilisateurs,DC=exemple,DC=com' } #<-# Pareil que pour Get-OU oskour.
    }
}

function Get-StaffGroup {
    <#
        .SYNOPSIS
            Retourne le groupe AD correspondant au rôle d’un adulte.

        .DESCRIPTION
            Associe chaque catégorie de personnel à un groupe spécifique afin
            de pouvoir gérer les droits et appartenances plus facilement. Les
            noms de groupes doivent exister dans Active Directory.

        .PARAMETER Role
            Catégorie du personnel.

        .OUTPUTS
            Nom du groupe correspondant.
    #>
    param([string]$Role)
    $r = $Role.ToUpperInvariant()
    switch ($r) {
        'EXTERIEUR' { return 'Extérieur' }
        'PERSONNEL' { return 'Personnel' }
        'PROF'      { return 'Profs' }
        'PROFS'     { return 'Profs' }
        'SURVEILLANT' { return 'Surveillant' }
        'AESH'      { return 'AESH' }
        default     { return 'Personnel' } #<-# Ici j'ai décidé de mettre "Personnel" par défaut mais possibilité d'ajuster au besoin.
    }
}

function New-SchoolADUser {
    <#
        .SYNOPSIS
            Crée un compte utilisateur dans Active Directory.

        .PARAMETER UserData
            Hashtable contenant les propriétés nécessaires à `New-ADUser` :
            GivenName, Surname, Name, DisplayName, SamAccountName,
            UserPrincipalName, Path, AccountPassword, Enabled,
            ChangePasswordAtLogon, Description, Department, etc.
        .PARAMETER WhatIf
            Si spécifié, exécute la cmdlet en mode simulation (`-WhatIf`).
    #>
    param(
        [hashtable]$UserData,
        [switch]$WhatIf
    )
    try {
        if ($PSBoundParameters.ContainsKey('WhatIf') -and $WhatIf) {
            New-ADUser @UserData -WhatIf -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -ErrorAction Stop
        } else {
            New-ADUser @UserData -ChangePasswordAtLogon $false -PasswordNeverExpires $true -CannotChangePassword $true -ErrorAction Stop
        }
        Write-Host "Compte créé : $($UserData.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Host "Erreur lors de la création du compte $($UserData.DisplayName)" -ForegroundColor Red
        Write-Host "Détails : $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Add-UserInteractive {
    <#
        .SYNOPSIS
            Ajoute un utilisateur en saisie interactive.

        .DESCRIPTION
            Demande les informations essentielles de l’élève : prénom, nom et
            code de classe. Calcule ensuite le SamAccountName et l’UPN,
            choisit l’OU et le groupe général en fonction du code de classe,
            génère ou collecte un mot de passe et crée le compte si aucun
            doublon n’existe.

        .PARAMETER WhatIf
            Permet de simuler la création de l’utilisateur sans l’effectuer.
    #>
    param(
        [switch]$WhatIf
    )
    Import-ADModule

    # Réinitialiser la liste interactive avant de commencer afin de ne pas cumuler les enregistrements
    # de publipostage entre plusieurs appels interactifs. Chaque ajout interactif génère un
    # nouveau fichier de publipostage isolé.
    $global:UtilisateursCreesInteractive = @()
    Write-Host "Ajout interactif d’un élève" -ForegroundColor Cyan
    $givenName = Read-Host -Prompt "Prénom"
    $surname   = Read-Host -Prompt "Nom"
    $codeClasse = Read-Host -Prompt "Code de classe (ex : 3C, BTS1G)"

    # Normaliser le prénom et le nom (suppression des accents et conversion en majuscules)
    $givenName = Normalize-String -InputString $givenName
    $surname   = Normalize-String -InputString $surname

    # Générer un mot de passe ou demander la saisie
    $pwChoice = Read-Host -Prompt "Voulez‑vous saisir le mot de passe ? (O/N)"
    [System.Security.SecureString]$passwordSecure = $null
    if ($pwChoice -match '^[oOyY]') {
        $passwordSecure = Read-Host -Prompt 'Mot de passe' -AsSecureString
    } else {
        $generated = New-RandomPassword
        Write-Host "Mot de passe généré : $generated" -ForegroundColor Magenta
        $passwordSecure = ConvertTo-SecureString -String $generated -AsPlainText -Force
    }

    # Générer l’identifiant
    $sam = New-SamAccountName -GivenName $givenName -Surname $surname
    $upn = New-UserPrincipalName -SamAccountName $sam -Domain $Domaine
    $displayName = "$givenName $surname"

    # Déterminer l’OU et le groupe
    $ou = Get-OU -CodeClasse $codeClasse
    $group = Get-GeneralGroup -CodeClasse $codeClasse
    # Calcule le niveau général (6EME, 5EME, etc.) à partir du code de classe
    $niveau = Get-GeneralGroup -CodeClasse $codeClasse
    # Si aucun niveau n’est trouvé, on conserve le code complet ; sinon on l’utilise dans la description
    $description = if ($niveau -eq 'Inconnu') { "$codeClasse $AnneeScolaire" } else { "$niveau $AnneeScolaire" }

    # Vérifier l’existence du compte【949893158897911†L210-L217】
    if (Get-ADUser -Filter { SamAccountName -eq $sam -or UserPrincipalName -eq $upn } -ErrorAction SilentlyContinue) {
        Write-Host "Un utilisateur avec l’identifiant $sam ou l’UPN $upn existe déjà." -ForegroundColor Yellow
        return
    }

    # Construction du hashtable pour New-ADUser
    $userProps = @{
        GivenName             = $givenName
        Surname               = $surname
        Name                  = $displayName
        DisplayName           = $displayName
        SamAccountName        = $sam
        UserPrincipalName     = $upn
        Path                  = $ou
        AccountPassword       = $passwordSecure
        Enabled               = $true
        Description           = $description
    }
    New-SchoolADUser -UserData $userProps -WhatIf:$WhatIf

    # Ajout au groupe général si connu
    if ($group -ne 'Inconnu') {
        try {
            if ($WhatIf) {
                Add-ADGroupMember -Identity $group -Members $sam -WhatIf
            } else {
                Add-ADGroupMember -Identity $group -Members $sam
            }
            Write-Host "Ajouté au groupe $group" -ForegroundColor Cyan
        } catch {
            Write-Host "Impossible d’ajouter au groupe $group : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Aucun groupe général pour le code de classe $codeClasse" -ForegroundColor Yellow
    }

    # -- Publipostage interactif et affichage supplémentaire --
    # Demander un chemin de publipostage la première fois ou utiliser celui déjà défini
    if (-not $global:PublipostageInteractivePath -or $global:PublipostageInteractivePath -eq '') {
        $pubPath = Read-Host -Prompt "Chemin du fichier de publipostage (laisser vide pour un nom automatique)"
        if ([string]::IsNullOrWhiteSpace($pubPath)) {
            $timestampPub = Get-Date -Format 'yyyyMMdd_HHmmss'
            $pubPath = "publipostage_interactif_$timestampPub.csv"
        }
        $global:PublipostageInteractivePath = $pubPath
    }
    # Convertir le mot de passe pour l’export
    $plainPwdInt = ConvertTo-PlainText -SecureString $passwordSecure
    # Préparer l’objet à exporter
    $utilObjInt = [PSCustomObject]@{
        NomPrenom      = $displayName
        NomUtilisateur = $sam
        MotDePasse     = $plainPwdInt
        Classe         = $codeClasse
    }
    $global:UtilisateursCreesInteractive += $utilObjInt
    # Informer l’administrateur et afficher l’identifiant
    Write-Host "$displayName a été ajouté(e) à la liste des utilisateurs pour le publipostage" -ForegroundColor Cyan
    Write-Host "SamAccountName : $sam" -ForegroundColor Cyan
    # Exporter la liste de publipostage pour l’interactif en UTF‑8 (conserver les accents)
    try {
        $global:UtilisateursCreesInteractive | Export-Csv -Path $global:PublipostageInteractivePath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    } catch {
        Write-Host "Impossible d’exporter le fichier de publipostage interactif : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Après export, vider la liste interactive afin d’éviter une accumulation lors des prochains appels
    $global:UtilisateursCreesInteractive = @()
}

function Add-UsersFromCsv {
    <#
        .SYNOPSIS
            Crée des comptes à partir d’un fichier CSV.

        .PARAMETER CSVPath
            Chemin vers le fichier CSV à importer (délimité par point-virgule).
        .PARAMETER WhatIf
            Active le mode simulation pour `New-ADUser` et `Add-ADGroupMember`.

        .DESCRIPTION
            Importe les données d’un fichier CSV, dérive les identifiants si
            nécessaire, vérifie l’existence des comptes et crée les
            utilisateurs. Les colonnes suivantes sont prises en charge :
            - GivenName (obligatoire)
            - Surname (obligatoire)
            - CodeClasse (obligatoire)
            - SamAccountName (facultatif)
            - UserPrincipalName (facultatif)
            - Password (facultatif)
            - Path (OU spécifique, facultatif)
            - Department, Description (facultatifs)
            - Groups (liste de groupes séparés par des virgules)
    #>
    param(
        [string]$CSVPath,
        [switch]$WhatIf,
        # Permet de spécifier l’encodage du fichier CSV. Par défaut UTF8. Vous pouvez
        # utiliser d’autres encodages (utf8BOM, Default, OEM, ansi, etc.) si le
        # fichier a été enregistré avec un encodage différent.
        [string]$Encoding = 'UTF8'
    )
    Import-ADModule

    # Réinitialiser la liste des comptes créés pour le publipostage afin d’éviter
    # l’accumulation de données entre différentes importations. Sans cette
    # réinitialisation, les listes de publipostage s’additionnent à chaque appel du
    # script, comme vous l’avez constaté.
    $global:UtilisateursCreesImport = @()
    if (-not (Test-Path $CSVPath)) {
        Write-Host "Fichier CSV introuvable : $CSVPath" -ForegroundColor Red
        return
    }
    # Détermination du fichier de publipostage : demander à l'utilisateur une seule fois par session
    if (-not $global:PublipostageImportPath -or $global:PublipostageImportPath -eq '') {
        $pubPath = Read-Host -Prompt "Chemin du fichier de publipostage (laisser vide pour un nom automatique)"
        if ([string]::IsNullOrWhiteSpace($pubPath)) {
            $timestampImp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $pubPath = "publipostage_import_$timestampImp.csv"
        }
        $global:PublipostageImportPath = $pubPath
    }
    # Importation des données en utilisant l’encodage spécifié. Par défaut UTF8.
    # Spécifier un encodage différent (p. ex. Default ou ansi) peut être nécessaire
    # lorsque le fichier CSV a été enregistré avec un jeu de caractères différent.
    $records = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding $Encoding
    foreach ($record in $records) {
        $givenName = $record.GivenName
        $surname   = $record.Surname
        $codeClasse = $record.CodeClasse

        # Normaliser le prénom et le nom provenant du fichier CSV
        $givenName = Normalize-String -InputString $givenName
        $surname   = Normalize-String -InputString $surname
        if (-not $givenName -or -not $surname -or -not $codeClasse) {
            Write-Host "Enregistrement incomplet pour $($record | ConvertTo-Json -Compress)" -ForegroundColor Yellow
            continue
        }
        # Calculer ou utiliser SamAccountName
        $sam = if ($record.SamAccountName) { $record.SamAccountName } else { New-SamAccountName -GivenName $givenName -Surname $surname }
        # Assurer l’unicité : incrémente si le sam existe déjà
        $uniqueSam = $sam
        $counter = 1
        while (Get-ADUser -Filter { SamAccountName -eq $uniqueSam } -ErrorAction SilentlyContinue) {
            $uniqueSam = "${sam}$counter"
            $counter++
        }
        $sam = $uniqueSam
        # Calculer ou utiliser UPN
        $upn = if ($record.UserPrincipalName) { $record.UserPrincipalName } else { New-UserPrincipalName -SamAccountName $sam -Domain $Domaine }
        # Déterminer OU
        $ou = if ($record.Path) { $record.Path } else { Get-OU -CodeClasse $codeClasse }
        # Déterminer groupe
        $group = Get-GeneralGroup -CodeClasse $codeClasse
        # Mot de passe
        [System.Security.SecureString]$pwdSecure = $null
        if ($record.Password) {
            $pwdSecure = ConvertTo-SecureString -String $record.Password -AsPlainText -Force
        } else {
            $plain = New-RandomPassword
            Write-Host "Mot de passe généré pour $sam : $plain" -ForegroundColor Magenta
            $pwdSecure = ConvertTo-SecureString -String $plain -AsPlainText -Force
        }
        # Détermination du niveau pour la description
        if ($record.Description) {
            $description = $record.Description
        } else {
            $niveauCsv = Get-GeneralGroup -CodeClasse $codeClasse
            $description = if ($niveauCsv -eq 'Inconnu') { "$codeClasse $AnneeScolaire" } else { "$niveauCsv $AnneeScolaire" }
        }
        # Vérifier existence
        if (Get-ADUser -Filter { SamAccountName -eq $sam -or UserPrincipalName -eq $upn } -ErrorAction SilentlyContinue) {
            Write-Host "Utilisateur déjà existant : $sam ($upn) – création ignorée." -ForegroundColor Yellow
            continue
        }
        # Construire hashtable
        $userProps = @{
            GivenName             = $givenName
            Surname               = $surname
            Name                  = "$givenName $surname"
            DisplayName           = "$givenName $surname"
            SamAccountName        = $sam
            UserPrincipalName     = $upn
            Path                  = $ou
            AccountPassword       = $pwdSecure
            Enabled               = $true
            Description           = $description
        }
        if ($record.Department) { $userProps.Department = $record.Department }
        if ($record.Title)      { $userProps.Title      = $record.Title }
        New-SchoolADUser -UserData $userProps -WhatIf:$WhatIf
        # Groupes supplémentaires
        # Ajout au groupe général si pertinent
        if ($group -ne 'Inconnu') {
            try {
                if ($WhatIf) {
                    Add-ADGroupMember -Identity $group -Members $sam -WhatIf
                } else {
                    Add-ADGroupMember -Identity $group -Members $sam
                }
                Write-Host "Ajouté au groupe $group" -ForegroundColor Cyan
            } catch {
                Write-Host "Erreur d’ajout au groupe $group : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        # Groupes listés dans la colonne Groups (séparés par des virgules)
        if ($record.Groups) {
            $grpNames = $record.Groups -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($grp in $grpNames) {
                try {
                    if ($WhatIf) {
                        Add-ADGroupMember -Identity $grp -Members $sam -WhatIf
                    } else {
                        Add-ADGroupMember -Identity $grp -Members $sam
                    }
                    Write-Host "Ajouté au groupe supplémentaire : $grp" -ForegroundColor Cyan
                } catch {
                    Write-Host "Erreur d’ajout au groupe $grp : $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        # -- Publipostage pour l'import CSV --
        # Ajouter le compte créé à la liste pour le publipostage en convertissant le mot de passe
        $plainPwdImp = ConvertTo-PlainText -SecureString $pwdSecure
        $utilObjImp = [PSCustomObject]@{
            NomPrenom      = "$givenName $surname"
            NomUtilisateur = $sam
            MotDePasse     = $plainPwdImp
            Classe         = $codeClasse
        }
        $global:UtilisateursCreesImport += $utilObjImp
        Write-Host "$givenName $surname a été ajouté(e) à la liste des utilisateurs pour le publipostage" -ForegroundColor Cyan
    }

    # Exporter les comptes créés via l'import CSV vers le fichier de publipostage
    if ($global:UtilisateursCreesImport.Count -gt 0) {
        try {
            $global:UtilisateursCreesImport | Export-Csv -Path $global:PublipostageImportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            Write-Host "Fichier de publipostage import sauvegardé : $($global:PublipostageImportPath)" -ForegroundColor Cyan
        } catch {
            Write-Host "Erreur lors de l’export du fichier de publipostage import : $($_.Exception.Message)" -ForegroundColor Yellow
        }
        # Vider la liste de publipostage après export afin de ne pas cumuler plusieurs listes lors des appels suivants.
        $global:UtilisateursCreesImport = @()
    }
}

function Add-StaffFromCsv {
    <#
        .SYNOPSIS
            Crée des comptes personnels (enseignants, personnels, AESH, surveillants, etc.) à partir d’un CSV.
            Ici c'est surtout pour avoir plus de détails sur la saisie des fiches adultes mais sera plus tard merge avec l'ajout élève.

        .DESCRIPTION
            Cette fonction lit un fichier CSV dont chaque ligne décrit un adulte à
            créer dans Active Directory. Les colonnes attendues sont :
            - GivenName (prénom)
            - Surname (nom)
            - Role (catégorie : Extérieur, Personnel, Prof, Surveillant, AESH)
            - SamAccountName (optionnel)
            - Password (optionnel)
            - Department, Title (optionnel)
            Les champs SamAccountName et Password peuvent être laissés vides pour que
            le script les génère automatiquement. Le rôle détermine l’OU et le
            groupe général.

        .PARAMETER CSVPath
            Chemin complet vers le fichier CSV. Le point‑virgule est utilisé comme séparateur.
        .PARAMETER WhatIf
            Exécute la création en mode simulation.
        .PARAMETER Encoding
            Encodage du fichier CSV (UTF8 par défaut). Utilisez utf8BOM, Default ou ansi
            si vos fichiers proviennent d’Excel ou d’une autre source et que les
            caractères accentués ne s’affichent pas correctement.
    #>
    param(
        [string]$CSVPath,
        [switch]$WhatIf,
        [string]$Encoding = 'UTF8'
    )
    Import-ADModule
    if (-not (Test-Path $CSVPath)) {
        Write-Host "Fichier CSV introuvable : $CSVPath" -ForegroundColor Red
        return
    }
    # Réinitialiser la liste pour ne pas accumuler les entrées
    $global:StaffCreesImport = @()
    # Déterminer le fichier de publipostage : demander à l’utilisateur une seule fois
    if (-not $global:PublipostageStaffImportPath -or $global:PublipostageStaffImportPath -eq '') {
        $pubPath = Read-Host -Prompt "Chemin du fichier de publipostage personnel (laisser vide pour un nom automatique)"
        if ([string]::IsNullOrWhiteSpace($pubPath)) {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $pubPath = "publipostage_staff_import_$timestamp.csv"
        }
        $global:PublipostageStaffImportPath = $pubPath
    }
    # Importer les données avec l’encodage spécifié
    $records = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding $Encoding
    foreach ($record in $records) {
        $givenName = $record.GivenName
        $surname   = $record.Surname
        $role      = $null
        if ($record.PSObject.Properties['Role']) { $role = $record.Role }
        elseif ($record.PSObject.Properties['Category']) { $role = $record.Category }
        elseif ($record.PSObject.Properties['Statut']) { $role = $record.Statut }
        if (-not $givenName -or -not $surname -or -not $role) {
            Write-Host "Enregistrement incomplet : $($record | ConvertTo-Json -Compress)" -ForegroundColor Yellow
            continue
        }
        # Normaliser les noms pour la génération des identifiants
        $normGiven = Normalize-String -InputString $givenName
        $normSurname = Normalize-String -InputString $surname
        # Utiliser l’identifiant fourni ou en créer un
        $sam = if ($record.SamAccountName) { $record.SamAccountName } else { New-SamAccountName -GivenName $normGiven -Surname $normSurname }
        # Assurer l’unicité du samAccountName
        $uniqueSam = $sam
        $cpt = 1
        while (Get-ADUser -Filter { SamAccountName -eq $uniqueSam } -ErrorAction SilentlyContinue) {
            $uniqueSam = "${sam}$cpt"
            $cpt++
        }
        $sam = $uniqueSam
        # UPN
        $upn = New-UserPrincipalName -SamAccountName $sam -Domain $Domaine
        # OU et groupe
        $ou = Get-StaffOU -Role $role
        $group = Get-StaffGroup -Role $role
        # Mot de passe
        [System.Security.SecureString]$pwdSecure = $null
        if ($record.Password) {
            $pwdSecure = ConvertTo-SecureString -String $record.Password -AsPlainText -Force
        } else {
            $plain = New-RandomPassword
            Write-Host "Mot de passe généré pour $sam : $plain" -ForegroundColor Magenta
            $pwdSecure = ConvertTo-SecureString -String $plain -AsPlainText -Force
        }
        # Description : rôle et année scolaire pour identifier rapidement
        $description = "$role $AnneeScolaire"
        # Vérifier l’existence
        if (Get-ADUser -Filter { SamAccountName -eq $sam -or UserPrincipalName -eq $upn } -ErrorAction SilentlyContinue) {
            Write-Host "Compte existant : $sam ($upn) – création ignorée." -ForegroundColor Yellow
            continue
        }
        # Préparer les propriétés
        $userProps = @{
            GivenName             = $givenName
            Surname               = $surname
            Name                  = "$givenName $surname"
            DisplayName           = "$givenName $surname"
            SamAccountName        = $sam
            UserPrincipalName     = $upn
            Path                  = $ou
            AccountPassword       = $pwdSecure
            Enabled               = $true
            Description           = $description
        }
        if ($record.Department) { $userProps.Department = $record.Department }
        if ($record.Title)      { $userProps.Title      = $record.Title }
        # Création du compte
        New-SchoolADUser -UserData $userProps -WhatIf:$WhatIf
        # Ajout au groupe associé
        if ($group) {
            try {
                if ($WhatIf) {
                    Add-ADGroupMember -Identity $group -Members $sam -WhatIf
                } else {
                    Add-ADGroupMember -Identity $group -Members $sam
                }
                Write-Host "Ajouté au groupe $group" -ForegroundColor Cyan
            } catch {
                Write-Host "Erreur d’ajout au groupe $group : $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        # Préparation de l'objet pour le publipostage
        $plainPwd = ConvertTo-PlainText -SecureString $pwdSecure
        $obj = [PSCustomObject]@{
            NomPrenom      = "$givenName $surname"
            NomUtilisateur = $sam
            MotDePasse     = $plainPwd
            Role           = $role
        }
        $global:StaffCreesImport += $obj
        Write-Host "$givenName $surname a été ajouté(e) à la liste du personnel pour le publipostage" -ForegroundColor Cyan
    }
    # Exporter la liste si nécessaire
    if ($global:StaffCreesImport.Count -gt 0) {
        try {
            $global:StaffCreesImport | Export-Csv -Path $global:PublipostageStaffImportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            Write-Host "Fichier de publipostage personnel import sauvegardé : $($global:PublipostageStaffImportPath)" -ForegroundColor Cyan
        } catch {
            Write-Host "Erreur lors de l’export du fichier de publipostage personnel : $($_.Exception.Message)" -ForegroundColor Yellow
        }
        # Vider la liste après export
        $global:StaffCreesImport = @()
    }
}

function Add-StaffInteractive {
    <#
        .SYNOPSIS
            Crée un compte personnel (enseignant, personnel, AESH, surveillant…) en
            saisie interactive.

        .DESCRIPTION
            Demande à l’utilisateur de saisir le prénom, le nom et le rôle du
            personnel. Génère un SamAccountName et un UPN si nécessaire, permet
            de saisir un mot de passe ou d’en générer un et crée le compte dans
            l’OU appropriée en l’ajoutant au groupe correspondant. Un fichier de
            publipostage est ensuite exporté pour conserver la liste des comptes
            créés (nom complet, identifiant, mot de passe et rôle).

        .PARAMETER WhatIf
            Permet de simuler la création des comptes sans les effectuer.
    #>
    param([switch]$WhatIf)
    Import-ADModule
    # Réinitialiser la liste pour cet ajout
    $global:StaffCreesInteractive = @()
    Write-Host "Ajout interactif d’un membre du personnel" -ForegroundColor Cyan
    $givenName = Read-Host -Prompt "Prénom"
    $surname   = Read-Host -Prompt "Nom"
    $roleInput = Read-Host -Prompt "Rôle (Extérieur/Personnel/Prof/Surveillant/AESH)"
    if (-not $givenName -or -not $surname -or -not $roleInput) {
        Write-Host "Saisie incomplète." -ForegroundColor Yellow
        return
    }
    # Normaliser pour les identifiants
    $normGiven = Normalize-String -InputString $givenName
    $normSurname = Normalize-String -InputString $surname
    $sam = New-SamAccountName -GivenName $normGiven -Surname $normSurname
    # Assurer unicité
    $uniqueSam = $sam
    $cpt = 1
    while (Get-ADUser -Filter { SamAccountName -eq $uniqueSam } -ErrorAction SilentlyContinue) {
        $uniqueSam = "${sam}$cpt"
        $cpt++
    }
    $sam = $uniqueSam
    $upn = New-UserPrincipalName -SamAccountName $sam -Domain $Domaine
    $ou = Get-StaffOU -Role $roleInput
    $group = Get-StaffGroup -Role $roleInput
    # Demander le mot de passe ou générer
    $pwChoice = Read-Host -Prompt "Voulez‑vous saisir le mot de passe ? (O/N)"
    [System.Security.SecureString]$passwordSecure = $null
    if ($pwChoice -match '^[oOyY]') {
        $passwordSecure = Read-Host -Prompt 'Mot de passe' -AsSecureString
    } else {
        $generated = New-RandomPassword
        Write-Host "Mot de passe généré : $generated" -ForegroundColor Magenta
        $passwordSecure = ConvertTo-SecureString -String $generated -AsPlainText -Force
    }
    # Description
    $description = "$roleInput $AnneeScolaire"
    # Vérifier l’existence
    if (Get-ADUser -Filter { SamAccountName -eq $sam -or UserPrincipalName -eq $upn } -ErrorAction SilentlyContinue) {
        Write-Host "Un utilisateur avec l’identifiant $sam ou l’UPN $upn existe déjà." -ForegroundColor Yellow
        return
    }
    # Préparer les propriétés
    $userProps = @{
        GivenName             = $givenName
        Surname               = $surname
        Name                  = "$givenName $surname"
        DisplayName           = "$givenName $surname"
        SamAccountName        = $sam
        UserPrincipalName     = $upn
        Path                  = $ou
        AccountPassword       = $passwordSecure
        Enabled               = $true
        Description           = $description
    }
    # Créer le compte
    New-SchoolADUser -UserData $userProps -WhatIf:$WhatIf
    # Ajouter au groupe
    if ($group) {
        try {
            if ($WhatIf) {
                Add-ADGroupMember -Identity $group -Members $sam -WhatIf
            } else {
                Add-ADGroupMember -Identity $group -Members $sam
            }
            Write-Host "Ajouté au groupe $group" -ForegroundColor Cyan
        } catch {
            Write-Host "Erreur d’ajout au groupe $group : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    # Gestion du publipostage pour le personnel (fichier) : demander un chemin une fois
    if (-not $global:PublipostageStaffInteractivePath -or $global:PublipostageStaffInteractivePath -eq '') {
        $pubPath = Read-Host -Prompt "Chemin du fichier de publipostage personnel (laisser vide pour un nom automatique)"
        if ([string]::IsNullOrWhiteSpace($pubPath)) {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $pubPath = "publipostage_staff_interactif_$timestamp.csv"
        }
        $global:PublipostageStaffInteractivePath = $pubPath
    }
    # Préparation de l’objet à exporter
    $plainPwd = ConvertTo-PlainText -SecureString $passwordSecure
    $obj = [PSCustomObject]@{
        NomPrenom      = "$givenName $surname"
        NomUtilisateur = $sam
        MotDePasse     = $plainPwd
        Role           = $roleInput
    }
    $global:StaffCreesInteractive += $obj
    Write-Host "$givenName $surname a été ajouté(e) à la liste du personnel pour le publipostage" -ForegroundColor Cyan
    Write-Host "SamAccountName : $sam" -ForegroundColor Cyan
    # Exporter la liste
    try {
        $global:StaffCreesInteractive | Export-Csv -Path $global:PublipostageStaffInteractivePath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    } catch {
        Write-Host "Impossible d’exporter le fichier de publipostage personnel interactif : $($_.Exception.Message)" -ForegroundColor Yellow
    }
    # Vider la liste après export
    $global:StaffCreesInteractive = @()
}
function Remove-StudentsExactFromCsv {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CSVPath,
        [string]$ReportFilePath = '',
        [switch]$WhatIf,
        # Encodage du fichier CSV (UTF8 par défaut). Ajustez selon l’encodage de votre fichier【582535038188393†L39-L75】.
        [string]$Encoding = 'UTF8'
    )
    Import-ADModule
    if (-not (Test-Path $CSVPath)) {
        Write-Host "Fichier CSV introuvable : $CSVPath" -ForegroundColor Red
        return
    }
    # Importer en utilisant l’encodage spécifié
    $records = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding $Encoding
    if (-not $records) {
        Write-Host "Aucune donnée dans le fichier CSV." -ForegroundColor Yellow
        return
    }
    $toDelete = @()
    $notFound = @()
    foreach ($r in $records) {
        $given = $null
        $sur   = $null
        if ($r.PSObject.Properties['GivenName']) { $given = $r.GivenName }
        elseif ($r.PSObject.Properties['Prenom']) { $given = $r.Prenom }
        if ($r.PSObject.Properties['Surname']) { $sur = $r.Surname }
        elseif ($r.PSObject.Properties['Nom']) { $sur = $r.Nom }
        if (-not $given -or -not $sur) { continue }
        $given = $given.Trim()
        $sur   = $sur.Trim()
        $given = Normalize-String -InputString $given
        $sur   = Normalize-String -InputString $sur
        $filter = "(GivenName -eq '$given' -and Surname -eq '$sur')"
        $found = @(Get-ADUser -Filter $filter -SearchBase $ElevesOU -SearchScope Subtree -Properties SamAccountName, GivenName, Surname)
        if ($found.Count -gt 0) {
            $toDelete += $found
        } else {
            $notFound += "$given $sur"
        }
    }
    Write-Host "Comptes trouvés : $($toDelete.Count)" -ForegroundColor Cyan
    foreach ($u in $toDelete) {
        Write-Host "$($u.GivenName) $($u.Surname) - $($u.SamAccountName)" -ForegroundColor Yellow
    }
    if ($notFound.Count -gt 0) {
        Write-Host "Aucun compte trouvé pour :" -ForegroundColor DarkGray
        $notFound | ForEach-Object { Write-Host "  $_" }
    }
    if ($toDelete.Count -eq 0) {
        Write-Host "Aucun compte à supprimer." -ForegroundColor Yellow
        return
    }
    if (-not $WhatIf) {
        $confirm = Read-Host -Prompt "Confirmez-vous la suppression de ces comptes ? (O/N)"
        if ($confirm -notmatch '^[oOyY]') {
            Write-Host "Suppression annulée." -ForegroundColor Yellow
            return
        }
    }
    $report = @()
    foreach ($user in $toDelete) {
        $entry = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            NomPrenom      = "$($user.GivenName) $($user.Surname)"
            Status         = ''
        }
        try {
            if ($WhatIf) {
                $entry.Status = 'Simulé'
            } else {
                Remove-ADUser -Identity $user -Confirm:$false -ErrorAction Stop
                $entry.Status = 'Supprimé'
            }
        } catch {
            $entry.Status = "Erreur : $($_.Exception.Message)"
        }
        $report += $entry
    }
    if ([string]::IsNullOrWhiteSpace($ReportFilePath)) {
        $timestampEx = (Get-Date -Format 'yyyyMMdd_HHmmss')
        $ReportFilePath = "suppression_eleves_exact_$timestampEx.csv"
    }
    try {
        $report | Export-Csv -Path $ReportFilePath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        Write-Host "Rapport sauvegardé dans : $ReportFilePath" -ForegroundColor Cyan
    } catch {
        Write-Host "Impossible d’enregistrer le rapport : $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Show-Menu {
    <#
        .SYNOPSIS
            Affiche le menu principal et gère la sélection de l’utilisateur. Le fameux menu... Oui il est à la fin du script... Et pourquoi pas ?
    #>
    do {
        Write-Host "============================================="
        Write-Host "  Gestion des comptes élèves et personnels - AD $AnneeScolaire" -ForegroundColor Cyan
        Write-Host "============================================="
        Write-Host "1. Ajouter un élève (mode interactif)"
        Write-Host "2. Importer des élèves depuis un fichier CSV"
        Write-Host "3. Promouvoir des élèves (passage de classe)"
        Write-Host "4. Réorganiser les groupes d’élèves"
        Write-Host "5. Informations détaillées sur un élève"
        Write-Host "6. Changer le mot de passe d’un élève"
        Write-Host "7. Rechercher un élève"
        Write-Host "8. Retour en arrière d’une promotion"
        Write-Host "9. Supprimer des élèves sortants (CSV)"
        Write-Host "10. Supprimer un élève par SamAccountName"
        Write-Host "11. Supprimer une liste d'élèves (Nom;Prénom)"
        Write-Host "12. Ajouter un membre du personnel (mode interactif)"
        Write-Host "13. Importer du personnel depuis un fichier CSV"
        Write-Host "0. Quitter"
        $choice = Read-Host -Prompt "Votre choix"
        switch ($choice) {
            '1' { Add-UserInteractive }
            '2' {
                $csvPath = Read-Host -Prompt "Chemin complet du fichier CSV (délimité par ;)"
                # Demander l'encodage utilisé pour le fichier. Laisser vide pour utiliser UTF8 par défaut.
                $enc = Read-Host -Prompt "Encodage du fichier (laisser vide pour UTF8, ex : utf8BOM, Default, ansi)"
                if ([string]::IsNullOrWhiteSpace($enc)) {
                    Add-UsersFromCsv -CSVPath $csvPath
                } else {
                    Add-UsersFromCsv -CSVPath $csvPath -Encoding $enc
                }
            }
            '3' { Promote-StudentsMenu }
            '4' { Reorganize-Groups }
            '5' { Get-UserInfo }
            '6' { Change-UserPassword }
            '7' { Search-User }
            '8' {
                $reportFile = Read-Host -Prompt "Chemin complet du fichier de rapport à annuler (CSV)"
                if (-not [string]::IsNullOrWhiteSpace($reportFile) -and (Test-Path $reportFile)) {
                    Revert-Promotion -ReportPath $reportFile
                } else {
                    Write-Host "Chemin de rapport invalide ou fichier introuvable." -ForegroundColor Yellow
                }
            }
            '9' {
                $csvPath = Read-Host -Prompt "Chemin complet du fichier CSV des élèves sortants"
                $reportPath = Read-Host -Prompt "Chemin du rapport de suppression (laisser vide pour un nom automatique)"
                # Demander l'encodage du fichier pour éviter des problèmes de caractères spéciaux
                $encDel = Read-Host -Prompt "Encodage du fichier (laisser vide pour UTF8, ex : utf8BOM, Default, ansi)"
                if ([string]::IsNullOrWhiteSpace($reportPath)) {
                    if ([string]::IsNullOrWhiteSpace($encDel)) {
                        Remove-StudentsFromCsv -CSVPath $csvPath
                    } else {
                        Remove-StudentsFromCsv -CSVPath $csvPath -Encoding $encDel
                    }
                } else {
                    if ([string]::IsNullOrWhiteSpace($encDel)) {
                        Remove-StudentsFromCsv -CSVPath $csvPath -ReportFilePath $reportPath
                    } else {
                        Remove-StudentsFromCsv -CSVPath $csvPath -ReportFilePath $reportPath -Encoding $encDel
                    }
                }
            }
            '10' {
             $report = Read-Host -Prompt "Chemin du rapport (laisser vide pour automatique)"
             if ([string]::IsNullOrWhiteSpace($report)) {
             Remove-StudentBySam
               } else {
                   Remove-StudentBySam -ReportFilePath $report
               }
             }
            '11' {
                $csvPathDel = Read-Host -Prompt "Chemin complet du fichier CSV des élèves à supprimer (Nom;Prénom)"
                $reportPathDel = Read-Host -Prompt "Chemin du rapport de suppression (laisser vide pour un nom automatique)"
                # Encodage du fichier pour l’import des noms/prénoms
                $encDel2 = Read-Host -Prompt "Encodage du fichier (laisser vide pour UTF8, ex : utf8BOM, Default, ansi)"
                if ([string]::IsNullOrWhiteSpace($reportPathDel)) {
                    if ([string]::IsNullOrWhiteSpace($encDel2)) {
                        Remove-StudentsExactFromCsv -CSVPath $csvPathDel
                    } else {
                        Remove-StudentsExactFromCsv -CSVPath $csvPathDel -Encoding $encDel2
                    }
                } else {
                    if ([string]::IsNullOrWhiteSpace($encDel2)) {
                        Remove-StudentsExactFromCsv -CSVPath $csvPathDel -ReportFilePath $reportPathDel
                    } else {
                        Remove-StudentsExactFromCsv -CSVPath $csvPathDel -ReportFilePath $reportPathDel -Encoding $encDel2
                    }
                }
            }
            '0' { Write-Host "Fin du programme."; return }
            '12' {
                # Ajout interactif d'un personnel (profs, AESH, surveillants, etc.)
                Add-StaffInteractive
            }
            '13' {
                # Importer du personnel depuis un fichier CSV
                $csvPathStaff = Read-Host -Prompt "Chemin complet du fichier CSV du personnel (délimité par ;)"
                # Demander l'encodage du fichier. Par défaut UTF8.
                $encStaff = Read-Host -Prompt "Encodage du fichier (laisser vide pour UTF8, ex : utf8BOM, Default, ansi)"
                if ([string]::IsNullOrWhiteSpace($encStaff)) {
                    Add-StaffFromCsv -CSVPath $csvPathStaff
                } else {
                    Add-StaffFromCsv -CSVPath $csvPathStaff -Encoding $encStaff
                }
            }
            default { Write-Host "Choix invalide. Veuillez recommencer." -ForegroundColor Yellow }
        }
    } while ($true)
}

<#
    Promotion des élèves vers la classe supérieure
    -------------------------------------------
    Cette fonction interactive propose différentes transitions (par exemple
    6EME→5EME, 5EME→4EME…) et appelle la fonction `Promote-Students` pour
    effectuer la mise à jour.
    Logique de promotion à adapter selon l'environnement mais généralement on passe bien de la 6EME à la 5EME en France ;)
    Possibilité également d'en ajouter, sert ici d'exemple ceux qui correspondent à mon cas.
#>
function Promote-StudentsMenu {
    $options = @{
        '1' = @{Current='6EME'; Next='5EME'}
        '2' = @{Current='5EME'; Next='4EME'}
        '3' = @{Current='4EME'; Next='3EME'}
        '4' = @{Current='3EME'; Next='2NDE'}
        '5' = @{Current='2NDE'; Next='1ERE'}
        '6' = @{Current='1ERE'; Next='TERMINALE'}
        '7' = @{Current='BTS1M'; Next='BTS2M'}
        '8' = @{Current='BTS1G'; Next='BTS2G'}
        '9' = @{Current='BTS1CCST'; Next='BTS2CCST'}
    }
    Write-Host "Promouvoir les élèves :"
    Write-Host "1. 6EME → 5EME"
    Write-Host "2. 5EME → 4EME"
    Write-Host "3. 4EME → 3EME"
    Write-Host "4. 3EME → 2NDE"
    Write-Host "5. 2NDE → 1ERE"
    Write-Host "6. 1ERE → TERMINALE"
    Write-Host "7. BTS1M → BTS2M"
    Write-Host "8. BTS1G → BTS2G"
    Write-Host "9. BTS1CCST → BTS2CCST"
    Write-Host "0. Retour"
    $choice = Read-Host -Prompt "Sélectionnez une option"
    if ($choice -eq '0') { return }
    if ($options.ContainsKey($choice)) {
        $curr = $options[$choice].Current
        $next = $options[$choice].Next
        # Demander un chemin de rapport ou laisser vide pour un nom automatique
        $reportPath = Read-Host -Prompt "Chemin complet du rapport (laisser vide pour un fichier automatique dans le dossier courant)"
        if ([string]::IsNullOrWhiteSpace($reportPath)) {
            Promote-Students -CurrentClass $curr -NewClass $next
        } else {
            Promote-Students -CurrentClass $curr -NewClass $next -ReportFilePath $reportPath
        }
    } else {
        Write-Host "Option invalide." -ForegroundColor Yellow
    }
}

<#
    Promote-Students
    ----------------
    Réalise la promotion des élèves d’une classe vers la classe supérieure. Les
    comptes sont recherchés dans l’OU définie par la variable $ElevesOU. Le
    code classe est extrait du premier mot du champ Description. Les élèves
    correspondants sont transférés dans le nouveau groupe, et leur description
    est mise à jour avec la nouvelle classe et l’année scolaire.
#>
function Promote-Students {
    param(
        [string]$CurrentClass,
        [string]$NewClass,
        [string]$ReportFilePath = ''
    )
    Import-ADModule
    Write-Host "Promotion des élèves de $CurrentClass vers $NewClass" -ForegroundColor Cyan
    # Préparation du rapport de modifications
    $changes = @()
    # Recherche de tous les élèves dans l'OU des élèves et leurs sous‑OU
    # La description peut contenir un code classe complet (par exemple 4B 2024-2025).
    # On extrait le premier élément de la description puis on détermine le niveau général
    # à l'aide de la fonction Get-GeneralGroup. Seuls les élèves dont le niveau
    # correspond à la classe actuelle ($CurrentClass) sont promus.
    $eleves = Get-ADUser -Filter * -SearchBase $ElevesOU -SearchScope Subtree -Properties Description
    if (-not $eleves) {
        Write-Host "Aucun élève trouvé dans $ElevesOU" -ForegroundColor Yellow
        return
    }
    # Vérifie l’existence du groupe de destination
    $targetGroup = Get-ADGroup -Identity $NewClass -ErrorAction SilentlyContinue
    if (-not $targetGroup) {
        Write-Host "Le groupe $NewClass n’existe pas dans Active Directory." -ForegroundColor Yellow
    }
    # Calculer la prochaine année scolaire pour mettre à jour les descriptions
    $nextYear = Get-NextSchoolYear -CurrentYear $AnneeScolaire
    foreach ($eleve in $eleves) {
        $desc = $eleve.Description
        if (-not $desc) { continue }
        # Ne promouvoir que les comptes dont la description contient l'année scolaire actuelle
        if ($desc -notlike "*$AnneeScolaire*") { continue }
        # Extraire le code classe (premier mot) et obtenir le niveau général
        $code = ($desc -split ' ')[0]
        $niveau = Get-GeneralGroup -CodeClasse $code
        # Si le niveau correspond à la classe actuelle (insensible à la casse), on promulgue
        if ($null -eq $niveau) { continue }
        if ($niveau.ToUpper() -ne $CurrentClass.ToUpper()) { continue }
        # Construire la nouvelle description en utilisant la prochaine année scolaire
        $nouvelleDescription = "$NewClass $nextYear"
        try {
            # Mettre à jour la description et conserver l’ancienne pour le rapport
            $ancienneDescription = $eleve.Description
            Set-ADUser -Identity $eleve -Description $nouvelleDescription -ErrorAction Stop
            # Retirer l’élève de l’ancien groupe (niveau actuel) si celui‑ci existe
            try {
                $grpExist = Get-ADGroup -Identity $niveau -ErrorAction SilentlyContinue
                if ($grpExist) {
                    Remove-ADGroupMember -Identity $niveau -Members $eleve.SamAccountName -Confirm:$false -ErrorAction SilentlyContinue
                }
            } catch {}
            # Déplacer l’élève vers la nouvelle OU si elle diffère (ex : passage collège → lycée) = oui je suis fénéant à ce point-là.
            try {
                $ancienneOU = Get-OU -CodeClasse $niveau
                $nouvelleOU = Get-OU -CodeClasse $NewClass
                if ($ancienneOU -and $nouvelleOU -and ($ancienneOU -ne $nouvelleOU)) {
                    # Utiliser l’objet utilisateur comme identité pour Move-ADObject. Cela déplace le compte vers l’OU ciblée.
                    Move-ADObject -Identity $eleve -TargetPath $nouvelleOU -Confirm:$false -ErrorAction SilentlyContinue
                }
            } catch {
                Write-Host "Erreur lors du déplacement de $($eleve.SamAccountName) vers $nouvelleOU : $($_.Exception.Message)" -ForegroundColor Yellow
            }
            # Ajouter l’élève au nouveau groupe s’il existe
            if ($targetGroup) {
                Add-ADGroupMember -Identity $NewClass -Members $eleve.SamAccountName -ErrorAction SilentlyContinue
            }
            # Ajouter une entrée au rapport
            $changes += [PSCustomObject]@{
                SamAccountName    = $eleve.SamAccountName
                AncienneDescription = $ancienneDescription
                NouvelleDescription = $nouvelleDescription
            }
            Write-Host "Élève $($eleve.SamAccountName) promu en $NewClass" -ForegroundColor Green
        } catch {
            Write-Host "Erreur lors de la promotion de $($eleve.SamAccountName) : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    # Enregistrer le rapport si des modifications ont été apportées
    if ($changes.Count -gt 0) {
        # Déterminer le chemin du fichier rapport si non fourni
        if ([string]::IsNullOrWhiteSpace($ReportFilePath)) {
            $timestamp = (Get-Date -Format 'yyyyMMdd_HHmmss')
            $ReportFilePath = "promotion_${CurrentClass}_vers_${NewClass}_$timestamp.csv"
        }
        try {
            $changes | Export-Csv -Path $ReportFilePath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            Write-Host "Rapport sauvegardé dans : $ReportFilePath" -ForegroundColor Cyan
        } catch {
            Write-Host "Impossible de sauvegarder le rapport : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

<#
    Réorganise les groupes d’élèves en fonction de leur code classe.
    Vide les groupes passés en paramètre puis réaffecte les élèves en
    inspectant le premier mot de leur Description (code classe).
#>
function Reorganize-Groups {
    [CmdletBinding()]
    param()
    Import-ADModule
    # Liste des groupes à gérer
    $groups = @('6EME','5EME','4EME','3EME','2NDE','1ERE','TERMINALE','BTS1M','BTS2M','BTS1G','BTS2G','BTS1CCST','BTS2CCST') #<-# Liste des groupes à modifier, passage en variable global dans les prochaines versions pour éviter la répétition.
    Write-Host "Réorganisation des groupes…" -ForegroundColor Cyan
    # Étape 1 : vider les groupes
    foreach ($grp in $groups) {
        try {
            $members = Get-ADGroupMember -Identity $grp -ErrorAction SilentlyContinue
            if ($members) {
                Remove-ADGroupMember -Identity $grp -Members $members -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "Groupe $grp vidé." -ForegroundColor Yellow
            } else {
                Write-Host "Groupe $grp ne contenait aucun membre." -ForegroundColor DarkGray
            }
        } catch {
            Write-Host "Erreur lors du vidage du groupe $grp : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    # Étape 2 : réaffecter les élèves
    $eleves = Get-ADUser -Filter * -SearchBase $ElevesOU -SearchScope Subtree -Properties Description
    foreach ($eleve in $eleves) {
        $desc = $eleve.Description
        if (-not $desc) { continue }
        # Extraire le premier mot de la description pour obtenir le code classe
        $code = ($desc -split ' ')[0]
        # Déterminer le niveau général (6EME, 5EME, etc.)
        $niveau = Get-GeneralGroup -CodeClasse $code
        if ($null -ne $niveau -and ($groups -contains $niveau.ToUpper())) {
            try {
                Add-ADGroupMember -Identity $niveau -Members $eleve.SamAccountName -ErrorAction Stop
                Write-Host "Élève $($eleve.SamAccountName) ajouté au groupe $niveau." -ForegroundColor Green
            } catch {
                Write-Host "Erreur lors de l’ajout de $($eleve.SamAccountName) au groupe $niveau : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

<#
    Affiche les informations détaillées d’un élève.
    Saisie du nom de famille, affichage des utilisateurs correspondant à la
    recherche. En cas de plusieurs résultats, l’utilisateur sélectionne
    l’élève voulu pour voir son SamAccountName et sa classe (Description).
#>
function Get-UserInfo {
    Import-ADModule
    $surname = Read-Host -Prompt "Nom de famille à rechercher"
    if (-not $surname) { return }
    # Construire une chaîne de filtre pour permettre l’injection de la variable dans la requête
    $filter = "Surname -like '$surname*'"
    $results = Get-ADUser -Filter $filter -SearchBase $ElevesOU -SearchScope Subtree -Properties GivenName, Surname, SamAccountName, Description
    # Convertir en tableau pour éviter les erreurs lorsque le résultat n’est pas un array
    $resultsArray = @($results)
    if (-not $resultsArray -or $resultsArray.Count -eq 0) {
        Write-Host "Aucun utilisateur trouvé pour $surname" -ForegroundColor Yellow
        return
    }
    if ($resultsArray.Count -eq 1) {
        $user = $resultsArray[0]
        Write-Host "Nom : $($user.GivenName) $($user.Surname)" -ForegroundColor Cyan
        Write-Host "Nom d’utilisateur : $($user.SamAccountName)"
        Write-Host "Description/Classe : $($user.Description)"
        return
    }
    # Plusieurs résultats – afficher un menu
    Write-Host "Résultats trouvés :"
    $i = 1
    foreach ($u in $resultsArray) {
        Write-Host "[$i] $($u.GivenName) $($u.Surname) - $($u.Description)"
        $i++
    }
    $sel = Read-Host -Prompt "Numéro de l’élève à afficher"
    if ($sel -match '^[0-9]+$') {
        $index = [int]$sel - 1
        if ($index -ge 0 -and $index -lt $resultsArray.Count) {
            $user = $resultsArray[$index]
            Write-Host "Nom : $($user.GivenName) $($user.Surname)" -ForegroundColor Cyan
            Write-Host "Nom d’utilisateur : $($user.SamAccountName)"
            Write-Host "Description/Classe : $($user.Description)"
        }
    }
}

<#
    Permet de changer rapidement le mot de passe d’un élève.
    Saisissez le SamAccountName, puis saisissez le nouveau mot de passe. Le
    mot de passe est transmis en SecureString via Read‑Host ; ensuite la
    cmdlet `Set-ADAccountPassword` est utilisée pour réinitialiser le
    mot de passe.
#>
function Change-UserPassword {
    Import-ADModule
    $sam = Read-Host -Prompt "Nom d’utilisateur (SamAccountName)"
    if (-not $sam) { return }
    $user = Get-ADUser -Identity $sam -ErrorAction SilentlyContinue
    if (-not $user) {
        Write-Host "Utilisateur introuvable : $sam" -ForegroundColor Yellow
        return
    }
    $securePwd = Read-Host -Prompt "Nouveau mot de passe" -AsSecureString
    try {
        Set-ADAccountPassword -Identity $sam -NewPassword $securePwd -Reset -ErrorAction Stop
        Write-Host "Mot de passe modifié pour $sam" -ForegroundColor Green
    } catch {
        Write-Host "Erreur lors de la modification du mot de passe : $($_.Exception.Message)" -ForegroundColor Red
    }
}

<#
    Recherche un élève par nom de famille et affiche les résultats sous forme
    de liste (SamAccountName et Description). Cette fonction est utile pour
    retrouver rapidement le nom d’utilisateur sans entrer dans le détail.
#>
function Search-User {
    Import-ADModule
    $surname = Read-Host -Prompt "Nom de famille à rechercher"
    if (-not $surname) { return }
    # Construire une chaîne de filtre pour inclure la variable dans la requête AD
    $filter = "Surname -like '$surname*'"
    $results = Get-ADUser -Filter $filter -SearchBase $ElevesOU -SearchScope Subtree -Properties SamAccountName, Description
    $resultsArray = @($results)
    if (-not $resultsArray -or $resultsArray.Count -eq 0) {
        Write-Host "Aucun utilisateur trouvé pour $surname" -ForegroundColor Yellow
        return
    }
    Write-Host "Utilisateurs trouvés :" -ForegroundColor Cyan
    foreach ($u in $resultsArray) {
        Write-Host "$($u.SamAccountName) - $($u.Description)"
    }
}

# -----------------------------------------------------------------------------
# Retour en arrière d’une promotion
# -----------------------------------------------------------------------------
<#
    Revert-Promotion
    ----------------
    Cette fonction permet d’annuler une promotion en se basant sur un
    fichier de rapport généré par la fonction `Promote-Students`. Chaque
    enregistrement du rapport contient le `SamAccountName`, l’ancienne
    description et la nouvelle description. La fonction va :

      • restaurer la description à sa valeur d’origine ;
      • déterminer le niveau correspondant à l’ancienne et la nouvelle
        description pour ajuster l’appartenance aux groupes ;
      • déplacer l’utilisateur dans l’OU appropriée si celle-ci diffère de
        l’OU actuelle (par exemple retour du lycée vers le collège).

        Perso j'appelle cette fonction le "Oups, j'ai fait une bêtise"

    .PARAMETER ReportPath
        Chemin complet vers le fichier CSV de rapport à utiliser pour le
        retour en arrière.
#>
function Revert-Promotion {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ReportPath
    )
    Import-ADModule
    if (-not (Test-Path $ReportPath)) {
        Write-Host "Fichier de rapport introuvable : $ReportPath" -ForegroundColor Red
        return
    }
    # Importer le rapport (UTF-8)
    try {
        $entries = Import-Csv -Path $ReportPath -Delimiter ';' -Encoding UTF8
    } catch {
        Write-Host "Erreur de lecture du fichier : $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    if (-not $entries) {
        Write-Host "Le rapport est vide." -ForegroundColor Yellow
        return
    }
    Write-Host "Restauration des élèves listés dans le rapport…" -ForegroundColor Cyan
    foreach ($entry in $entries) {
        $sam = $entry.SamAccountName
        $oldDesc = $entry.AncienneDescription
        $newDesc = $entry.NouvelleDescription
        if (-not $sam -or -not $oldDesc -or -not $newDesc) {
            continue
        }
        $user = Get-ADUser -Identity $sam -Properties Description,DistinguishedName -ErrorAction SilentlyContinue
        if (-not $user) {
            Write-Host "Utilisateur introuvable : $sam" -ForegroundColor Yellow
            continue
        }
        # Déterminer les niveaux (groupes) à partir des descriptions
        $oldCode = ($oldDesc -split ' ')[0]
        $newCode = ($newDesc -split ' ')[0]
        $oldLevel = Get-GeneralGroup -CodeClasse $oldCode
        $newLevel = Get-GeneralGroup -CodeClasse $newCode
        # Restaurer la description
        try {
            Set-ADUser -Identity $sam -Description $oldDesc -ErrorAction Stop
        } catch {
            Write-Host "Erreur lors de la restauration de la description pour $sam : $($_.Exception.Message)" -ForegroundColor Yellow
        }
        # Déplacer dans l’OU d’origine si elle diffère de l’OU actuelle
        try {
            $oldOU = Get-OU -CodeClasse $oldLevel
            $newOU = Get-OU -CodeClasse $newLevel
            if ($oldOU -and $newOU -and ($oldOU -ne $newOU)) {
                # Déplacer l’objet utilisateur vers l’ancienne OU
                Move-ADObject -Identity $user -TargetPath $oldOU -Confirm:$false -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Host "Erreur lors du déplacement de $sam vers l’OU $oldOU : $($_.Exception.Message)" -ForegroundColor Yellow
        }
        # Ajuster l’appartenance aux groupes
        try {
            if ($newLevel -and $newLevel -ne 'Inconnu') {
                Remove-ADGroupMember -Identity $newLevel -Members $sam -Confirm:$false -ErrorAction SilentlyContinue
            }
        } catch {}
        try {
            if ($oldLevel -and $oldLevel -ne 'Inconnu') {
                Add-ADGroupMember -Identity $oldLevel -Members $sam -ErrorAction SilentlyContinue
            }
        } catch {}
        Write-Host "Compte restauré : $sam" -ForegroundColor Green
    }
    Write-Host "Retour en arrière terminé." -ForegroundColor Cyan
}

# -----------------------------------------------------------------------------
# Suppression en masse d’élèves sortants depuis un fichier CSV
# -----------------------------------------------------------------------------
<#
    Remove-StudentsFromCsv
    ----------------------
    Supprime des comptes d’élèves dans Active Directory à partir d’une liste
    fournie dans un fichier CSV. Le fichier doit contenir au minimum les
    colonnes `GivenName` (prénom) et `Surname` (nom). Une colonne
    supplémentaire (par exemple `Status`) peut mentionner « sortie » mais elle
    n’est pas utilisée directement. Pour chaque enregistrement :
      – le script recherche un ou plusieurs comptes correspondant au
        prénom/nom dans l’OU des élèves ;
      – compile une liste des comptes trouvés et de ceux introuvables ;
      – affiche un récapitulatif et demande confirmation avant suppression ;
      – supprime les comptes et enregistre un rapport des actions.
    .PARAMETER CSVPath
        Chemin vers le fichier CSV contenant la liste des élèves sortants.
    .PARAMETER ReportFilePath
        Chemin où enregistrer le rapport de suppression (facultatif).
    .PARAMETER WhatIf
        Si spécifié, simule les suppressions sans modifier Active Directory.

    Celle-ci elle déconne pas, donc il faut bien vérifier le contenu avant, je n'ai pas encore conçu le retour en arrière. Mais j'ai déjà préparé le log pour la future fonction.
#>
function Remove-StudentsFromCsv {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CSVPath,
        [string]$ReportFilePath = '',
        [switch]$WhatIf,
        # Encodage du fichier CSV (UTF8 par défaut). Ajustez si les noms comportent des caractères accentués mal interprétés【582535038188393†L39-L75】.
        [string]$Encoding = 'UTF8'
    )
    Import-ADModule
    if (-not (Test-Path $CSVPath)) {
        Write-Host "Fichier CSV introuvable : $CSVPath" -ForegroundColor Red
        return
    }
    # Importer le fichier CSV avec l’encodage spécifié
    $records = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding $Encoding
    if (-not $records) {
        Write-Host "Aucune donnée dans le fichier CSV." -ForegroundColor Yellow
        return
    }
    $toDelete = @()
    $notFound = @()
    foreach ($r in $records) {
        # Récupérer le prénom et le nom en tenant compte de différents en-têtes possibles (GivenName/Prenom, Surname/Nom)
        $given = $null
        $sur   = $null
        if ($r.PSObject.Properties['GivenName']) { $given = $r.GivenName }
        elseif ($r.PSObject.Properties['Prenom']) { $given = $r.Prenom }
        if ($r.PSObject.Properties['Surname']) { $sur = $r.Surname }
        elseif ($r.PSObject.Properties['Nom']) { $sur = $r.Nom }
        if (-not $given -or -not $sur) { continue }
        # Éliminer les espaces superflus
        $given = $given.Trim()
        $sur   = $sur.Trim()

        # Normaliser le prénom et le nom (suppression des accents et conversion en majuscules)
        $given = Normalize-String -InputString $given
        $sur   = Normalize-String -InputString $sur
        # Si une colonne de classe actuelle est renseignée (Codeclasse ou CodeClasse), ignorer la suppression pour éviter une erreur
        $currentClass = $null
        if ($r.PSObject.Properties['Codeclasse']) { $currentClass = $r.Codeclasse }
        elseif ($r.PSObject.Properties['Code classe']) { $currentClass = $r.'Code classe' }
        elseif ($r.PSObject.Properties['CodeClasse']) { $currentClass = $r.CodeClasse }
        if ($currentClass -and $currentClass.Trim() -ne '') {
            Write-Host "Élève $given $sur possède une classe actuelle ($currentClass). Suppression ignorée." -ForegroundColor DarkYellow
            continue
        }
        # Déterminer la classe précédente s’il existe une colonne correspondante (Codeclasseprec, CodeClassePrec, Code Classe prec., etc.)
        $prevClass = $null
        if ($r.PSObject.Properties['Codeclasseprec']) { $prevClass = $r.Codeclasseprec }
        elseif ($r.PSObject.Properties['CodeClasse prec.']) { $prevClass = $r.'CodeClasse prec.' }
        elseif ($r.PSObject.Properties['Code Classe prec.']) { $prevClass = $r.'Code Classe prec.' }
        elseif ($r.PSObject.Properties['CodeClassePrec']) { $prevClass = $r.CodeClassePrec }
        # Construire un filtre plus précis si la classe précédente est fournie
        if ($prevClass) {
            $prevGroup = Get-GeneralGroup -CodeClasse $prevClass
            # Déterminer le filtre sur la description : on inclut la classe complète
            # et, si connu, le niveau général (ex. TERMINALE) pour couvrir les deux cas
            $descFilter = "(Description -like '$prevClass*')"
            if ($prevGroup -and $prevGroup -ne 'Inconnu') {
                $descFilter = "($descFilter -or Description -like '$prevGroup*')"
            }
            $filter = "(GivenName -like '$given*' -and Surname -like '$sur*' -and $descFilter)"
        } else {
            $filter = "(GivenName -like '$given*' -and Surname -like '$sur*')"
        }
        $found = @(Get-ADUser -Filter $filter -SearchBase $ElevesOU -SearchScope Subtree -Properties SamAccountName, GivenName, Surname, Description)
        if ($found.Count -gt 0) {
            $toDelete += $found
        } else {
            $notFound += "$given $sur"
        }
    }
    Write-Host "Comptes trouvés : $($toDelete.Count)" -ForegroundColor Cyan
    foreach ($u in $toDelete) {
        Write-Host "$($u.GivenName) $($u.Surname) - $($u.SamAccountName)" -ForegroundColor Yellow
    }
    if ($notFound.Count -gt 0) {
        Write-Host "Aucun compte trouvé pour :" -ForegroundColor DarkGray
        $notFound | ForEach-Object { Write-Host "  $_" }
    }
    if ($toDelete.Count -eq 0) {
        Write-Host "Aucun compte à supprimer." -ForegroundColor Yellow
        return
    }
    # Demander confirmation sauf en mode WhatIf
    if (-not $WhatIf) {
        $confirm = Read-Host -Prompt "Confirmez-vous la suppression de ces comptes ? (O/N)"
        if ($confirm -notmatch '^[oOyY]') {
            Write-Host "Suppression annulée." -ForegroundColor Yellow
            return
        }
    }
    $report = @()
    foreach ($user in $toDelete) {
        $entry = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            NomPrenom      = "$($user.GivenName) $($user.Surname)"
            Status         = ''
        }
        try {
            if ($WhatIf) {
                $entry.Status = 'Simulé'
            } else {
                Remove-ADUser -Identity $user -Confirm:$false -ErrorAction Stop
                $entry.Status = 'Supprimé'
            }
        } catch {
            $entry.Status = "Erreur : $($_.Exception.Message)"
        }
        $report += $entry
    }
    # Exporter le rapport
    if ([string]::IsNullOrWhiteSpace($ReportFilePath)) {
        $timestamp = (Get-Date -Format 'yyyyMMdd_HHmmss')
        $ReportFilePath = "suppression_eleves_$timestamp.csv"
    }
    try {
        $report | Export-Csv -Path $ReportFilePath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        Write-Host "Rapport sauvegardé dans : $ReportFilePath" -ForegroundColor Cyan
    } catch {
        Write-Host "Impossible d’enregistrer le rapport : $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# -----------------------------------------------------------------------------
# Suppression d’un élève par SamAccountName (saisie utilisateur) #<-# Pareil pour celle-ci pas de retour en arrière... pour le moment.
# -----------------------------------------------------------------------------
function Remove-StudentBySam {
    param(
        [string]$ReportFilePath = ''
    )
    Import-ADModule

    $sam = (Read-Host -Prompt "SamAccountName de l'élève à supprimer").Trim()
    if ([string]::IsNullOrWhiteSpace($sam)) {
        Write-Host "SamAccountName vide. Abandon." -ForegroundColor Yellow
        return
    }

    # 1) Recherche ciblée sous Eleves
    $results = @()
    try {
        $u = Get-ADUser -Identity $sam -SearchBase $ElevesOU -SearchScope Subtree -Properties GivenName,Surname,Description,DistinguishedName
        if ($u) { $results += $u }
    } catch { }

    # 2) Repli : recherche dans tout le domaine si rien trouvé
    if ($results.Count -eq 0) {
        try {
            $results = @(Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" -Properties GivenName,Surname,Description,DistinguishedName)
        } catch { $results = @() }
    }

    if ($results.Count -eq 0) {
        Write-Host "Aucun utilisateur '$sam' trouvé dans le domaine." -ForegroundColor Yellow
        return
    }

    # 3) Sélection si plusieurs correspondances
    $user = $null
    if ($results.Count -gt 1) {
        Write-Host "Plusieurs comptes trouvés :" -ForegroundColor Cyan
        $i = 1
        foreach ($r in $results) {
            Write-Host ("[{0}] {1} {2}  |  {3}  |  {4}" -f $i, $r.GivenName, $r.Surname, $r.SamAccountName, $r.DistinguishedName)
            $i++
        }
        do {
            $idx = Read-Host "Numéro à supprimer"
        } while (-not ($idx -as [int]) -or [int]$idx -lt 1 -or [int]$idx -gt $results.Count)
        $user = $results[[int]$idx - 1]
    } else {
        $user = $results[0]
    }

    # 4) Récap + confirmation
    Write-Host "Utilisateur sélectionné :" -ForegroundColor Cyan
    Write-Host ("  Nom : {0} {1}" -f $user.GivenName, $user.Surname)
    Write-Host ("  SAM : {0}" -f $user.SamAccountName)
    Write-Host ("  Desc: {0}" -f $user.Description)
    Write-Host ("  DN  : {0}" -f $user.DistinguishedName)

    $confirm = Read-Host -Prompt "Confirmez-vous la suppression de ce compte ? (O/N)"
    if ($confirm -notmatch '^[oOyY]') {
        Write-Host "Suppression annulée." -ForegroundColor Yellow
        return
    }

    # 5) Suppression + rapport
    $status = ''
    try {
        Remove-ADUser -Identity $user -Confirm:$false -ErrorAction Stop
        $status = 'Supprimé'
        Write-Host "Compte supprimé : $($user.SamAccountName)" -ForegroundColor Green
    } catch {
        $status = "Erreur : $($_.Exception.Message)"
        Write-Host $status -ForegroundColor Red
    }

    if ([string]::IsNullOrWhiteSpace($ReportFilePath)) {
        $timestamp = (Get-Date -Format 'yyyyMMdd_HHmmss')
        $ReportFilePath = "suppression_eleve_unique_$timestamp.csv"
    }
    try {
        [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            NomPrenom      = "$($user.GivenName) $($user.Surname)"
            Description    = $user.Description
            DN             = $user.DistinguishedName
            Status         = $status
        } | Export-Csv -Path $ReportFilePath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        Write-Host "Rapport sauvegardé dans : $ReportFilePath" -ForegroundColor Cyan
    } catch {
        Write-Host "Impossible d’enregistrer le rapport : $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Lancer le menu principal
# Car oui c'est toujours mieux d'avoir un menu à lancer.
try {
    Import-ADModule
    Show-Menu
} catch {
    Write-Host "Erreur : $($_.Exception.Message)" -ForegroundColor Red
}
