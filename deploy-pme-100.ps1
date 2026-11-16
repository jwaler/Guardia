<#
.SYNOPSIS
    deploy-pme-100.ps1 : peuple guardia.lab comme une PME de 100 personnes.

.DESCRIPTION
    Version script de la procédure /p/ad-pme-100-personnes (guardia-studio,
    importer/proc_ad_pme.py) et du visuel ad-pme-script.html. Il crée :

      22 OU          GUARDIA, ses 5 OU de premier niveau, les 9 OU de service,
                     Metier / Ressources / Administration, Postes-Fixes /
                     Portables / Serveurs, et Quarantaine à la racine du domaine
      22 groupes     10 GG_ (9 services + GG_Tous-Salaries), 9 DL_, T0_ T1_ T2_,
                     et les emboîtements GG_ -> DL_ (AGDLP)
      100 comptes    p.nom@guardia.lab, rangés dans l'OU de leur service et dans
                     leur GG_, mot de passe aléatoire à changer au premier logon
      6 comptes adm- un par informaticien, T0_Admins-AD + Admins du domaine,
                     et le PSO-Admins (20 caractères, 90 jours)
      38 ordinateurs PC-COM-01..18 (Postes-Fixes), PT-NOM-01..20 (Portables),
                     puis redircmp / redirusr

    Le script s'arrête à l'annuaire. Le serveur de fichiers, les partages,
    les droits NTFS sur les groupes DL_, les dossiers personnels et les
    lecteurs réseau se font ensuite à la main, dans la suite de la procédure.

    Idempotent : chaque objet est testé avant d'être créé ; relancer le script
    affiche « deja la » pour ce qui existe et ne crée que ce qui manque.

    Reprise du TP Finance (/p/exo-ad-finance, puis /p/ad-partage-smb) : si les
    élèves ont déjà monté, à la main, l'OU Finance à la racine, le groupe
    G_Finance, les comptes j.moreau / k.benali / l.petit / s.durand et les
    GPO U_Finance_*, le script ADOPTE ce travail au lieu de le doubler :
      - les comptes trouvés partent dans OU=Finance,OU=Utilisateurs,OU=GUARDIA
        et comptent parmi les 8 de la Finance (le script ne génère que le
        reste) ; ni mot de passe, ni état, ni expiration, ni gestionnaire, ni
        fonction ne sont touchés, et ils n'entrent pas dans mots-de-passe.csv ;
      - ils entrent dans GG_Finance ; G_Finance est conservé (le partage du TP
        lui donne ses droits) et rangé dans Groupes > Metier ;
      - les 3 GPO U_Finance_* sont reliées à la nouvelle OU, puis déliées de
        l'ancienne ; l'ancienne OU est supprimée si elle est vide.
    Sans TP, rien de tout cela ne s'affiche ni ne change.

    Entrée  : C:\Lab\salaries.csv (Prenom;Nom;Service;Fonction;Bureau). S'il
              n'existe pas, le script l'écrit avec une liste FIXE (pas de
              hasard : deux lancements donnent les mêmes 100 personnes).
    Sortie  : C:\Lab\mots-de-passe.csv, les mots de passe des comptes créés
              pendant CE lancement. À supprimer après distribution.

    À lancer sur DC01, dans un PowerShell ouvert en administrateur, avec un
    compte Admins du domaine.

    Les chaînes affichées sont volontairement SANS accents : la console de
    Windows PowerShell 5.1 les abîme selon la page de code.

.PARAMETER Domaine
    Nom DNS du domaine (défaut guardia.lab).

.PARAMETER WhatIf
    Simulation : le script lit l'annuaire, affiche ce qu'il ferait,
    et n'écrit RIEN (ni objet AD, ni fichier).

.EXAMPLE
    Set-ExecutionPolicy -Scope Process Bypass; .\deploy-pme-100.ps1

.EXAMPLE
    .\deploy-pme-100.ps1 -WhatIf
#>
param(
    [string]$Domaine = 'guardia.lab',
    [switch]$WhatIf
)

# ====================================================================== réglages
$DossierLab   = 'C:\Lab'
$CsvSalaries  = Join-Path $DossierLab 'salaries.csv'
$CsvMdp       = Join-Path $DossierLab 'mots-de-passe.csv'

# Les 9 services : nom de l'OU (= nom du groupe GG_), code, effectif, lettre de bureau
$Services = @(
    @{ Nom = 'Direction';    Code = 'DIR'; Effectif = 4;  Bureau = 'D' }
    @{ Nom = 'Commercial';   Code = 'COM'; Effectif = 18; Bureau = 'A' }
    @{ Nom = 'Marketing';    Code = 'MKT'; Effectif = 7;  Bureau = 'M' }
    @{ Nom = 'Technique';    Code = 'TEC'; Effectif = 24; Bureau = 'C' }
    @{ Nom = 'Production';   Code = 'PRD'; Effectif = 16; Bureau = 'P' }
    @{ Nom = 'Informatique'; Code = 'INF'; Effectif = 6;  Bureau = 'I' }
    @{ Nom = 'RH';           Code = 'RH';  Effectif = 5;  Bureau = 'R' }
    @{ Nom = 'Finance';      Code = 'FIN'; Effectif = 8;  Bureau = 'B' }
    @{ Nom = 'Logistique';   Code = 'LOG'; Effectif = 12; Bureau = 'L' }
)
$NomsServices = $Services | ForEach-Object { $_.Nom }

# Les personnes que la procédure cite par leur nom : elles gardent ce nom.
# (m.durand en Finance, p.lefevre au Commercial, s.marchand à la Technique,
#  t.dubois à l'Informatique, dont le compte d'admin est adm-tdubois)
$PersonnesCitees = @(
    @{ Prenom = 'Marie';  Nom = 'Durand';   Service = 'Finance';      Fonction = 'Comptable';               Bureau = 'B12' }
    @{ Prenom = 'Paul';   Nom = 'Lefevre';  Service = 'Commercial';   Fonction = 'Charge d affaires';       Bureau = 'A04' }
    @{ Prenom = 'Sophie'; Nom = 'Marchand'; Service = 'Technique';    Fonction = 'Dessinatrice projeteuse'; Bureau = 'C21' }
    @{ Prenom = 'Thomas'; Nom = 'Dubois';   Service = 'Informatique'; Fonction = 'Administrateur systemes'; Bureau = 'I01' }
)

# Les 96 autres : listes FIXES, sans accents, noms tous différents
# (donc identifiants p.nom tous différents). La personne n° i reçoit le
# nom n° i et le prénom n° (i modulo 40).
$Prenoms = @(
    'Julie','Nicolas','Camille','Julien','Laura','Antoine','Sarah','Maxime','Emma','Hugo',
    'Chloe','Lucas','Manon','Louis','Ines','Nathan','Lea','Theo','Clara','Romain',
    'Pauline','Alexandre','Elodie','Kevin','Celine','Mathieu','Audrey','Vincent','Claire','Damien',
    'Lucie','Guillaume','Margaux','Benoit','Helene','Florian','Amandine','Quentin','Aurelie','Jerome'
)
$Noms = @(
    'Martin','Bernard','Petit','Robert','Richard','Moreau','Laurent','Simon','Michel','Leroy',
    'Garcia','David','Bertrand','Roux','Vincent','Fournier','Morel','Girard','Andre','Mercier',
    'Dupont','Lambert','Bonnet','Francois','Martinez','Legrand','Garnier','Faure','Rousseau','Blanc',
    'Guerin','Muller','Henry','Roussel','Perrin','Morin','Mathieu','Clement','Gauthier','Dumont',
    'Lopez','Fontaine','Chevalier','Robin','Masson','Sanchez','Gerard','Nguyen','Boyer','Denis',
    'Lemaire','Duval','Joly','Gautier','Roger','Roche','Roy','Noel','Meyer','Meunier',
    'Perez','Marchal','Dufour','Blanchard','Barbier','Brun','Dumas','Brunet','Schmitt','Leroux',
    'Colin','Fernandez','Renard','Arnaud','Rolland','Caron','Aubert','Giraud','Leclerc','Vidal',
    'Bourgeois','Renaud','Lemoine','Picard','Gaillard','Lacroix','Fabre','Dupuis','Rodriguez','Hubert',
    'Carpentier','Moulin','Poirier','Deschamps','Huet','Vasseur'
)
$Fonctions = @{
    Direction    = 'Directeur general','Directrice administrative','Directeur des operations','Assistante de direction'
    Commercial   = 'Charge d affaires','Commercial sedentaire','Technico-commercial','Responsable grands comptes','Assistant commercial'
    Marketing    = 'Chef de produit','Charge de communication','Graphiste','Webmaster'
    Technique    = 'Ingenieur d etudes','Dessinateur projeteur','Technicien methodes','Chef de projet'
    Production   = 'Chef d atelier','Operateur de production','Technicien de maintenance','Controleur qualite'
    Informatique = 'Administrateur systemes','Technicien support','Administrateur reseau','Responsable informatique'
    RH           = 'Responsable RH','Charge de recrutement','Gestionnaire paie'
    Finance      = 'Comptable','Controleur de gestion','Tresorier','Assistant comptable'
    Logistique   = 'Magasinier','Responsable logistique','Gestionnaire des stocks','Cariste'
}

# Le TP « le service Finance arrive » (/p/exo-ad-finance), fait à la main
# AVANT ce script : OU Finance à la racine du domaine, groupe G_Finance, quatre
# comptes et trois GPO liées à cette OU. Les noms sont ceux de l'énoncé ; les
# prénoms et noms servent seulement à écrire salaries.csv (sans accents, et
# Get-Login retombe bien sur l'identifiant).
$TpGroupe  = 'G_Finance'
$TpComptes = @(
    @{ Login = 'j.moreau'; Prenom = 'Julie'; Nom = 'Moreau'; Fonction = 'Responsable financiere' }
    @{ Login = 'k.benali'; Prenom = 'Karim'; Nom = 'Benali'; Fonction = 'Comptable' }
    @{ Login = 'l.petit';  Prenom = 'Lea';   Nom = 'Petit';  Fonction = 'Controleuse de gestion' }
    @{ Login = 's.durand'; Prenom = 'Sami';  Nom = 'Durand'; Fonction = 'Stagiaire' }
)
$TpGpo     = @('U_Finance_PanneauConfig', 'U_Finance_EcranVeille', 'U_Finance_USB')
$LoginsTp  = @($TpComptes | ForEach-Object { $_.Login })

# Les emboîtements AGDLP, exactement ceux de la procédure. Le script crée les
# 9 groupes DL_ et y met les GG_ ; les droits de ces DL_ sur les dossiers et
# les partages se posent plus tard, à la main, dans la procédure.
$Emboitements = @(
    @{ DL = 'DL_Commun_Modif';      GG = @('GG_Tous-Salaries') }
    @{ DL = 'DL_Echange_Modif';     GG = @('GG_Tous-Salaries') }
    @{ DL = 'DL_Logiciels_Lecture'; GG = @('GG_Tous-Salaries') }
    @{ DL = 'DL_Direction_Modif';   GG = @('GG_Direction') }
    @{ DL = 'DL_RH_Modif';          GG = @('GG_RH') }
    @{ DL = 'DL_Finance_Modif';     GG = @('GG_Finance') }
    @{ DL = 'DL_Commercial_Modif';  GG = @('GG_Commercial') }
    @{ DL = 'DL_Technique_Modif';   GG = @('GG_Technique', 'GG_Production') }
    @{ DL = 'DL_Logiciels_Modif';   GG = @('GG_Informatique') }
)

# ====================================================================== affichage et compteurs
$script:Simu    = $WhatIf.IsPresent
$script:AdOk    = $false
$script:Bloc    = ''
$script:Stats   = [ordered]@{}
$script:Totaux  = @{ OU = 0; Comptes = 0; Groupes = 0; Ordinateurs = 0; Adm = 0 }
$script:NbErreurs = 0
$script:MotsDePasse = New-Object System.Collections.Generic.List[object]

function Section([string]$Titre) {
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host (' ' + $Titre) -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor Cyan
    $script:Bloc = $Titre
    $script:Stats[$Titre] = @{ Crees = 0; Deja = 0; Erreurs = 0 }
}

function Fin-Section([string]$Fait = 'crees', [string]$AFaire = 'a creer') {
    $s = $script:Stats[$script:Bloc]
    $verbe = if ($script:Simu) { $AFaire } else { $Fait }
    $couleur = if ($s.Erreurs) { 'Red' } else { 'Cyan' }
    Write-Host ('  -> {0} {1}, {2} deja la, {3} erreur(s)' -f $s.Crees, $verbe, $s.Deja, $s.Erreurs) -ForegroundColor $couleur
}

function Compter([string]$Cat) { if ($Cat) { $script:Totaux[$Cat]++ } }

function Note-Deja([string]$Libelle, [string]$Cat) {
    Write-Host ('  deja la   ' + $Libelle) -ForegroundColor DarkGray
    $script:Stats[$script:Bloc].Deja++
    Compter $Cat
}

function Note-Erreur([string]$Libelle) {
    Write-Host ('  ERREUR    ' + $Libelle) -ForegroundColor Red
    $script:Stats[$script:Bloc].Erreurs++
    $script:NbErreurs++
}

# Exécute $Bloc_ (ou l'annonce seulement, en simulation). Renvoie $true si c'est fait.
# $Verbe : le mot affiché quand c'est fait (cree, deplace, relie, retire...).
function Faire([string]$Libelle, [scriptblock]$Bloc_, [string]$Cat, [string]$Verbe = 'cree') {
    if ($script:Simu) {
        Write-Host ('  [simul]   ' + $Libelle) -ForegroundColor Yellow
        $script:Stats[$script:Bloc].Crees++
        Compter $Cat
        return $true
    }
    try {
        & $Bloc_ | Out-Null
        Write-Host (('  {0,-10}' -f $Verbe) + $Libelle) -ForegroundColor Green
        $script:Stats[$script:Bloc].Crees++
        Compter $Cat
        return $true
    } catch {
        Note-Erreur ($Libelle + ' : ' + $_.Exception.Message)
        return $false
    }
}

function Arret([string]$Message) {
    Write-Host ''
    Write-Host ('ARRET : ' + $Message) -ForegroundColor Red
    Write-Host 'Rien n a ete cree. Corriger le prerequis, puis relancer le script.' -ForegroundColor Red
    exit 1
}

# ====================================================================== outils
function Sans-Accents([string]$s) {
    ($s.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '')
}

function Get-Login([string]$Prenom, [string]$Nom) {
    $l = ('{0}.{1}' -f (Sans-Accents $Prenom).Substring(0, 1), (Sans-Accents $Nom)).ToLower()
    $l -replace '[^a-z0-9.]', ''
}

$script:Rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
function Get-Alea([int]$Max) {
    $b = New-Object byte[] 4
    $script:Rng.GetBytes($b)
    [int]([BitConverter]::ToUInt32($b, 0) % [uint32]$Max)
}

# Mot de passe aléatoire : au moins une majuscule, une minuscule, un chiffre et
# un symbole. Ni ; ni " (le CSV), ni = + - @ en tête (formules Excel).
function New-MotDePasse([int]$Longueur) {
    $jeux = 'ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!#%?*$'
    $tous = -join $jeux
    $car = New-Object System.Collections.Generic.List[char]
    foreach ($j in $jeux) { $car.Add($j[(Get-Alea $j.Length)]) }
    while ($car.Count -lt $Longueur) { $car.Add($tous[(Get-Alea $tous.Length)]) }
    for ($i = $car.Count - 1; $i -gt 0; $i--) {
        $r = Get-Alea ($i + 1); $t = $car[$i]; $car[$i] = $car[$r]; $car[$r] = $t
    }
    if ('!#%?*$'.Contains([string]$car[0])) { $t = $car[0]; $car[0] = $car[1]; $car[1] = $t }
    -join $car
}

function Test-DN([string]$Dn) {
    if (-not $script:AdOk) { return $false }
    try { $null = Get-ADObject -Identity $Dn -ErrorAction Stop; $true } catch { $false }
}

function Get-Groupe([string]$Identite) {
    if (-not $script:AdOk) { return $null }
    try { Get-ADGroup -Identity $Identite -Properties member -ErrorAction Stop } catch { $null }
}

function Get-ParSam([string]$Sam) {
    if (-not $script:AdOk) { return $null }
    Get-ADObject -Filter "sAMAccountName -eq '$Sam'" -ErrorAction SilentlyContinue
}

function Ensure-OU([string]$Nom, [string]$Parent) {
    $dnOU = "OU=$Nom,$Parent"
    if (Test-DN $dnOU) { Note-Deja "OU $dnOU" 'OU' }
    else {
        [void](Faire "OU $dnOU" { New-ADOrganizationalUnit -Name $Nom -Path $Parent -ProtectedFromAccidentalDeletion $true } 'OU')
    }
    $dnOU
}

function Ensure-Groupe([string]$Nom, [string]$Portee, [string]$Chemin, [string]$Desc) {
    if (Get-Groupe $Nom) { Note-Deja "groupe $Nom" 'Groupes' }
    else {
        [void](Faire "groupe $Nom ($Portee)" {
            New-ADGroup -Name $Nom -SamAccountName $Nom -GroupScope $Portee -GroupCategory Security -Path $Chemin -Description $Desc
        } 'Groupes')
    }
}

# Met $Membre (sAMAccountName) dans $Groupe (nom ou SID), s'il n'y est pas déjà.
function Ensure-Membre([string]$Groupe, [string]$Membre, [string]$NomAffiche) {
    if (-not $NomAffiche) { $NomAffiche = $Groupe }
    $lib = "$Membre dans $NomAffiche"
    $g = Get-Groupe $Groupe
    $m = Get-ParSam $Membre
    if ($g -and $m -and ($g.member -contains $m.DistinguishedName)) { Note-Deja $lib; return }
    if (-not $script:Simu -and (-not $g -or -not $m)) { Note-Erreur "$lib : groupe ou membre introuvable"; return }
    [void](Faire $lib { Add-ADGroupMember -Identity $g.DistinguishedName -Members $m.DistinguishedName })
}

# Le DN du parent : tout après la première virgule NON échappée
# (un CN peut contenir « \, » : CN=Moreau\, Julie,OU=...).
function Get-Parent([string]$Dn) {
    if ($Dn -match '^(?:[^,\\]|\\.)+,(.+)$') { $Matches[1] } else { '' }
}

# Déplace un objet sans rien changer d'autre. S'il est protégé contre la
# suppression accidentelle, le déplacement serait refusé : on lève la
# protection le temps du Move-ADObject, puis on la remet.
function Deplacer([string]$Dn, [string]$Cible, [bool]$Protege) {
    if ($Protege) { Set-ADObject -Identity $Dn -ProtectedFromAccidentalDeletion $false }
    $o = Move-ADObject -Identity $Dn -TargetPath $Cible -PassThru
    if ($Protege) { Set-ADObject -Identity $o.DistinguishedName -ProtectedFromAccidentalDeletion $true }
}

# Les liens de GPO posés sur une OU (vide si l'OU n'existe pas encore)
function Get-Liens([string]$DnOU) {
    if (-not $DnOU -or -not (Test-DN $DnOU)) { return @() }
    @((Get-GPInheritance -Target $DnOU -Domain $Domaine).GpoLinks)
}

# ====================================================================== 0. prérequis
Section '0. Prerequis'
if ($script:Simu) {
    Write-Host '  MODE SIMULATION (-WhatIf) : lecture seule, rien ne sera cree.' -ForegroundColor Yellow
}

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin -and -not $script:Simu) { Arret 'ouvrir PowerShell en tant qu administrateur.' }

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module GroupPolicy -ErrorAction Stop
    $ad = Get-ADDomain -Identity $Domaine -ErrorAction Stop
    $script:AdOk = $true
} catch {
    if (-not $script:Simu) { Arret ("annuaire $Domaine injoignable ou modules RSAT absents : " + $_.Exception.Message) }
    Write-Host '  [simul]   annuaire injoignable : tout sera affiche comme a creer' -ForegroundColor Yellow
}

if ($script:AdOk) {
    $dn      = $ad.DistinguishedName
    $netbios = $ad.NetBIOSName
    $sidDom  = $ad.DomainSID.Value
    $daNom   = (Get-ADGroup -Identity "$sidDom-512").Name      # Admins du domaine, quelle que soit la langue
} else {
    $dn      = 'DC=' + ($Domaine -replace '\.', ',DC=')
    $netbios = ($Domaine.Split('.')[0]).ToUpper()
    $sidDom  = ''
    $daNom   = 'Admins du domaine'
}
Write-Host ("  domaine   : $Domaine ($dn), NetBIOS $netbios") -ForegroundColor Gray
Write-Host '  prerequis controles' -ForegroundColor Gray

# ====================================================================== détection du TP Finance
# Lecture seule : rien n'est modifié ici. Le TP est « trouvé » dès qu'il en
# reste une trace : l'OU Finance à la racine, le groupe G_Finance, ou l'un des
# quatre comptes (où qu'il soit rangé dans l'annuaire).
$script:Tp = $null
$dnTpOU = "OU=Finance,$dn"
if ($script:AdOk) {
    $tpOuExiste = Test-DN $dnTpOU
    $tpGrp = $null
    try { $tpGrp = Get-ADGroup -Identity $TpGroupe -Properties member, ProtectedFromAccidentalDeletion -ErrorAction Stop } catch { }
    $tpCpt = New-Object System.Collections.Generic.List[object]
    foreach ($c in $TpComptes) {
        $u = Get-ADUser -Filter "SamAccountName -eq '$($c.Login)'" -Properties Title, ProtectedFromAccidentalDeletion -ErrorAction SilentlyContinue
        if (-not $u) { continue }
        # la fonction saisie par l'élève, sans accents ni ; (elle ne sert qu'au CSV)
        $fct = if ($u.Title) { (Sans-Accents $u.Title) -replace '[;"]', ' ' } else { $c.Fonction }
        $tpCpt.Add([pscustomobject]@{
            Login = $c.Login; Prenom = $c.Prenom; Nom = $c.Nom; Fonction = $fct
            Dn = $u.DistinguishedName; Protege = [bool]$u.ProtectedFromAccidentalDeletion
        })
    }
    if ($tpOuExiste -or $tpGrp -or $tpCpt.Count) {
        $script:Tp = [pscustomobject]@{
            DnOU    = $(if ($tpOuExiste) { $dnTpOU } else { '' })
            Groupe  = $tpGrp
            Comptes = @($tpCpt)
        }
    }
}

# ====================================================================== 1. le fichier des salariés
Section '1. Fichier des salaries'
if ($script:Tp) {
    $traces = @()
    if ($script:Tp.DnOU)   { $traces += "OU $dnTpOU" }
    if ($script:Tp.Groupe) { $traces += $TpGroupe }
    if ($script:Tp.Comptes.Count) { $traces += ($script:Tp.Comptes | ForEach-Object { $_.Login }) }
    Write-Host ('  TP Finance trouve : ' + ($traces -join ', ')) -ForegroundColor Gray
    Write-Host ('  -> repris en section 3 bis ; ses {0} compte(s) comptent parmi les 8 de la Finance' -f $script:Tp.Comptes.Count) -ForegroundColor Gray
}
$salaries = $null
if (Test-Path -LiteralPath $CsvSalaries) {
    $salaries = @(Import-Csv -LiteralPath $CsvSalaries -Delimiter ';' -Encoding UTF8)
    Note-Deja ("$CsvSalaries ({0} lignes), il est lu tel quel" -f $salaries.Count)
} else {
    # génération déterministe : les personnes citées d'abord, puis les comptes
    # du TP Finance s'il y en a, puis la liste fixe pour compléter l'effectif
    $liste = New-Object System.Collections.Generic.List[object]
    $iNom = 0
    foreach ($s in $Services) {
        $deja = @($PersonnesCitees | Where-Object { $_.Service -eq $s.Nom })
        foreach ($p in $deja) { $liste.Add([pscustomobject]$p) }
        $occupes = $deja.Count
        if ($s.Nom -eq 'Finance' -and $script:Tp) {
            foreach ($c in $script:Tp.Comptes) {
                $occupes++
                $liste.Add([pscustomobject]@{
                    Prenom = $c.Prenom; Nom = $c.Nom; Service = 'Finance'
                    Fonction = $c.Fonction; Bureau = ('{0}{1:D2}' -f $s.Bureau, $occupes)
                })
            }
        }
        for ($n = $deja.Count + 1; $n -le $s.Effectif; $n++) {
            # Le nom n° $iNom est consommé même quand la place est prise par un
            # compte du TP : les services suivants gardent ainsi exactement les
            # mêmes noms, avec ou sans TP.
            $prenom = $Prenoms[$iNom % $Prenoms.Count]
            $nom    = $Noms[$iNom]
            $iNom++
            if ($n -le $occupes) { continue }
            # collision avec un compte du TP : impossible avec les listes fixes
            # (ni Benali ni Durand, et Moreau / Petit n'y portent pas J / L),
            # mais on le vérifie au lieu de le supposer
            if ($LoginsTp -contains (Get-Login $prenom $nom)) {
                Note-Erreur "nom genere $prenom $nom : identifiant $(Get-Login $prenom $nom) deja pris par le TP Finance, ligne ignoree"
                continue
            }
            $f = $Fonctions[$s.Nom]
            $liste.Add([pscustomobject]@{
                Prenom   = $prenom
                Nom      = $nom
                Service  = $s.Nom
                Fonction = $f[($n - 1) % $f.Count]
                Bureau   = ('{0}{1:D2}' -f $s.Bureau, $n)
            })
        }
    }
    $salaries = @($liste | Select-Object Prenom, Nom, Service, Fonction, Bureau)
    [void](Faire ("$CsvSalaries ({0} salaries, liste fixe)" -f $salaries.Count) {
        if (-not (Test-Path -LiteralPath $DossierLab)) { New-Item -ItemType Directory -Path $DossierLab -Force }
        $lignes = @('Prenom;Nom;Service;Fonction;Bureau')
        $lignes += $salaries | ForEach-Object { '{0};{1};{2};{3};{4}' -f $_.Prenom, $_.Nom, $_.Service, $_.Fonction, $_.Bureau }
        $lignes | Out-File -LiteralPath $CsvSalaries -Encoding UTF8
    })
}

# contrôle du fichier : service connu, identifiant unique et de 20 caractères maximum
$valides = New-Object System.Collections.Generic.List[object]
$vus = @{}
foreach ($l in $salaries) {
    $login = Get-Login $l.Prenom $l.Nom
    if ($NomsServices -notcontains $l.Service) { Note-Erreur "ligne $($l.Prenom) $($l.Nom) : service '$($l.Service)' inconnu"; continue }
    if ($login.Length -gt 20) { Note-Erreur "ligne $($l.Prenom) $($l.Nom) : identifiant $login de plus de 20 caracteres"; continue }
    if ($vus.ContainsKey($login)) { Note-Erreur "ligne $($l.Prenom) $($l.Nom) : identifiant $login en double"; continue }
    if ($LoginsTp -contains $login -and $l.Service -ne 'Finance') {
        Note-Erreur "ligne $($l.Prenom) $($l.Nom) : identifiant $login reserve au TP Finance, service '$($l.Service)' refuse"; continue
    }
    $vus[$login] = $true
    $valides.Add([pscustomobject]@{
        Prenom = $l.Prenom; Nom = $l.Nom; Service = $l.Service
        Fonction = $l.Fonction; Bureau = $l.Bureau; Login = $login
    })
}
foreach ($s in $Services) {
    $n = @($valides | Where-Object { $_.Service -eq $s.Nom }).Count
    Write-Host ('  {0,-13} {1,3}   (attendu {2})' -f $s.Nom, $n, $s.Effectif) -ForegroundColor Gray
}
Write-Host ('  total         {0,3}   (attendu 100)' -f $valides.Count) -ForegroundColor Gray
# Un salaries.csv écrit AVANT le TP ne le connaît pas : on ne le modifie pas,
# les comptes du TP sont déplacés quand même, mais hors des 100 du fichier.
$loginsValides = @($valides | ForEach-Object { $_.Login })
$tpHorsFichier = @()
if ($script:Tp) { $tpHorsFichier = @($script:Tp.Comptes | Where-Object { $loginsValides -notcontains $_.Login }) }
if ($tpHorsFichier.Count) {
    Write-Host ('  attention : {0} absent(s) du fichier. Ils seront deplaces dans la Finance quand meme,' -f (($tpHorsFichier | ForEach-Object { $_.Login }) -join ', ')) -ForegroundColor Yellow
    Write-Host '              mais ils s ajoutent aux salaries du fichier au lieu d en faire partie.' -ForegroundColor Yellow
}
Fin-Section

# ====================================================================== 2. les 22 OU
Section '2. Unites d organisation (22)'
$dnG    = Ensure-OU 'GUARDIA' $dn
$dnUtil = Ensure-OU 'Utilisateurs' $dnG
$dnGrp  = Ensure-OU 'Groupes' $dnG
$dnOrd  = Ensure-OU 'Ordinateurs' $dnG
$dnAdm  = Ensure-OU 'Comptes-Admin' $dnG
$null   = Ensure-OU 'Comptes-Service' $dnG
foreach ($s in $NomsServices) { $null = Ensure-OU $s $dnUtil }
$dnMetier = Ensure-OU 'Metier' $dnGrp
$dnRess   = Ensure-OU 'Ressources' $dnGrp
$dnAdmin  = Ensure-OU 'Administration' $dnGrp
$dnFixes  = Ensure-OU 'Postes-Fixes' $dnOrd
$dnPort   = Ensure-OU 'Portables' $dnOrd
$null     = Ensure-OU 'Serveurs' $dnOrd
$null     = Ensure-OU 'Quarantaine' $dn
Fin-Section

# ====================================================================== 3. les 22 groupes et AGDLP
Section '3. Groupes (22) et emboitements AGDLP'
foreach ($s in $NomsServices) { Ensure-Groupe "GG_$s" 'Global' $dnMetier "Salaries du service $s" }
Ensure-Groupe 'GG_Tous-Salaries' 'Global' $dnMetier 'Les 9 groupes de service'
foreach ($e in $Emboitements) { Ensure-Groupe $e.DL 'DomainLocal' $dnRess "Acces : $($e.DL)" }
# Tiers : portée GLOBALE, car un PSO ne s'applique qu'à un utilisateur ou à un
# groupe de sécurité GLOBAL (sur un groupe de domaine local il serait ignoré).
Ensure-Groupe 'T0_Admins-AD'       'Global' $dnAdmin 'Tier 0 : annuaire et DC'
Ensure-Groupe 'T1_Admins-Serveurs' 'Global' $dnAdmin 'Tier 1 : serveurs membres'
Ensure-Groupe 'T2_Support-Poste'   'Global' $dnAdmin 'Tier 2 : postes de travail'
foreach ($s in $NomsServices) { Ensure-Membre 'GG_Tous-Salaries' "GG_$s" }
foreach ($e in $Emboitements) { foreach ($gg in $e.GG) { Ensure-Membre $e.DL $gg } }
Fin-Section

# ====================================================================== 3 bis. reprise du TP Finance
# Placée APRÈS les OU et les groupes (la cible doit exister) et AVANT la boucle
# des comptes : les comptes du TP y arrivent déjà rangés, la boucle les voit
# « deja la », ne leur donne aucun mot de passe et les compte dans les 100.
$dnFinance = "OU=Finance,$dnUtil"
if ($script:Tp) {
    Section '3 bis. Reprise du TP Finance'
    $partis   = @{}          # DN des objets qui quittent l'ancienne OU (utile en simulation)
    $repris   = @()
    $reliees  = @()

    # --- a) les comptes : déplacés, et RIEN d'autre (mot de passe, état,
    #        expiration, gestionnaire, fonction restent ceux du TP)
    Write-Host '  --- comptes du TP' -ForegroundColor Cyan
    if (-not $script:Tp.Comptes.Count) { Write-Host '  aucun des 4 comptes du TP dans l annuaire' -ForegroundColor DarkGray }
    foreach ($c in $script:Tp.Comptes) {
        $parent = Get-Parent $c.Dn
        if ($parent -eq $dnFinance) { Note-Deja "$($c.Login) dans $dnFinance" }
        elseif (Faire "$($c.Login) : $parent -> $dnFinance" { Deplacer $c.Dn $dnFinance $c.Protege } '' 'deplace') {
            $partis[$c.Dn] = $true
        }
        $repris += $c.Login
    }
    # GG_Finance : la boucle des comptes s'en charge pour ceux du fichier ;
    # ici seulement ceux qu'un salaries.csv plus ancien ne connaît pas.
    foreach ($c in $tpHorsFichier) { Ensure-Membre 'GG_Finance' $c.Login }

    # --- b) G_Finance : CONSERVÉ tel quel, et rangé avec les groupes métier.
    # Pourquoi le garder : le partage \\DC01\Finance de /p/ad-partage-smb donne
    # son droit NTFS Modification à G_Finance (par SID). Le supprimer ou le
    # remplacer casserait ce partage. Le déplacer ne change pas son SID.
    # Pourquoi ne PAS l'imbriquer dans GG_Finance (ni l'inverse) : les quatre
    # comptes sont déjà membres directs de GG_Finance (boucle des comptes),
    # donc G_Finance dans GG_Finance n'apporterait rien ; GG_Finance dans
    # G_Finance ouvrirait le partage du TP aux 8 salariés de la Finance, ce que
    # le TP n'a pas prévu. Les droits de la PME passent par GG_Finance ->
    # DL_Finance_Modif (AGDLP), posés à la main plus tard dans la procédure.
    Write-Host "  --- groupe $TpGroupe" -ForegroundColor Cyan
    $g = $script:Tp.Groupe
    if ($g) {
        $parentG = Get-Parent $g.DistinguishedName
        if ($parentG -eq $dnMetier) { Note-Deja "$TpGroupe dans $dnMetier" }
        elseif (Faire "$TpGroupe : $parentG -> $dnMetier" { Deplacer $g.DistinguishedName $dnMetier ([bool]$g.ProtectedFromAccidentalDeletion) } '' 'deplace') {
            $partis[$g.DistinguishedName] = $true
        }
        Write-Host "  $TpGroupe conserve : le partage du TP reste valable (membres inchanges, pas d imbrication avec GG_Finance)" -ForegroundColor Gray
    } else {
        Write-Host "  $TpGroupe absent : rien a deplacer, les comptes du TP passent par GG_Finance" -ForegroundColor DarkGray
    }

    # --- c) les 3 GPO du TP : liées à la nouvelle OU (même état activé /
    #        appliqué que l'ancien lien), puis déliées de l'ancienne
    Write-Host '  --- GPO du TP' -ForegroundColor Cyan
    $liensNeuf  = @(Get-Liens $dnFinance)
    $liensVieux = @(Get-Liens $script:Tp.DnOU)
    foreach ($nomGpo in $TpGpo) {
        $existe = $null
        try { $existe = Get-GPO -Name $nomGpo -Domain $Domaine -ErrorAction Stop } catch { }
        if (-not $existe) { Write-Host "  absente   GPO $nomGpo : rien a relier" -ForegroundColor DarkGray; continue }
        $vieux = @($liensVieux | Where-Object { $_.DisplayName -eq $nomGpo }) | Select-Object -First 1
        $lieOk = $true
        if (@($liensNeuf | ForEach-Object { $_.DisplayName }) -contains $nomGpo) { Note-Deja "lien $nomGpo sur $dnFinance" }
        else {
            $actif = if ($vieux -and -not $vieux.Enabled) { 'No' } else { 'Yes' }
            $force = if ($vieux -and $vieux.Enforced) { 'Yes' } else { 'No' }
            $lieOk = Faire "lien $nomGpo sur $dnFinance" {
                New-GPLink -Name $nomGpo -Target $dnFinance -Domain $Domaine -LinkEnabled $actif -Enforced $force
            } '' 'relie'
        }
        # l'ancien lien ne part que si le nouveau est en place
        if ($vieux -and $lieOk) {
            [void](Faire "lien $nomGpo retire de $($script:Tp.DnOU)" {
                Remove-GPLink -Name $nomGpo -Target $script:Tp.DnOU -Domain $Domaine
            } '' 'retire')
        }
        if ($lieOk) { $reliees += $nomGpo }
    }

    # --- d) l'ancienne OU Finance à la racine : supprimée si elle est vide
    Write-Host '  --- ancienne OU Finance' -ForegroundColor Cyan
    if (-not $script:Tp.DnOU) {
        Note-Deja "ancienne OU $dnTpOU : absente, rien a supprimer"
        $sortOU = 'absente (deja supprimee ou jamais creee)'
    } else {
        $restent = @(Get-ADObject -SearchBase $dnTpOU -SearchScope OneLevel -Filter * -ErrorAction SilentlyContinue |
                     Where-Object { -not $partis.ContainsKey($_.DistinguishedName) })
        if ($restent.Count -eq 0) {
            $ok = Faire "ancienne OU $dnTpOU : vide, protection levee puis supprimee" {
                Set-ADOrganizationalUnit -Identity $dnTpOU -ProtectedFromAccidentalDeletion $false
                Remove-ADOrganizationalUnit -Identity $dnTpOU -Confirm:$false
            } '' 'supprime'
            $sortOU = if ($ok) { 'vide, supprimee' } else { 'vide, mais la suppression a echoue (voir ERREUR)' }
        } else {
            Write-Host ("  conservee ancienne OU $dnTpOU : elle contient encore {0} objet(s)" -f $restent.Count) -ForegroundColor Yellow
            foreach ($r in $restent) { Write-Host ("            {0,-22} {1}" -f $r.ObjectClass, $r.Name) -ForegroundColor Yellow }
            $sortOU = "conservee, {0} objet(s) dedans (liste ci-dessus)" -f $restent.Count
        }
    }

    # --- e) le bilan
    Write-Host '  --- bilan de la reprise' -ForegroundColor Cyan
    Write-Host ('  comptes repris : ' + $(if ($repris.Count) { $repris -join ', ' } else { 'aucun' })) -ForegroundColor Gray
    Write-Host ('  GPO reliees    : ' + $(if ($reliees.Count) { $reliees -join ', ' } else { 'aucune' })) -ForegroundColor Gray
    Write-Host ('  ancienne OU    : ' + $sortOU) -ForegroundColor Gray
    Write-Host '  inchanges      : mots de passe, etat, expiration, gestionnaire et fonction des comptes du TP' -ForegroundColor Gray
    Fin-Section 'faits' 'a faire'
}

# ====================================================================== 4. les 100 comptes
Section '4. Comptes salaries (100)'
foreach ($u in $valides) {
    if (Get-ParSam $u.Login) { Note-Deja "$($u.Login) ($($u.Service))" 'Comptes' }
    else {
        $mdp = New-MotDePasse 14
        $ok = Faire "$($u.Login)@$Domaine  $($u.Prenom) $($u.Nom), OU $($u.Service)" {
            New-ADUser -Name "$($u.Prenom) $($u.Nom)" -GivenName $u.Prenom -Surname $u.Nom `
                -SamAccountName $u.Login -UserPrincipalName "$($u.Login)@$Domaine" `
                -DisplayName "$($u.Prenom) $($u.Nom)" -Title $u.Fonction -Department $u.Service `
                -Office $u.Bureau -Path "OU=$($u.Service),$dnUtil" `
                -AccountPassword (ConvertTo-SecureString $mdp -AsPlainText -Force) `
                -ChangePasswordAtLogon $true -Enabled $true
        } 'Comptes'
        if ($ok -and -not $script:Simu) {
            $script:MotsDePasse.Add([pscustomobject]@{ Login = $u.Login; MotDePasse = $mdp; Type = 'salarie' })
        }
    }
    Ensure-Membre "GG_$($u.Service)" $u.Login
}
Fin-Section

# ====================================================================== 5. comptes d'administration
Section '5. Comptes adm- (6) et PSO-Admins'
$informaticiens = @($valides | Where-Object { $_.Service -eq 'Informatique' })
foreach ($u in $informaticiens) {
    $adm = ('adm-' + (Sans-Accents $u.Prenom).Substring(0, 1) + (Sans-Accents $u.Nom)).ToLower() -replace '[^a-z0-9-]', ''
    if ($adm.Length -gt 20) { $adm = $adm.Substring(0, 20) }
    if (Get-ParSam $adm) { Note-Deja "$adm (compte d admin de $($u.Login))" 'Adm' }
    else {
        $mdp = New-MotDePasse 24
        $ok = Faire "$adm  dans Comptes-Admin (compte d admin de $($u.Login))" {
            New-ADUser -Name $adm -SamAccountName $adm -UserPrincipalName "$adm@$Domaine" `
                -DisplayName "$adm ($($u.Prenom) $($u.Nom))" -Description "Compte d administration de $($u.Prenom) $($u.Nom)" `
                -Path $dnAdm -AccountPassword (ConvertTo-SecureString $mdp -AsPlainText -Force) `
                -ChangePasswordAtLogon $true -Enabled $true
        } 'Adm'
        if ($ok -and -not $script:Simu) {
            $script:MotsDePasse.Add([pscustomobject]@{ Login = $adm; MotDePasse = $mdp; Type = 'admin' })
        }
    }
    Ensure-Membre 'T0_Admins-AD' $adm
    Ensure-Membre "$sidDom-512" $adm $daNom
}

$pso = if ($script:AdOk) { Get-ADFineGrainedPasswordPolicy -Filter "Name -eq 'PSO-Admins'" } else { $null }
if ($pso) { Note-Deja 'PSO-Admins' }
else {
    [void](Faire 'PSO-Admins (20 caracteres, 90 jours, 5 essais / 60 min, priorite 10)' {
        New-ADFineGrainedPasswordPolicy -Name 'PSO-Admins' -Precedence 10 `
            -MinPasswordLength 20 -MaxPasswordAge '90.00:00:00' `
            -LockoutThreshold 5 -LockoutDuration '01:00:00' `
            -LockoutObservationWindow '01:00:00' -ComplexityEnabled $true
    })
}
foreach ($cible in 'T0_Admins-AD', 'T1_Admins-Serveurs') {
    $sujets = @()
    if ($script:AdOk -and (Get-ADFineGrainedPasswordPolicy -Filter "Name -eq 'PSO-Admins'")) {
        $sujets = @(Get-ADFineGrainedPasswordPolicySubject -Identity 'PSO-Admins' | ForEach-Object { $_.Name })
    }
    if ($sujets -contains $cible) { Note-Deja "PSO-Admins applique a $cible" }
    else { [void](Faire "PSO-Admins applique a $cible" { Add-ADFineGrainedPasswordPolicySubject -Identity 'PSO-Admins' -Subjects $cible }) }
}
# Le compte Administrateur d'origine n'est PAS renommé : la procédure le
# recommande en texte sans donner de nom ni de commande, on le laisse au formateur.
Fin-Section

# ====================================================================== 6. les 38 comptes d'ordinateur
Section '6. Comptes d ordinateur (38) et redirection des conteneurs'
$machines = @()
1..18 | ForEach-Object { $machines += @{ Nom = ('PC-COM-{0:D2}' -f $_); Dn = $dnFixes; Ou = 'Postes-Fixes' } }
1..20 | ForEach-Object { $machines += @{ Nom = ('PT-NOM-{0:D2}' -f $_); Dn = $dnPort;  Ou = 'Portables' } }
foreach ($m in $machines) {
    $existe = $script:AdOk -and (Get-ADComputer -Filter "Name -eq '$($m.Nom)'")
    if ($existe) { Note-Deja "$($m.Nom) ($($m.Ou))" 'Ordinateurs' }
    else { [void](Faire "$($m.Nom) dans $($m.Ou)" { New-ADComputer -Name $m.Nom -SamAccountName "$($m.Nom)`$" -Path $m.Dn } 'Ordinateurs') }
}
$conteneurs = if ($script:AdOk) { Get-ADDomain -Identity $Domaine } else { $null }
if ($conteneurs -and $conteneurs.ComputersContainer -eq $dnFixes) { Note-Deja "redircmp : nouveaux ordinateurs vers $dnFixes" }
else {
    [void](Faire "redircmp : nouveaux ordinateurs vers $dnFixes" {
        $sortie = & redircmp.exe $dnFixes 2>&1
        if ($LASTEXITCODE -ne 0) { throw "redircmp : $sortie" }
    })
}
if ($conteneurs -and $conteneurs.UsersContainer -eq $dnUtil) { Note-Deja "redirusr : nouveaux comptes vers $dnUtil" }
else {
    [void](Faire "redirusr : nouveaux comptes vers $dnUtil" {
        $sortie = & redirusr.exe $dnUtil 2>&1
        if ($LASTEXITCODE -ne 0) { throw "redirusr : $sortie" }
    })
}
Fin-Section

# ====================================================================== 7. export des mots de passe
Section '7. Mots de passe initiaux'
if ($script:Simu) {
    Write-Host "  [simul]   $CsvMdp ne sera ecrit qu au vrai lancement" -ForegroundColor Yellow
} elseif ($script:MotsDePasse.Count -eq 0) {
    Write-Host '  aucun compte cree pendant ce lancement : rien a exporter' -ForegroundColor DarkGray
} else {
    try {
        if (-not (Test-Path -LiteralPath $DossierLab)) { New-Item -ItemType Directory -Path $DossierLab -Force | Out-Null }
        $ajout = Test-Path -LiteralPath $CsvMdp
        $script:MotsDePasse | Export-Csv -LiteralPath $CsvMdp -NoTypeInformation -Delimiter ';' -Encoding UTF8 -Append:$ajout
        Write-Host ('  {0} mot(s) de passe ecrit(s) dans {1}' -f $script:MotsDePasse.Count, $CsvMdp) -ForegroundColor Green
    } catch {
        Note-Erreur ("export $CsvMdp : " + $_.Exception.Message)
    }
}
Fin-Section

# ====================================================================== résumé
$t = $script:Totaux
Write-Host ''
Write-Host ('=' * 72) -ForegroundColor Cyan
$prefixe = if ($script:Simu) { 'SIMULATION, apres lancement : ' } else { 'Resultat : ' }
$bilan = '{0} OU, {1} comptes, {2} groupes, {3} ordinateurs, {4} adm' -f `
    $t.OU, $t.Comptes, $t.Groupes, $t.Ordinateurs, $t.Adm
$attendu = '22 OU, 100 comptes, 22 groupes, 38 ordinateurs, 6 adm'
$couleur = if ($bilan -eq $attendu -and $script:NbErreurs -eq 0) { 'Green' } else { 'Yellow' }
Write-Host (' ' + $prefixe + $bilan) -ForegroundColor $couleur
if ($bilan -ne $attendu) { Write-Host (' attendu : ' + $attendu) -ForegroundColor Yellow }
Write-Host ' Suite a la main, dans la procedure : serveur de fichiers, partages, droits NTFS des DL_, lecteurs.' -ForegroundColor Gray
if ($script:Tp) {
    $etatG = if ($script:Tp.Groupe) { "$TpGroupe conserve" } else { "$TpGroupe absent" }
    Write-Host (' TP Finance repris : {0} compte(s) dont {1} parmi les 100, {2} (voir section 3 bis)' -f `
        $script:Tp.Comptes.Count, ($script:Tp.Comptes.Count - $tpHorsFichier.Count), $etatG) -ForegroundColor $couleur
}
if ($script:NbErreurs) { Write-Host (" $($script:NbErreurs) erreur(s) : relire les lignes ERREUR ci-dessus, puis relancer") -ForegroundColor Red }
Write-Host ('=' * 72) -ForegroundColor Cyan
if (-not $script:Simu -and (Test-Path -LiteralPath $CsvMdp)) {
    Write-Host ''
    Write-Host " ATTENTION : $CsvMdp contient des mots de passe EN CLAIR." -ForegroundColor Red
    Write-Host ' Le distribuer par un canal separe, puis le supprimer : a supprimer en fin de seance.' -ForegroundColor Red
    Write-Host "   Remove-Item $CsvMdp" -ForegroundColor Red
}
if ($script:NbErreurs) { exit 2 }
