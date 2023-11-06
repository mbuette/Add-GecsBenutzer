# Datei als Array einlesen und zeilenweise abarbeiten
# WICHTIG: CSV-Datei muss im UTF-8 Format gespeichert sein oder der Parameter -Encoding muss angegeben werden
$SkriptOrdner = Split-Path $MyInvocation.MyCommand.Path -Parent
$Datei = Import-Csv "$SkriptOrdner\GecsBenutzer.csv"
# Alternative: CSV-Datei unter Excel als "CSV UTF-8" speichern
# WICHTIG: Eyxcel verwendet standardmäßig das Semikolon zur Trennung der Spalten, der Parameter -Delimiter wird benötigt
# $Datei = Import-Csv -Delimiter ';' C:\Add-GecsBenutzer\GecsBenutzer.csv

# Zur Ereignisprotokollierung ist einmalig eine neue Quelle anzulegen
# Beispiel für eine If-Bedingung ohne eine Hilfsvariable:
If ( ( Get-EventLog -LogName "Application" -Source "Add-GecsBenutzer.ps1" -Newest 1 -ErrorAction SilentlyContinue ) -eq $null ) {
    New-EventLog -LogName "Application" -Source "Add-GecsBenutzer.ps1"
    Write-EventLog -LogName "Application" -Source "Add-GecsBenutzer.ps1"-EventId 100 -Message "Quelle angelegt"
}

# Sicherstellen, dass die ActiveDirectory CmdLets verfügbar sind
Import-Module ActiveDirectory

# Ausgabe von Meldungen auf der Kommandozeile und Eintrag in das Ergeignisprotokoll
# Auslagerung in eine Fuktion, weil der Code mehrmals benötigt wird
# Ereignis-ID ist 100 für neue Objekte und 101 für fehlgeschlagene Aktionen
Function Write-Result {
    Param ([string]$Nachricht, [int]$EreignisId)
    If ( $EreignisId -eq 100) {
        Write-Output $Nachricht
        Write-EventLog -LogName "Application" -Source "Add-GecsBenutzer.ps1" -EntryType Information -EventId $EreignisId -Message $Nachricht
    }
    Else {
        Write-Warning $Nachricht
        Write-EventLog -LogName "Application" -Source "Add-GecsBenutzer.ps1" -EntryType Warning -EventId $EreignisId -Message $Nachricht
    }
}
#
# Hauptprogramm: Die Datei zeilenweise abarbeiten
#
ForEach ( $Zeile in $Datei) {
    # Attributer der aktuellen Zeile an Variablen zuweisen
    $Vorname = $Zeile.Vorname
    $Nachname = $Zeile.Nachname
    $Standort = $Zeile.Standort
    $Abteilung = $Zeile.Abteilung
    # PowerShell akzeptiert keine Klartext-Kennwörter
    $Passwort = ConvertTo-SecureString $Zeile.Passwort -AsPlainText -Force
    # Aktuellen Benutzer über eine Funktion des .NET Framework ermitteln
    $AktuellerAdmin = [System.Environment]::UserName
    # Zusätzliche Angaben, die nicht direkt in der CSV-Datei stehen
    $Name = $Vorname+" "+$Nachname
    $GruppenName = "G_"+$Abteilung
    # Domänennamen in Kleinschrift aus den Umgebungsvariablen
    $DomainName = $env:USERDNSDOMAIN.ToLower()
    # Aus dem DNS Domänennamen (2 Stufen) den distiguished name für die Domäne ermitteln
    $Dcs = $DomainName.Split('.') # Teilt den String an der Zeichenkette in einzelne Komponenten
    $DomainDn = "DC="+$Dcs[0]+",DC="+$Dcs[1]
    # Definition der untergeordneten LDAP Distinguished Names (DNs)
    $StandortDn = "OU="+$Standort+","+$DomainDn
    $AbteilungsDn = "OU="+$Abteilung+","+$StandortDn
    $BenutzerDn = "CN="+$Name+","+$AbteilungsDn
    $GruppenDn = "OU=Gruppen,"+$DomainDn

    # Überprüfen ob OUs existieren, sonst anlegen. Erst Standort, dann die untergeordnete Abteilung
    $StandortOu = Get-AdOrganizationalUnit -Filter 'DistinguishedName -eq $StandortDn'
    $AbteilungsOu = Get-AdOrganizationalUnit -Filter 'DistinguishedName -eq $AbteilungsDn'
    If ( $StandortOu -eq $null ) {
        New-ADOrganizationalUnit -Name $Standort -Path $DomainDn -ProtectedFromAccidentalDeletion $false
        Write-Result "Eine neue Standort-OU mit dem DN $StandortDn wurde angelegt" 100
    }
    If ( $AbteilungsOu -eq $null ) {
        New-ADOrganizationalUnit -Name $Abteilung -Path $StandortDn -ProtectedFromAccidentalDeletion $false
        Write-Result "Eine neue Abteilungs-OU mit dem DN $AbteilungsDn wurde angelegt" 100
    }

    # Existieren die Gruppen-OU und die Gruppe, sonst anlegen
    $GruppenOu = Get-ADOrganizationalUnit -Filter 'DistinguishedName -eq $GruppenDn'
    If ( $GruppenOu -eq $null ) {
        New-ADOrganizationalUnit -Name "Gruppen" -Path $DomainDn -ProtectedFromAccidentalDeletion $false
        Write-Result "Eine neue Gruppen-OU mit dem DN $DomainDn wurde angelegt" 100
    }
    # -Filter statt -Identity unterdrückt die Fehlermeldung, falls die Gruppe noch nicht existiert
    $GruppenObjekt = Get-ADGroup -Filter 'SamAccountName -eq $GruppenName'
    If ( $GruppenObjekt -eq $null ) {
        # Gruppe als Globale Sicherheitsgruppe anlegen
        New-ADGroup -Name $GruppenName -Path $GruppenDn -GroupCategory Security -GroupScope Global
        Write-Result "Eine neue Gruppe mit dem Namen $GruppenName wurde angelegt" 100
    }
    
    # Anmeldenamen aus 1. Buchstabe des Vornames, Nachname und ggf. laufender Nummer generieren
    # In Kleinschreibung konvertieren und deutsche Umlaute ersetzen
    $Anmeldename = (((($Vorname.Substring(0,1)+$Nachname).ToLower()).Replace('ä','ae')).Replace('ö','oe')).Replace('ü','ue')
    $BenutzerTest = Get-ADUser -Filter 'SamAccountName -eq $Anmeldename'
    # Anmeldename existiert bereits
    If ( $BenutzerTest -ne $null ) {
        # Laufende Nummer 2++ hochzählen
        $BenutzerNummer = 2
        Do {
            $VNachnameN = $Anmeldename+$BenutzerNummer
            $BenutzerTest = Get-ADUser -Filter 'SamAccountName -eq $VNachnameN'
            $BenutzerNummer++
        }
        Until ( $BenutzerTest -eq $null )
        $Anmeldename = $VNachnameN
    }

    # Kann erst jetzt definiert werden
    $Upn = $Anmeldename+"@"+$DomainName
    # Überprüfen ob das Benutzerobjekt existiert, sonst anlegen
    $BenutzerObjekt = Get-ADUser -Filter 'SamAccountName -eq $Anmeldename'
    $BenutzerDnObjekt = Get-ADUser -Filter 'DistinguishedName -eq $BenutzerDn'
    If ( ($Benutzerobjekt -eq $null) -and ($BenutzerDnObjekt -eq $null ) ) {
        New-ADUser -Name $Name -DisplayName $Name -GivenName $Vorname -Surname $Nachname `
            -SamAccountName $Anmeldename -UserPrincipalName $Upn -Path $AbteilungsDn -AccountPassword $Passwort `
            -Enabled $true 2> $null
        # Rückgabewert der 1. Funktion merken
        $FunktionsErgebnis = $?
        # Benutzer zur Abteilungsgruppe hinzufügen
        Add-ADGroupMember -Identity $GruppenName -Members $Anmeldename 2> $null
        # Beide Ergebnisse mit UND verknüpfen
        $FunktionsErgebnis = $FunktionsErgebnis -and $?
        If ( $FunktionsErgebnis ) {
            Write-Result "Ein neuer Benutzer mit dem Namen $Vorname $Nachname ($Anmeldename) wurde angelegt und zur Gruppe $GruppenName hinzugefügt" 100
        }
        Else {
            Write-Result "Beim Anlegen des Benutzers $Vorname $Nachname ($Anmeldename) ist ein Fehler aufgetreten" 101
        }
    }
    Else {
        Write-Result "Der Benutzer mit dem Benutzer-Anmeldenamen $Anmeldename und DN $BenutzerDn existiert bereits" 101
    }
}