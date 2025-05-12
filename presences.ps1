[CmdletBinding()]
Param(
    [Parameter(ValueFromPipeline)]
    [ValidateScript({Test-Path -Path $_ -PathType Container})]
    [string] $Path = (Join-Path -Path $Home -ChildPath "Desktop"),
    [string] $Cours,
    [string] $Groupe
)

if(-not (Get-Module -Name AdsiPS -ListAvailable)) {
    Install-Module -Name AdsiPS -Scope CurrentUser -Force
}

$InputFile = New-TemporaryFile

$arguments = @(
    "/process pcinssui.exe",
    "/type listview",
    "/visible Yes",
    "/scomma $($InputFile.FullName)"
)

Start-Process -FilePath ".\sysexp.exe" -ArgumentList $arguments -Wait

Import-Csv -Path $InputFile.FullName -Encoding Default | ForEach-Object {
    if ($sam = $_."Nom d'utilisateur") {
        [PSCustomObject]@{
            Username = ($_."Nom d'utilisateur")
            NomÉtudiant = (Get-ADSIUser -Identity $sam).DisplayName
            Ordinateur = $_.Nom
            DateHeure = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Programme = (Get-ADSIUser -Identity $sam).description
        }
    }
} | Sort-Object -Property NomÉtudiant | Tee-Object -Variable "Rapport" | Format-Table

Write-Host "Nombre de sessions:", $Rapport.count -ForegroundColor "Yellow"

$Rapport | Export-Csv -Path "$Path\Rapport_Présences_$(Get-Date -Format "yyyyMMdd_HHmm").csv" -Encoding UTF8 -UseCulture -NoTypeInformation

$InputFile | Remove-Item -Force
