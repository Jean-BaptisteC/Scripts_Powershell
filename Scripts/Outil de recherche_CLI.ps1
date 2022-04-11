#Script qui recherche à l'emplacement spécifié les fichiers qui ont été modifiés il y a un jour ou moins.
#Version 1.6

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

$NbDay = -1
Write-Output "Ce script affiche les derniers fichiers modifiés dans le dossier spécifié." `n
$Directory = Read-Host "Chemin d'accès (sans guillemets)"

if (Test-Path $Directory)
{
    if ((Get-ChildItem -Path $Directory -Recurse -Force -File | Where-Object {$_.LastWriteTime -ge (Get-Date).AddDays($NbDay)} | Measure-Object).Count -eq 0)
    {
        Write-Output "Aucun résultat"
        Start-Sleep 1
        exit
    }
    else
    {
        Get-ChildItem -Path $Directory -Recurse -Force -File | Where-Object {$_.LastWriteTime -ge (Get-Date).AddDays($NbDay)} | Format-Table -AutoSize -Property @{L='Emplacement du fichier';E={$_.Fullname}},@{L='Modifié le';E={$_.LastWriteTime}}
        $OpenDir = Read-Host "Voulez-vous ouvrir l'emplacement spécifié ?(y/n) "
        if ($OpenDir -eq "y")
        {
            Start-Process -FilePath $Directory
        }
        else
        {
            exit
        }
    }
}
else
{
    Write-Output "Le chemin d'accès spécifié est introuvable."
    Start-Sleep 1
    exit
}