#Script qui recherche à l'emplacement spécifié les fichiers qui ont été modifiés il y a un jour ou moins.
#Version 2.6

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

Add-Type -AssemblyName System.Windows.Forms
[Windows.Forms.Application]::EnableVisualStyles()

$NbDay = -1
Write-Output "Ce script affiche les derniers fichiers modifiés à l'emplacement spécifié."

$FormsDir = New-Object System.Windows.Forms.FolderBrowserDialog
$FormsDir.RootFolder = "MyComputer"
if ($FormsDir.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Directory = $FormsDir.SelectedPath
    if ((Get-ChildItem -Path $Directory -Recurse -Force -File | Where-Object {$_.LastWriteTime -ge (Get-Date).AddDays($NbDay)} | Measure-Object).Count -eq 0)
    {
        [System.Windows.Forms.MessageBox]::Show("Aucun résultat","Information",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Start-Sleep 1
        exit
    }
    else
    {
        $Title_GridView = 'Derniers fichiers modifiés dans ' + [System.IO.Path]::GetFileName($Directory)
        Get-ChildItem -Path $Directory -Recurse -Force -File | Where-Object {$_.LastWriteTime -ge (Get-Date).AddDays($NbDay)} | Select-Object @{L='Emplacement du fichier';E={$_.FullName}},@{L='Modifié le';E={$_.LastWriteTime}} | Out-GridView -Title $Title_GridView -Wait
    }
        if ([System.Windows.Forms.MessageBox]::Show("Voulez-vous ouvrir l'emplacement spécifié ?", "Ouvrir l'emplacement spécifié",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq [System.Windows.Forms.DialogResult]::Yes)
        {
            Start-Process $Directory
        }
        else{
            exit
        }
}