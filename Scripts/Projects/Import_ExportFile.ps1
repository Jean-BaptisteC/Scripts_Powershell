﻿#Import-ExportFile 1.2

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

$Dir = 'C:'

$FormFile = New-Object System.Windows.Forms.OpenFileDialog
$FormFile.CheckFileExists = $true
$FormFile.CheckPathExists = $true
$FormFile.Filter = 'Fichier (*.csv, *.txt)|*.csv;*.txt|CSV|*.csv|Fichier texte|*.txt'
$FormFile.InitialDirectory = "C:\Users\$env:USERNAME\Desktop\"
$FormFile.Multiselect = $false
$FormFile.Title = 'Sélectionner un fichier à exporter'

$SavePath = New-Object System.Windows.Forms.SaveFileDialog
$SavePath.CheckPathExists = $true
$SavePath.RestoreDirectory = $false

$script:balloon = New-Object System.Windows.Forms.NotifyIcon
Register-ObjectEvent -InputObject $script:balloon -EventName BalloonTipClicked -Action{Start-Process $script:Output} | Out-Null

function Open-File {
    if($FormFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $File = $FormFile.Filename
        Get-Extension
    }
    else
    {
        exit
    }
}

function Get-Extension {
    $Dir = [System.IO.Path]::GetDirectoryName($File)
    if([System.IO.Path]::GetExtension($File) -eq '.txt')
    {
        $Nom = [System.IO.Path]::GetFileNameWithoutExtension($File) + '.csv'
        $Extension = 'txt'
    }
    else
    {
        $Nom = [System.IO.Path]::GetFileNameWithoutExtension($File) + '.txt'
        $Extension = 'csv'
    }
    $Check = "$Dir\$Nom"
    if(Test-Path $Check)
    {
        if([System.Windows.Forms.MessageBox]::Show('Attention ce fichier existe déja voulez vous écraser son contenu.','Modification de fichier',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::Yes)
        {
            $script:Output = New-Item -Path $Dir -Name $Nom -ItemType File -Force
            Export-File
            }
            else
            {
                $SavePath.InitialDirectory = $Dir
                if($SavePath.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
                {
                    $script:Output = $SavePath.FileName
                    Export-File
                }
                else
                {
                    exit
                }
            }
    }
    else
    {
        $script:Output = New-Item -Path $Dir -Name $Nom -ItemType File
        Export-File
    }
}

function Export-File{
    param(
        $extension
    )
    if($extension -eq 'txt')
    {
        $script:Output = New-Item -Path $Dir -Name $Nom -ItemType File -Force
        Save-Txt -File $File
    }
    else
    {
        $script:Output = New-Item -Path $Dir -Name $Nom -ItemType File -Force
        Save-Csv -File $File
    }
}

function Save-Txt {#Export txt
    param(
        $File
    )
    Import-Csv -Path $File -Encoding Default | Export-Csv -Path $script:Output -Encoding UTF8 -Delimiter ',' -NoTypeInformation
    (Get-Content -Path $script:Output -Raw).replace('"','') | Set-Content -Path $script:Output
    $Dir = [System.IO.Path]::GetDirectoryName($script:Output)
    Show-Notification
}

function Save-Csv {#Export csv
        param(
        $File
    )
    Import-Csv -Path $File -Encoding Default | Export-Csv -Path $script:Output -Encoding UTF8 -Delimiter ',' -NoTypeInformation
    (Get-Content -Path $script:Output -Raw).replace('"','') | Set-Content -Path $script:Output
    $Dir = [System.IO.Path]::GetDirectoryName($script:Output)
    Show-Notification
}

function Show-Notification {
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\explorer.exe")
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $balloon.BalloonTipText = "Cliquez ici pour ouvrir"
    $balloon.BalloonTipTitle = "Fichier Exporté"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(5000)
}
Open-File
Start-Sleep -Seconds 1