<#Lecteur audio compatible aac, flac, m4a, mp3, wav, wma
Version 1.3 Beta

Author: Jean-Baptiste CHARRON
#>

Set-StrictMode -Version Latest

Add-Type -AssemblyName presentationCore
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

#Initialize variables
[Object]$MusicPath = @()
$script:Playlist = New-Object System.Collections.ArrayList
[Int32]$script:Index = 0
[bool]$script:Files = 0
[bool]$script:Statut = 0
[Object]$Transfer = @()
[Int32]$script:EndedPlaylist = 0

function Add-File {#Import file
        $FormFile = New-Object System.Windows.Forms.OpenFileDialog
        $FormFile.CheckFileExists = $true
        $FormFile.CheckPathExists = $true
        $FormFile.Filter = "All Audio Files (*.*)|*.aac;*.flac;*.m4a;*.mp3;*.wav;*.wma|Free Lossless Audio Codec (*.flac)|*.flac|MPEG 4 Audio(*.m4a)|*.m4a|MPEG Audio (*.mp3)|*.mp3|Raw AAC (*.aac)|*.aac|Waveform Audio File Format (*.wav)|*.wav|Windows Media Audio (*.wma)|*.wma"
        $FormFile.InitialDirectory = "C:\Users\$env:USERNAME\Music\"
        $FormFile.Multiselect = $true
        $FormFile.Title = 'Sélectionner un fichier audio à ouvrir'
        if ($FormFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $MusicPath = $FormFile.FileNames
            Invoke-Playlist -Path $MusicPath
        }
}

function Add-Folder {#Importer directory
    $FormFolder = New-Object System.Windows.Forms.FolderBrowserDialog
    $FormFolder.RootFolder = 'MyComputer'
    if ($FormFolder.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $MusicPath = $FormFolder.SelectedPath
        Invoke-Playlist -Path $MusicPath
    }
}

function Open-Path {#Open path where read file is save
    param (
        $File
        )
    if (Test-Path $File)
    {
        $Argument = '/select,' + $File
        Start-Process explorer.exe -ArgumentList $Argument
    }
    else
    {
        [System.Windows.MessageBox]::Show("Impossible d'ouvrir l'emplacement du fichier",'Erreur',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Clear-Playlist {#Reset player with statut message
    param (
        $Message
        )
    $MediaPlayer.Close()
    $Timer.Stop()
    $Timer.Enabled = $false
    $script:Statut = 0
    $script:Playlist.Clear()
    $script:Index = 0
    $script:Files = 0
    $OpenPath.Enabled = $false
    $TrackTitle.Text = 'Aucune piste en cours'
    $TrackDuration.Location = New-Object System.Drawing.Size(158, 68)
    $TrackDuration.Text = '00:00'
    $ButtonPlay.Text = $StartPlayback.Text = 'Play'
    $ButtonPrevious.Enabled = $ButtonNext.Enabled = $PreviousPlayback.Enabled = $NextPlayback.Enabled = $Random.Enabled = $true
    $StatusLabel.Text = $Message
    $Transfer = @()
}

function Invoke-About {#Generate About window
    $MainAbout = New-Object System.Windows.Forms.Form
    $MainAbout.ControlBox = $false
    $MainAbout.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $MainAbout.Height = 160
    $MainAbout.MaximizeBox = $false
    $MainAbout.MinimizeBox = $false
    $MainAbout.ShowIcon = $false
    $MainAbout.ShowInTaskbar = $false
    $MainAbout.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $MainAbout.Text = 'À propos'
    $MainAbout.Width = 360

    $AboutTitle = New-Object System.Windows.Forms.Label
    $AboutTitle.AutoSize = $true
	$AboutTitle.Font = New-Object Drawing.Font('SegoeUI', 10, [System.Drawing.FontStyle]::Bold)
    $AboutTitle.Text = 'Lecteur Audio v1.3'
    $AboutTitle.Location = New-Object System.Drawing.Size(110, 18)
    $MainAbout.Controls.Add($AboutTitle)

    $Developer = New-Object System.Windows.Forms.Label
    $Developer.AutoSize = $true
    $Developer.Font = New-Object Drawing.Font('SegoeUI', 9)
    $Developer.Text = 'Développé par Jean-Baptiste CHARRON'
    $Developer.Location = New-Object System.Drawing.Size(65, 45)
    $MainAbout.Controls.Add($Developer)

    $AboutExit = New-Object System.Windows.Forms.Button
    $AboutExit.Location = New-Object System.Drawing.Size(135, 90)
    $AboutExit.Text = 'OK'
    $MainAbout.CancelButton = $AboutExit
    $MainAbout.Controls.Add($AboutExit)

    [void]$MainAbout.ShowDialog()
    $MainAbout.Dispose()
}

function Invoke-Playlist {#function to generate playlist
    param (
        $Path
        )
    $Elements = ForEach-Object -InputObject $Path{(Get-ChildItem -Path $_ -Include *.aac,*.flac,*.m4a,*.mp3,*.wav,*.wma -Recurse -File | Measure-Object).Count}
    if ($Elements -gt 0)#Detect if path contains audio files
    {
        $Path | ForEach-Object{(Get-ChildItem -Path $_ -Include *.aac,*.flac,*.m4a,*.mp3,*.wav,*.wma -Recurse -File).FullName} | Sort-Object | ForEach-Object{
        $script:Playlist.Add($_)
        }
        if ($script:Playlist.Count -gt 1)#Detect files number
        {
            $ButtonPrevious.Enabled = $ButtonNext.Enabled = $PreviousPlayback.Enabled = $NextPlayback.Enabled = $Random.Enabled = $true
        }
        else
        {
            $ButtonPrevious.Enabled = $ButtonNext.Enabled = $PreviousPlayback.Enabled = $NextPlayback.Enabled = $Random.Enabled = $false
        }
        if ($script:Files -eq 0)#Detect if player have already files in playlist
        {
            $script:Files = 1
            $OpenPath.Enabled = $true
            Read-Music
        }
        else
        {
            Get-MetaData -Path $script:Playlist[$script:Index]
        }
    }
    else
    {
        $StatusLabel.Text = "Aucune piste audio n'a été trouvé"
    }
}

function Read-Music {#function to read audio file
    $MediaPlayer.Open($script:Playlist[$script:Index])
    $MediaPlayer.Play()
    $script:Statut = 1
    $ButtonPlay.Text = $StartPlayback.Text = 'Pause'
    Get-MetaData -Path $script:Playlist[$script:Index]
    Get-Duration -Path $script:Playlist[$script:Index]
}

function Get-MetaData {#Show tags
    param (
        $Path
    )
    $Shell = New-Object -ComObject Shell.Application
    $ShellFolder = $Shell.Namespace($(Split-Path -Path $Path))
    $ShellFile = $ShellFolder.ParseName($(Split-Path -Path $Path -Leaf))
    if ($ShellFolder.GetDetailsOf($ShellFile, 21))#Detect if tag 'Title' exists on played file
    {
        $TrackTitle.Text = $ShellFolder.GetDetailsOf($ShellFile, 21) + ' - ' + $ShellFolder.GetDetailsOf($ShellFile, 14)
    }
    else
    {
        $TrackTitle.Text = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    }
    if ($script:Index -lt ($script:Playlist.Count)-1)#Detect if played file is the latest
    {
        $ShellFolder = $Shell.Namespace($(Split-Path -Path $script:Playlist[$script:Index + 1]))
        $ShellFile = $ShellFolder.ParseName($(Split-Path -Path $script:Playlist[$script:Index + 1] -Leaf))
        if ($ShellFolder.GetDetailsOf($ShellFile, 21))#Detect if tag 'Title' exists on next played file
        {
            $StatusLabel.Text = 'Prochaine piste: ' + $ShellFolder.GetDetailsOf($ShellFile, 26) + ' - ' + $ShellFolder.GetDetailsOf($ShellFile, 21)
        }
        else
        {
            $StatusLabel.Text = 'Prochaine piste: ' + [System.IO.Path]::GetFileNameWithoutExtension($script:Playlist[$script:Index + 1])
        }
    }
    else
    {
        if ($script:EndedPlaylist -eq 0)
        {
            $StatusLabel.Text = 'Fin de la playlist'
        }
        else
        {
            $StatusLabel.Text = 'Le lecteur va se fermer à la fin de cette piste'
        }
    }
}

function Get-Duration {#Show track duration and timer
    param (
        $Path
    )
    $Timer.Enabled = $true
    $Shell = New-Object -COMObject Shell.Application
    $ShellFolder = $Shell.Namespace($(Split-Path -Path $Path))
    $ShellFile = $ShellFolder.ParseName($(Split-Path -Path $Path -Leaf))
    if ($ShellFolder.GetDetailsOf($ShellFile, 27) -ne '' -and [System.IO.Path]::GetExtension($Path) -ne '.wma')#Detect if file have duration or file have extension 'wma'
    {
        $script:Duration = $ShellFolder.GetDetailsOf($ShellFile, 27)
        $Timer.Add_Tick({$TrackDuration.Text = "$($MediaPlayer.Position.Minutes.ToString('00')):$($MediaPlayer.Position.Seconds.ToString('00')) / " + $script:Duration.SubString($script:Duration.Length - 5)})
        $TrackDuration.Location = New-Object System.Drawing.Size(143, 68)
    }
    else
    {
        $Timer.Add_Tick({$TrackDuration.Text = "$($MediaPlayer.Position.Minutes.ToString('00')):$($MediaPlayer.Position.Seconds.ToString('00'))"})
        $TrackDuration.Location = New-Object System.Drawing.Size(158, 68)
    }
    $Timer.Start()
}

function Get-StatusPlay {#manage button play/pause
    if ($script:Playlist.Count -ge 1)#Detect if file is in the playlist
    {
        if ($ButtonPlay.Text -eq 'Play')#Detect button statut
        {
            $MediaPlayer.Play()
            $Timer.Start()
            $script:Statut = 1
            $ButtonPlay.Text = $StartPlayback.Text = 'Pause'
        }
        else
        {
            $MediaPlayer.Pause()
            $Timer.Stop()
            $script:Statut = 0
            $ButtonPlay.Text = $StartPlayback.Text = 'Play'
        }
    }
    else
    {
        Add-File
    }
}

function Open-Next {#Prepare next track
    $Timer.Stop()
        if ($script:Index -lt ($script:Playlist.Count) - 1)
        {
            $script:Index += 1
            $MediaPlayer.Close()
            Read-Music
        }
}

function Open-Previous {#Prepare previous track
    if ($script:Index -gt 0)
    {
        $script:Index -= 1
        $MediaPlayer.Close()
        Read-Music
    }
}

$PlayerGUI = New-Object System.Windows.Forms.Form #Main Window
$PlayerGUI.AllowDrop = $true
$PlayerGUI.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$PlayerGUI.Height = 190
$PlayerGUI.MaximizeBox = $false
$PlayerGUI.ShowIcon = $false
$PlayerGUI.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$PlayerGUI.Text = 'Lecteur audio v1.3 Beta'
$PlayerGUI.Width = 370

$PlayerGUI_DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
	{
	    $_.Effect = 'Copy'
	}
    else
	{
	    $_.Effect = 'None'
	}
}
$PlayerGUI_DragDrop = [System.Windows.Forms.DragEventHandler]{
        $MusicPath=@()
	    ForEach ($FileName in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
        {
            $MusicPath += $FileName
	    }
        Invoke-Playlist -Path $MusicPath
}

$PlayerGUI.Add_DragOver($PlayerGUI_DragOver)
$PlayerGUI.Add_DragDrop($PlayerGUI_DragDrop)
$PlayerGUI.Add_Closing({$MediaPlayer.Close()
Clear-Playlist})

$TrackTitle = New-Object System.Windows.Forms.Label
$TrackTitle.AutoSize = $false
$TrackTitle.Dock = [System.Windows.Forms.DockStyle]::Top
$TrackTitle.Font = New-Object System.Drawing.Font('SegoeUI', 9)
$TrackTitle.Height = 40
$TrackTitle.Text = 'Aucune piste en cours'
$TrackTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$PlayerGUI.Controls.Add($TrackTitle)

$MainMenuStrip = New-Object System.Windows.Forms.MenuStrip
$PlayerGUI.Controls.Add($MainMenuStrip)

$FileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$FileMenu.Text = '&Fichier'
[void]$MainMenuStrip.Items.Add($FileMenu)

$OpenFile = New-Object System.Windows.Forms.ToolStripMenuItem
$OpenFile.ShortcutKeys = 'Control, O'
$OpenFile.Text = 'Ouvrir un fichier'
$OpenFile.Add_Click({Add-File})
[void]$FileMenu.DropDownItems.Add($OpenFile)

$OpenDir = New-Object System.Windows.Forms.ToolStripMenuItem
$OpenDir.ShortcutKeys = 'Control, F'
$OpenDir.Text = 'Ouvrir un dossier'
$OpenDir.Add_Click({Add-Folder})
[void]$FileMenu.DropDownItems.Add($OpenDir)

$OpenPath = New-Object System.Windows.Forms.ToolStripMenuItem
$OpenPath.Text = "Ouvrir l'emplacement"
$OpenPath.Enabled = $false
$OpenPath.Add_Click({
    if ($script:Playlist.Count -cge 1)
    {
        Open-Path -File $script:Playlist[$script:Index]
    }
})
[void]$FileMenu.DropDownItems.Add($OpenPath)

$QuitPlayer = New-Object System.Windows.Forms.ToolStripMenuItem
$QuitPlayer.ShortcutKeys = 'Control, Q'
$QuitPlayer.Text = 'Quitter'
$QuitPlayer.Add_Click({$PlayerGUI.Close()})
[void]$FileMenu.DropDownItems.Add($QuitPlayer)

$PlaybackMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$PlaybackMenu.Text = '&Lecture'
[void]$MainMenuStrip.Items.Add($PlaybackMenu)

$StartPlayback = New-Object System.Windows.Forms.ToolStripMenuItem
$StartPlayback.ShortcutKeys = 'Control, Space'
$StartPlayback.Text = 'Play'
$StartPlayback.Add_Click({Get-StatusPlay})
[void]$PlaybackMenu.DropDownItems.Add($StartPlayback)

$PreviousPlayback = New-Object System.Windows.Forms.ToolStripMenuItem
$PreviousPlayback.ShortcutKeys = 'Control, Left'
$PreviousPlayback.Text = 'Précédent'
$PreviousPlayback.Add_Click({Open-Previous})
[void]$PlaybackMenu.DropDownItems.Add($PreviousPlayback)

$NextPlayback = New-Object System.Windows.Forms.ToolStripMenuItem
$NextPlayback.ShortcutKeys = 'Control, Right'
$NextPlayback.Text = 'Suivant'
$NextPlayback.Add_Click({Open-Next})
[void]$PlaybackMenu.DropDownItems.Add($NextPlayback)

$Random = New-Object System.Windows.Forms.ToolStripMenuItem
$Random.Text = 'Aléatoire'
$Random.Add_Click({
    if ($script:Playlist.Count -ge 2)
    {
        $MediaPlayer.Close()
        $Transfer = $script:Playlist | Sort-Object{Get-Random}
        $script:Playlist.Clear()
        $Transfer | ForEach-Object{$script:Playlist.Add($_)}
        $Transfer.Clear()
        $script:Index = 0
        Read-Music
    }
})
[void]$PlaybackMenu.DropDownItems.Add($Random)

$CleanPlaylist = New-Object System.Windows.Forms.ToolStripMenuItem
$CleanPlaylist.Text = "Vider la file d'attente"
$CleanPlaylist.Add_Click({
    Clear-Playlist -Message "File d'attente vidé"
})
[void]$PlaybackMenu.DropDownItems.Add($CleanPlaylist)

$OutPlayer = New-Object System.Windows.Forms.ToolStripMenuItem
$OutPlayer.Text = 'Quitter à la fin'
$OutPlayer.Checked = $false
$OutPlayer.Add_Click({
    if ($OutPlayer.Checked -eq $false)
    {
        $script:EndedPlaylist = 1
        $OutPlayer.Checked = $true
    }
    else
    {
        $script:EndedPlaylist = 0
        $OutPlayer.Checked = $false
    }
    if ($script:Statut -eq 1)
    {
        Get-MetaData -Path $script:Playlist[$script:Index]
    }
})
[void]$PlaybackMenu.DropDownItems.Add($OutPlayer)

$HelpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$HelpMenu.Text = '&Aide'
[void]$MainMenuStrip.Items.Add($HelpMenu)

$OpenAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$OpenAbout.ShortcutKeys = 'F1'
$OpenAbout.Text = 'À propos'
$OpenAbout.Add_Click({Invoke-About})
[void]$HelpMenu.DropDownItems.Add($OpenAbout)

$StatusStrip = New-Object System.Windows.Forms.StatusStrip
$StatusStrip.SizingGrip = $false
$PlayerGUI.Controls.Add($StatusStrip)

$StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$StatusLabel.AutoSize = $true
$StatusLabel.Text = 'Prêt'
[void]$StatusStrip.Items.Add($StatusLabel)

$TrackDuration = New-Object System.Windows.Forms.Label
$TrackDuration.Anchor = [System.Windows.Forms.AnchorStyles]::None
$TrackDuration.AutoSize = $true
$TrackDuration.Font = New-Object System.Drawing.Font('SegoeUI', 9)
$TrackDuration.Location = New-Object System.Drawing.Size(158, 68)
$TrackDuration.Text = '00:00'
$PlayerGUI.Controls.Add($TrackDuration)

$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 500

$ButtonPrevious = New-Object System.Windows.Forms.Button
$ButtonPrevious.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$ButtonPrevious.AutoSize = $true
$ButtonPrevious.Location = New-Object System.Drawing.Size(58, 97)
$ButtonPrevious.Text = 'Précédent'
$ButtonPrevious.Add_Click({Open-Previous})
$PlayerGUI.Controls.Add($ButtonPrevious)

$ButtonPlay = New-Object System.Windows.Forms.Button
$ButtonPlay.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$ButtonPlay.AutoSize = $true
$ButtonPlay.Location = New-Object System.Drawing.Size(138, 97)
$ButtonPlay.Text = 'Play'
$ButtonPlay.Add_Click({Get-StatusPlay})
$PlayerGUI.Controls.Add($ButtonPlay)

$ButtonNext = New-Object System.Windows.Forms.Button
$ButtonNext.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$ButtonNext.AutoSize = $true
$ButtonNext.Location = New-Object System.Drawing.Size(218, 97)
$ButtonNext.Text = 'Suivant'
$ButtonNext.Add_Click({Open-Next})
$PlayerGUI.Controls.Add($ButtonNext)

$MediaPlayer = New-Object System.Windows.Media.MediaPlayer
$MediaPlayer.Volume = 1
$handler_MediaPlayer_MediaEnded=#Detect ended file and read next file
{
    $Timer.Stop()
    if ($script:EndedPlaylist -eq 1 -and $script:Index -eq ($script:Playlist.Count) - 1)
    {
        $PlayerGUI.Close()
    }
    if ($script:Index -le ($script:Playlist.Count))
    {
        Open-Next
    }
}
$MediaPlayer.add_MediaEnded($handler_MediaPlayer_MediaEnded)

$handler_MediaPlayer_MediaFailed=#Detect if file cannot be read
{
    Clear-Playlist -Message 'Erreur de lecture'
    [System.Windows.Forms.MessageBox]::Show('Impossible de lire le fichier audio (format non supporté ou fichier corrompu)','Erreur de lecture',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
}
$MediaPlayer.add_MediaFailed($handler_MediaPlayer_MediaFailed)

$PlayerGUI.Add_Shown({$PlayerGUI.Activate()})
$PlayerGUI.Add_Shown({$ButtonPlay.Select()})
[void]$PlayerGUI.ShowDialog()
$PlayerGUI.Dispose()