<#Lecteur audio compatible aac, flac, m4a, mp3, wav, wma
Version 1.3 Beta

Author: Jean-Baptiste
#>

Set-StrictMode -Version Latest

Add-Type -AssemblyName presentationCore
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

#Initialize variables
[array]$MusicPath = @()
$script:Playlist = [System.Collections.ArrayList]@()
[Int32]$script:Index = 0
[bool]$script:Files = 0
$Transfer = [System.Collections.ArrayList]@()
[Int32]$script:EndedPlaylist = 0

function Add-File {#Import files
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
    $FormFile.Dispose()
}

function Add-Folder {#Import directory
    $FormFolder = New-Object System.Windows.Forms.FolderBrowserDialog
    $FormFolder.RootFolder = 'MyComputer'
    if ($FormFolder.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $MusicPath = $FormFolder.SelectedPath
        Invoke-Playlist -Path $MusicPath
    }
    $FormFolder.Dispose()
}

function Open-Path {#Open file path
    param (
        $File
    )
    $Argument = '/select,' + $File
    Start-Process explorer.exe -ArgumentList $Argument
}

function Clear-Playlist {#Reset player with statut message
    param (
        $Message
    )
    $MediaPlayer.Close()
    $Timer.Stop()
    $Timer.Enabled = $false
    $script:Playlist.Clear()
    $script:Index = 0
    $script:Files = 0
    $OpenPath.IsEnabled = $false
    $TrackTitle.Content = 'Aucune piste en cours'
    $TrackDuration.Content = '00:00'
    $ButtonPlay.Content = $StartPlayback.Header = 'Play'
    $ButtonPrevious.IsEnabled = $ButtonNext.IsEnabled = $PreviousPlayback.IsEnabled = $NextPlayback.IsEnabled = $Random.IsEnabled = $true
    $StatusLabel.Text = $Message
    $Transfer.Clear()
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
    $AboutTitle.Location = New-Object System.Drawing.Size(110, 18)
    $AboutTitle.Text = 'Lecteur Audio v1.4'
    $MainAbout.Controls.Add($AboutTitle)

    $Developer = New-Object System.Windows.Forms.Label
    $Developer.AutoSize = $true
    $Developer.Font = New-Object Drawing.Font('SegoeUI', 9)
    $Developer.Location = New-Object System.Drawing.Size(65, 55)
    $Developer.Text = 'Développé par Jean-Baptiste CHARRON'
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
        $Path | ForEach-Object {(Get-ChildItem -Path $_ -Include *.aac,*.flac,*.m4a,*.mp3,*.wav,*.wma -Recurse -File).FullName} | Sort-Object | ForEach-Object {
        $script:Playlist.Add($_)
        }
        if ($script:Playlist.Count -gt 1)#Detect files number
        {
            $ButtonPrevious.IsEnabled = $ButtonNext.IsEnabled = $PreviousPlayback.IsEnabled = $NextPlayback.IsEnabled = $Random.IsEnabled = $true
        }
        else
        {
            $ButtonPrevious.IsEnabled = $ButtonNext.IsEnabled = $PreviousPlayback.IsEnabled = $NextPlayback.IsEnabled = $Random.IsEnabled = $false
        }
        if ($script:Files -eq 0)#Detect if player have already files in playlist
        {
            $script:Files = 1
            $OpenPath.IsEnabled = $true
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
    $ButtonPlay.Content = $StartPlayback.Header = 'Pause'
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
        $TrackTitle.Content = $ShellFolder.GetDetailsOf($ShellFile, 21) + ' - ' + $ShellFolder.GetDetailsOf($ShellFile, 14)
    }
    else
    {
        $TrackTitle.Content = [System.IO.Path]::GetFileNameWithoutExtension($Path)
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
        $Timer.Add_Tick({$TrackDuration.Content = "$($MediaPlayer.Position.Minutes.ToString('00')):$($MediaPlayer.Position.Seconds.ToString('00')) / " + $script:Duration.SubString($script:Duration.Length - 5)})
    }
    else
    {
        $Timer.Add_Tick({$TrackDuration.Content = "$($MediaPlayer.Position.Minutes.ToString('00')):$($MediaPlayer.Position.Seconds.ToString('00'))"})
    }
    $Timer.Start()
}

function Get-StatusPlay {#manage button play/pause
    if ($script:Playlist.Count -ge 1)#Detect if file is in the playlist
    {
        if ($ButtonPlay.Content -eq 'Play')#Detect button statut
        {
            $MediaPlayer.Play()
            $Timer.Start()
            $ButtonPlay.Content = $StartPlayback.Header = 'Pause'
        }
        else
        {
            $MediaPlayer.Pause()
            $Timer.Stop()
            $ButtonPlay.Content = $StartPlayback.Header = 'Play'
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

[xml]$XML= @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Names"
    AllowDrop="true" ResizeMode="CanMinimize" Height="190" Title="Lecteur audio v1.4 Beta" Width="370" WindowStartupLocation="CenterScreen">
    <Window.Resources>
    </Window.Resources>
    <StackPanel>
    <Menu>
        <MenuItem Header="_Fichier">
            <MenuItem Name="OpenFile" Header="_Ouvrir un fichier" InputGestureText="Ctrl+O"/>
            <MenuItem Name="OpenDir" Header="_Ouvrir un dossier" InputGestureText="Ctrl+F"/>
            <MenuItem Name="OpenPath" Header="_Ouvrir l'emplacement" IsEnabled="False"/>
            <MenuItem Name="QuitPlayer" Header="_Quitter" InputGestureText="Ctrl+Q"/>
        </MenuItem>
        <MenuItem Header="_Lecture">
            <MenuItem Name="StartPlayback" Header="_Play" InputGestureText="Ctrl+Space"/>
            <MenuItem Name="PreviousPlayback" Header="_Précédent" InputGestureText="Ctrl+Left"/>
            <MenuItem Name="NextPlayback" Header="_Suivant" InputGestureText="Ctrl+Right"/>
            <MenuItem Name="Random" Header="_Aléatoire"/>
            <MenuItem Name="CleanPlaylist" Header="_Vider la file d'attente"/>
            <MenuItem Name="OutPlayer" Header="_Quitter à la fin" IsCheckable="True" IsChecked="False"/>
        </MenuItem>
        <MenuItem Header="_Aide">
            <MenuItem Name="OpenAbout" Header="_À propos" InputGestureText="F1"/>
        </MenuItem>
    </Menu>
        <DockPanel>
            <Label Name="TrackTitle" DockPanel.Dock="Top" Height="40" HorizontalAlignment ="Center" VerticalContentAlignment="Center" Content="Aucune piste en cours" FontFamily="Segoe UI" FontSize="13"/>
        </DockPanel>
        <Label Name="TrackDuration" Height="30" HorizontalAlignment ="Center" VerticalContentAlignment="Center" Content="00:00" FontFamily="Segoe UI" FontSize="13"/>
        <StackPanel Orientation="Horizontal">
        <Button Name="ButtonPrevious" HorizontalAlignment ="Left" VerticalAlignment="Bottom" Width="75" Height="24">Précédent</Button>
        <Button Name="ButtonPlay" HorizontalAlignment ="Center" VerticalAlignment="Center" Width="75" Height="24">Play</Button>
        <Button Name="ButtonNext" HorizontalAlignment ="Center" VerticalAlignment="Center" Width="75" Height="24">Suivant</Button>
        </StackPanel>
                <StatusBar>
        <StatusBarItem>
			<TextBlock Name="StatusLabel" Text="Prêt"/>
	    </StatusBarItem>
        </StatusBar>
        </StackPanel>
    </Window>
"@

$FormXML = (New-Object System.Xml.XmlNodeReader $XML)
$Player = [Windows.Markup.XamlReader]::Load($FormXML)
<#$Player.Add_PreviewDragOver({
	if ($_.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop))
	{
	    $_.Effect = 'Copy'
	}
    else
	{
	    $_.Effect = 'None'
	}
})#>
$Player.Add_Drop({
        $MusicPath=@()
	    ForEach ($FileName in $_.Data.GetData([System.Windows.DataFormats]::FileDrop))
        {
            $MusicPath += $FileName
	    }
        Invoke-Playlist -Path $MusicPath
})

$Player.Add_Closing({
    $MediaPlayer.Close()
    Clear-Playlist
})

$OpenFile = $Player.FindName("OpenFile")
$OpenFile.Add_Click{(Add-File)}

$OpenDir = $Player.FindName("OpenDir")
$OpenDir.Add_Click({Add-Folder})

$OpenPath = $Player.FindName("OpenPath")
$OpenPath.Add_Click({
    if ($script:Playlist.Count -cge 1)
    {
        Open-Path -File $script:Playlist[$script:Index]
    }
})

$QuitPlayer = $Player.FindName("QuitPlayer")
$QuitPlayer.Add_Click({$Player.Close()})

$StartPlayback = $Player.FindName("StartPlayback")
$StartPlayback.Add_Click({Get-StatusPlay})

$PreviousPlayback = $Player.FindName("PreviousPlayback")
$PreviousPlayback.Add_Click({Open-Previous})

$NextPlayback = $Player.FindName("NextPlayback")
$NextPlayback.Add_Click({Open-Next})

$Random = $Player.FindName("Random")
$Random.Add_Click({
    if ($script:Playlist.Count -ge 2)
    {
        $MediaPlayer.Close()
        $Transfer = $script:Playlist | Sort-Object {Get-Random}
        $script:Playlist.Clear()
        $Transfer | ForEach-Object {$script:Playlist.Add($_)}
        $Transfer.Clear()
        $script:Index = 0
        Read-Music
    }
})

$CleanPlaylist = $Player.FindName("CleanPlaylist")
$CleanPlaylist.Add_Click({
    Clear-Playlist -Message "File d'attente vidé"
})

$OutPlayer = $Player.FindName("OutPlayer")
$OutPlayer.Add_Click({
    if ($OutPlayer.IsChecked -eq 'False')
    {
        $script:EndedPlaylist = 1
        $OutPlayer.IsChecked = $true
    }
    else
    {
        $script:EndedPlaylist = 0
        $OutPlayer.IsChecked = $false
    }
    if ($script:Playlist.Count -gt 0)
    {
        Get-MetaData -Path $script:Playlist[$script:Index]
    }
})

$OpenAbout = $Player.FindName("OpenAbout")
$OpenAbout.Add_Click({Invoke-About})

$TrackTitle = $Player.FindName("TrackTitle")

$TrackDuration = $Player.FindName("TrackDuration")

$ButtonPrevious = $Player.FindName("ButtonPrevious")
$ButtonPrevious.Add_Click({Open-Previous})

$ButtonPlay = $Player.FindName("ButtonPlay")
$ButtonPlay.Add_Click({Get-StatusPlay})

$ButtonNext = $Player.FindName("ButtonNext")
$ButtonNext.Add_Click({Open-Next})

$StatusLabel = $Player.FindName("StatusLabel")

$PlayerGUI = New-Object System.Windows.Forms.Form #Main Window
$PlayerGUI.AllowDrop = $true
$PlayerGUI.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$PlayerGUI.Height = 190
$PlayerGUI.MaximizeBox = $false
$PlayerGUI.ShowIcon = $false
$PlayerGUI.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$PlayerGUI.Text = 'Lecteur audio v1.3 Beta'
$PlayerGUI.Width = 370

<#$TrackTitle = New-Object System.Windows.Forms.Label
$TrackTitle.AutoSize = $false
$TrackTitle.Dock = [System.Windows.Forms.DockStyle]::Top
$TrackTitle.Font = New-Object System.Drawing.Font('SegoeUI', 9)
$TrackTitle.Height = 40
$TrackTitle.Text = 'Aucune piste en cours'
$TrackTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$PlayerGUI.Controls.Add($TrackTitle)#>

<#$TrackDuration = New-Object System.Windows.Forms.Label
$TrackDuration.Anchor = [System.Windows.Forms.AnchorStyles]::None
$TrackDuration.AutoSize = $true
$TrackDuration.Font = New-Object System.Drawing.Font('SegoeUI', 9)
$TrackDuration.Location = New-Object System.Drawing.Size(158, 68)
$TrackDuration.Text = '00:00'
$PlayerGUI.Controls.Add($TrackDuration)#>

$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 500

$MediaPlayer = New-Object System.Windows.Media.MediaPlayer
$MediaPlayer.Volume = 1
$handler_MediaPlayer_MediaEnded=#Detect ended file and read next file
{
    $Timer.Stop()
    if ($script:EndedPlaylist -eq 1 -and $script:Index -eq ($script:Playlist.Count) - 1)
    {
        $Player.Close()
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

$Player.ShowActivated = $true
[void]$Player.ShowDialog()