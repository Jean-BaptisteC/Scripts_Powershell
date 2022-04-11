Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = $Formtitle
$form.Size = '400,320'
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = $form.Size
$form.MaximizeBox = $False
$form.Topmost = $True
$form.AutoSize = $True
$form.AllowDrop = $True

$Form_DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
	{
	    $_.Effect = 'Copy'
	}
    Else
	{
	    $_.Effect = 'None'
	}
    }

    $Form_DragDrop = [System.Windows.Forms.DragEventHandler]{
	    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
        {
		    Start-Process $filename
	    }
    }

$form.Add_DragOver($Form_DragOver)
$form.Add_DragDrop($Form_DragDrop)
$form.ShowDialog()