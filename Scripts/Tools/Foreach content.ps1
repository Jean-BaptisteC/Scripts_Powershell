$File = "C:\example.txt"
$Data = @(
    $script:OSName.Text
    $script:OSManufacturer.Text
    $script:OSModel.Text
    $script:SystemBox.Text
    $script:OSArchitecture.Text
    $script:OSVersion.Text
    $script:OSSerialNumber.Text
    $script:OSInstallDate.Text
    $script:OSLanguages.Text
    $script:OSBootDevice.Text
    $script:OSWindowsDrive.Text
    $script:OSWindowsDirectory.Text
    $script:OSLastBoot.Text
    $script:Domain.Text
    $script:OSNBRAM.Text
    $script:CPUBox.Text
    $script:CPUManufacturer.Text
    $script:CPUDataWidth.Text
    $script:CPUCCS.Text
    $script:CPUSD.Text
    $script:CPUArchitecture.Text
    $script:CPUNbCores.Text
    $script:CPUNbLP.Text
    $script:CPUVirtualisation.Text
    )

function Write-Content{
    param(
        $Path,
        $Content
    )
    $Content | ForEach-Object{
        Add-Content -Path $Path -Value $_ -Encoding UTF8
    }
}

Write-Content -Path $File -Content $Data