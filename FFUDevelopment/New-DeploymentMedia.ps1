[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateScript({ Test-Path $_ })]
    [String]$FFUDevelopmentPath = "C:\FFUDevelopment",
    [String]$DeployISO,
    [String]$DeployISOPath
    )


#Set default values for variables that depend on other parameters

if (-not $LogFile) { $LogFile = "$FFUDevelopmentPath\FFUDevelopment.log" }
if (-not $adkPath) { $adkPath = "C:\Program Files (x86)\Windows Kits\10\" }
if (-not $Date) {$Date = get-date -UFormat %m%d%Y }
if (-not $DeployISO) { $DeployISO = "WinPE_FFU_Deploy-$Date.iso"}
if (-not $WindowsLang) { $WindowsLang = "En-US"}
if (-not $DeployISOPath) { $DeployISOPath = "$FFUDevelopmentPath\$DeployISO" }

function WriteLog($LogText) { 
    Add-Content -path $LogFile -value "$((Get-Date).ToString()) $LogText" -Force -ErrorAction SilentlyContinue
    Write-Verbose $LogText
}

function Invoke-Process {
    [CmdletBinding(SupportsShouldProcess)]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ArgumentList
    )

    $ErrorActionPreference = 'Stop'

    try {
        $stdOutTempFile = "$env:TEMP\$((New-Guid).Guid)"
        $stdErrTempFile = "$env:TEMP\$((New-Guid).Guid)"

        $startProcessParams = @{
            FilePath               = $FilePath
            ArgumentList           = $ArgumentList
            RedirectStandardError  = $stdErrTempFile
            RedirectStandardOutput = $stdOutTempFile
            Wait                   = $true;
            PassThru               = $true;
            NoNewWindow            = $true;
        }
        if ($PSCmdlet.ShouldProcess("Process [$($FilePath)]", "Run with args: [$($ArgumentList)]")) {
            $cmd = Start-Process @startProcessParams
            $cmdOutput = Get-Content -Path $stdOutTempFile -Raw
            $cmdError = Get-Content -Path $stdErrTempFile -Raw
            if ($cmd.ExitCode -ne 0) {
                if ($cmdError) {
                    throw $cmdError.Trim()
                }
                if ($cmdOutput) {
                    throw $cmdOutput.Trim()
                }
            }
            else {
                if ([string]::IsNullOrEmpty($cmdOutput) -eq $false) {
                    WriteLog $cmdOutput
                }
            }
        }
    }
    catch {
        #$PSCmdlet.ThrowTerminatingError($_)
        WriteLog $_
        Write-Host "Script failed - $Logfile for more info"
        throw $_

    }
    finally {
        Remove-Item -Path $stdOutTempFile, $stdErrTempFile -Force -ErrorAction Ignore
		
    }
	
}

function New-DeploymentMedia {
    param (
        [Parameter()]
        [String]$FFUDevelopmentPath = "C:\FFUDevelopment",
        [Parameter()]
        [Switch]$Deploy,
        [String]$WindowsLang = 'en-us'
    )
    #Need to use the Demployment and Imaging tools environment to create winPE media
    $DandIEnv = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\DandISetEnv.bat"
    $WinPEFFUPath = "$FFUDevelopmentPath\WinPE"
     $WinPEFFUMedia = "$WinPEFFUPath\Media"
    If (Test-path -Path "$WinPEFFUPath") {
        WriteLog "Removing old WinPE path at $WinPEFFUPath"
        Remove-Item -Path "$WinPEFFUPath" -Recurse -Force | out-null
    }

    WriteLog "Copying WinPE files to $WinPEFFUPath" 
    & cmd /c """$DandIEnv"" && copype amd64 $WinPEFFUPath" | Out-Null
    WriteLog "Copying tools"
    Copy-Item -Path "$FFUDevelopmentPath\DriverExporter.exe" -Destination "$WinPEFFUMedia\DriverExporter.exe"
    Copy-Item -Path "$FFUDevelopmentPath\DriverInstaller.exe" -Destination "$WinPEFFUMedia\DriverInstaller.exe"
    WriteLog 'Files copied successfully'
    if($WindowsLang -eq 'en-us'){
        WriteLog "if default (En-US) Clean up unused language folders from WinPE media folder"
        $UnusedFolder = Get-ChildItem -Path "$WinPEFFUMedia" -Directory | Where-Object {$_.Name -like "*-*"}
        foreach($folder in $UnusedFolder){
        remove-item -Path $folder.fullname -Force -Confirm: $false -Recurse
        }
    }else{
        WriteLog "Cleanup all language folders except $WindowsLang"
        $UnusedFolder = Get-ChildItem -Path "$WinPEFFUMedia" -Directory | Where-Object {$_.Name -like "*-*" -and $_.Name -notmatch $WindowsLang}
        foreach($folder in $UnusedFolder){
        remove-item -Path $folder.fullname -Force -Confirm: $false -Recurse
        }
    }
    Mount-WindowsImage -ImagePath "$WinPEFFUPath\media\sources\boot.wim" -Index 1 -Path "$WinPEFFUPath\mount" | Out-Null
    WriteLog 'Mounting complete'
    
    $Packages = @(
        "WinPE-WMI.cab",
        "en-us\WinPE-WMI_en-us.cab",
        "WinPE-NetFX.cab",
        "en-us\WinPE-NetFX_en-us.cab",
        "WinPE-Scripting.cab",
        "en-us\WinPE-Scripting_en-us.cab",
        "WinPE-PowerShell.cab",
        "en-us\WinPE-PowerShell_en-us.cab",
        "WinPE-StorageWMI.cab",
        "en-us\WinPE-StorageWMI_en-us.cab",
        "WinPE-DismCmdlets.cab",
        "en-us\WinPE-DismCmdlets_en-us.cab"
    )

    $PackagePathBase = "$adkPath`Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\"

    foreach ($Package in $Packages) {
        $PackagePath = Join-Path $PackagePathBase $Package
        WriteLog "Adding Package $Package"
        Add-WindowsPackage -Path "$WinPEFFUPath\mount" -PackagePath $PackagePath | Out-Null
        WriteLog "Adding package complete"
 
  
        WriteLog "Copying $FFUDevelopmentPath\WinPEDeployFFUFiles\* to WinPE deploy media"
        Copy-Item -Path "$FFUDevelopmentPath\WinPEDeployFFUFiles\*" -Destination "$WinPEFFUPath\mount" -Recurse -Force | Out-Null
        WriteLog 'Copy complete'
        if ($CopyPEDrivers) {
            WriteLog "Adding drivers to WinPE media"
            try {
                Add-WindowsDriver -Path "$WinPEFFUPath\Mount" -Driver "$FFUDevelopmentPath\PEDrivers" -Recurse -ErrorAction SilentlyContinue | Out-null
            }
            catch {
                WriteLog 'Some drivers failed to be added to the FFU. This can be expected. Continuing.'
            }
            WriteLog "Adding drivers complete"
                
    
    WriteLog 'Dismounting WinPE media' 
    Dismount-WindowsImage -Path "$WinPEFFUPath\mount" -Save | Out-Null
    WriteLog 'Dismount complete'
    #Make ISO
    $OSCDIMGPath = "$adkPath`Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"
    $OSCDIMG = "$OSCDIMGPath\oscdimg.exe"
    WriteLog "Creating WinPE ISO at $FFUDevelopmentPath\$ISOFileName"
    Invoke-Process $OSCDIMG "-m -o -u2 -udfver102 -bootdata:2`#p0,e,b`"$OSCDIMGPath\etfsboot.com`"`#pEF,e,b`"$OSCDIMGPath\Efisys_noprompt.bin`" `"$WinPEFFUPath\media`" `"$DeployISOPath`""
    WriteLog "ISO created successfully"
    WriteLog "Cleaning up $WinPEFFUPath"
    Remove-Item -Path "$WinPEFFUPath" -Recurse -Force
    WriteLog 'Cleanup complete'
    
}
}
}
if (Test-Path -Path $Logfile) {
    Remove-item -Path $LogFile -Force
}

WriteLog 'Begin Logging'
New-DeploymentMedia -Deploy