# Single Threaded 
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    if ($PSCommandPath) {
        powershell -STA -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
        exit
    } else {
        Write-Warning "Relaunching a new STA PowerShell..."
        powershell -STA -NoProfile -ExecutionPolicy Bypass
        exit
    }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$logoUrl = "https://www.freenome.com/wp-content/uploads/2024/02/FreenomeLogo1200x630.jpg"
$tempLogoPath = Join-Path $env:TEMP "freenome_logo.jpg"
try {
    Invoke-WebRequest -Uri $logoUrl -OutFile $tempLogoPath -UseBasicParsing -ErrorAction Stop
} catch {
    $tempLogoPath = $null
}

$bitmap = $null
if ($tempLogoPath -and (Test-Path $tempLogoPath)) {
    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.UriSource  = [Uri]::new($tempLogoPath)
    $bitmap.EndInit()
    if ($bitmap.CanFreeze) { $bitmap.Freeze() }
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Freenome IOQ Export Assistant"
        Height="640" Width="720"
        WindowStartupLocation="CenterScreen"
        Background="#F6F8FA">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" CornerRadius="16" Padding="16" Background="#FFFFFF" BorderBrush="#E5E7EB" BorderThickness="1">
            <DockPanel LastChildFill="False">
                <StackPanel DockPanel.Dock="Left" Orientation="Vertical" Margin="0,0,12,0">
                    <!-- NOTE: & must be escaped in XML -->
                    <TextBlock Text="IOQ Export &amp; Upload" FontSize="22" FontWeight="Bold" Foreground="#111827" />
                    <TextBlock Text="Enter your name and NetApp password, then export or upload the CSV." 
                               Margin="0,6,0,0" Foreground="#4B5563" />
                </StackPanel>
                <Image Name="Logo" Height="60" Width="200" Stretch="Uniform" DockPanel.Dock="Right" />
            </DockPanel>
        </Border>

        <!-- Card -->
        <Border Grid.Row="1" Margin="0,12,0,12" CornerRadius="18" Padding="18" Background="#FFFFFF" BorderBrush="#E5E7EB" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Margin="0,0,0,12">
                    <TextBlock Text="Technician Information" FontWeight="SemiBold" Foreground="#111827" />
                    <TextBlock Text="Used to tag the CSV (Submitted By + date)." FontSize="12" Foreground="#6B7280"/>
                </StackPanel>

                <!-- Name -->
                <StackPanel Grid.Row="1" Margin="0,0,0,16">
                    <TextBlock Text="Your Name" FontSize="12" Foreground="#374151" Margin="0,0,0,6"/>
                    <TextBox Name="UserNameBox" Height="32" Padding="8" />
                </StackPanel>

                <!-- Password -->
                <StackPanel Grid.Row="2" Margin="0,0,0,16">
                    <TextBlock FontSize="12" Foreground="#374151" Margin="0,0,0,6">
                        <Run Text="NetApp Password for " />
                        <Run Text="freenomelab" FontWeight="Bold"/>
                    </TextBlock>

                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="110"/>
                        </Grid.ColumnDefinitions>

                        <Grid Grid.Column="0">
                            <PasswordBox Name="PasswordBox" Height="32" Padding="8" />
                            <TextBox Name="PasswordTextBox" Height="32" Padding="8" Visibility="Collapsed" IsReadOnly="True"/>
                        </Grid>

                        <Button Name="ShowPasswordBtn" Grid.Column="1" Height="32" Margin="8,0,0,0"
                                Content="Hold to Show" ToolTip="Hold mouse down to reveal"
                                Background="#111827" Foreground="White" FontWeight="SemiBold"
                                BorderBrush="#111827" />
                    </Grid>
                </StackPanel>

                <!-- Buttons + Status -->
                <StackPanel Grid.Row="3" Orientation="Vertical">
                    <UniformGrid Columns="2" Rows="1" Margin="0,0,0,12">
                        <Button Name="UploadButton" Height="44" Margin="0,0,8,0"
                                Content="1) Export + Upload to NetApp (Default)"
                                Background="#31B39D" Foreground="White" FontWeight="Bold" BorderBrush="#31B39D"/>
                        <Button Name="ExportButton" Height="44" Margin="8,0,0,0"
                                Content="2) Export Locally to Desktop (Backup)"
                                Background="#2563EB" Foreground="White" FontWeight="Bold" BorderBrush="#2563EB"/>
                    </UniformGrid>

                    <TextBlock Text="Status" FontWeight="SemiBold" Foreground="#111827" Margin="0,0,0,6"/>
                    <Border CornerRadius="10" BorderBrush="#E5E7EB" BorderThickness="1" Background="#F9FAFB">
                        <TextBox Name="StatusBox" Height="180" Padding="8"
                                 IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                                 TextWrapping="Wrap" Background="#F9FAFB" Foreground="#111827"/>
                    </Border>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Footer -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <TextBlock Text="Freenome IT • IOQ" Foreground="#6B7280" FontSize="12"/>
        </StackPanel>
    </Grid>
</Window>
"@

try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "XAML parse failed: $($_.Exception.Message)"
    throw
}

if (-not $window) {
    throw "XAML failed to load (window is null). Check the XAML string."
}

$Logo            = $window.FindName("Logo")
$ExportBtn       = $window.FindName("ExportButton")
$UploadBtn       = $window.FindName("UploadButton")
$PasswordBox     = $window.FindName("PasswordBox")
$PasswordTextBox = $window.FindName("PasswordTextBox")
$ShowPasswordBtn = $window.FindName("ShowPasswordBtn")
$StatusBox       = $window.FindName("StatusBox")
$UserNameBox     = $window.FindName("UserNameBox")

$all = @{
    Logo=$Logo; ExportButton=$ExportBtn; UploadButton=$UploadBtn; PasswordBox=$PasswordBox;
    PasswordTextBox=$PasswordTextBox; ShowPasswordBtn=$ShowPasswordBtn; StatusBox=$StatusBox; UserNameBox=$UserNameBox
}
$missing = $all.GetEnumerator() | Where-Object { -not $_.Value } | Select-Object -ExpandProperty Key
if ($missing) {
    throw "Missing named controls in XAML: $($missing -join ', ')"
}

if ($Logo -and $bitmap) { $Logo.Source = $bitmap }


$global:UploadAttempts = 0
$MaxUploadAttempts = 5

$handler = [System.Windows.RoutedEventHandler]{ $PasswordTextBox.Text = $PasswordBox.Password }
$PasswordBox.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, $handler)
$ShowPasswordBtn.Add_PreviewMouseDown({
    $PasswordTextBox.Text = $PasswordBox.Password
    $PasswordBox.Visibility = 'Collapsed'
    $PasswordTextBox.Visibility = 'Visible'
})
$ShowPasswordBtn.Add_PreviewMouseUp({
    $PasswordBox.Visibility = 'Visible'
    $PasswordTextBox.Visibility = 'Collapsed'
})

function Write-Log($msg) {
    $StatusBox.AppendText("$msg`r`n")
    $StatusBox.ScrollToEnd()
}

function pcData {
    $pc = $env:COMPUTERNAME
    $serial = (Get-CimInstance Win32_BIOS).SerialNumber
    $model = (Get-CimInstance Win32_ComputerSystem).Model
    $os = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $osv = "$($os.CurrentBuildNumber).$($os.UBR)"
    $domain = (Get-CimInstance Win32_ComputerSystem).Domain
    $userName = $UserNameBox.Text
    $csvDate = Get-Date -Format "yyyy-MM-dd"

    function Get-App($name) {
        $paths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        foreach ($p in $paths) {
            $apps = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue
            foreach ($app in $apps) {
                if ($app.DisplayName -like "*$name*") { return $app.DisplayVersion }
            }
        }
        return "Not Installed"
    }

    function Get-ElasticVersion {
        $yaml = Get-ChildItem "C:\Program Files\Elastic\Agent\data\elastic-agent-*\manifest.yaml" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($yaml) {
            foreach ($line in Get-Content $yaml.FullName) {
                if ($line -match "^\s*version:\s*([\d\.]+)") { return $Matches[1] }
            }
        }
        return "Not Installed"
    }

    function Get-SecureBoot {
        try { if (Confirm-SecureBootUEFI) { "Enabled" } else { "Disabled" } } catch { "Unsupported" }
    }

    function Get-BitLocker {
        try {
            $bit = Get-BitLockerVolume -MountPoint "C:"
            if ($bit.ProtectionStatus -eq 1) { "Enabled" } else { "Disabled" }
        } catch { "Unknown" }
    }

    function Get-PasswordRequired {
        try {
            $key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
            $auto = Get-ItemProperty -Path $key -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
            if ($auto.AutoAdminLogon -eq "1") { "No" } else { "Yes" }
        } catch { "Unknown" }
    }

    return @(
        [PSCustomObject]@{Field="PC"; Value=$pc},
        [PSCustomObject]@{Field="Model"; Value=$model},
        [PSCustomObject]@{Field="Serial"; Value=$serial},
        [PSCustomObject]@{Field="OS Version"; Value=$osv},
        [PSCustomObject]@{Field="Domain"; Value=$domain},
        [PSCustomObject]@{Field="Password Required"; Value=Get-PasswordRequired},
        [PSCustomObject]@{Field="Secure Boot"; Value=Get-SecureBoot},
        [PSCustomObject]@{Field="BitLocker (C:)"; Value=Get-BitLocker},
        [PSCustomObject]@{Field="Chrome"; Value=Get-App "Google Chrome"},
        [PSCustomObject]@{Field="Elastic Search Agent"; Value=Get-ElasticVersion},
        [PSCustomObject]@{Field="Excel"; Value=Get-App "Microsoft Excel"},
        [PSCustomObject]@{Field="Word"; Value=Get-App "Microsoft Word"},
        [PSCustomObject]@{Field="Freshservice"; Value=Get-App "Freshservice"},
        [PSCustomObject]@{Field="ManageEngine UEMS"; Value=Get-App "ManageEngine UEMS"},
        [PSCustomObject]@{Field="Idemeum"; Value=Get-App "Idemeum"},
        [PSCustomObject]@{Field="NiceLabel"; Value=Get-App "NiceLabel"},
        [PSCustomObject]@{Field="Submission Date"; Value=$csvDate},
        [PSCustomObject]@{Field="Submitted By"; Value=$userName}
    )
}

$UploadBtn.Add_Click({
    $global:UploadAttempts++
    $pc = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $networkUser = "freenomelab"
    $networkPath = "\\10.15.44.15\fn-IT\IOQ_Export"
    $netFileName = "${pc}_${date}.csv"
    $localTempPath = Join-Path $env:TEMP $netFileName
    $desktopPath   = Join-Path ([Environment]::GetFolderPath("Desktop")) $netFileName

    if (-not $PasswordBox.SecurePassword -or $PasswordBox.SecurePassword.Length -eq 0) {
        Write-Log "Please enter a password before uploading."
        return
    }
    if (-not $UserNameBox.Text) {
        Write-Log "Please enter your name before exporting."
        return
    }

    if (-not (Test-Path $localTempPath)) {
        try {
            pcData | Export-Csv -Path $localTempPath -NoTypeInformation -Force
            Write-Log "CSV generated: $localTempPath"
        } catch {
            Write-Log "Failed to generate CSV: $($_.Exception.Message)"
            return
        }
    }

    Write-Log "Upload attempt $UploadAttempts of $MaxUploadAttempts..."

    try {
        $plainPwd = $PasswordBox.Password
        $netUseCmd = "net use \\10.15.44.15\fn-IT /user:$networkUser $plainPwd"
        $netUseResult = cmd /c $netUseCmd 2>&1
        if ($netUseResult -match "error|failed|incorrect|denied") {
            throw "Invalid password or access denied."
        }

        $cred = New-Object System.Management.Automation.PSCredential ($networkUser, $PasswordBox.SecurePassword)
        New-PSDrive -Name "Z" -PSProvider FileSystem -Root $networkPath -Credential $cred -ErrorAction Stop | Out-Null
        Copy-Item -Path $localTempPath -Destination "Z:\$netFileName" -Force -ErrorAction Stop
        Remove-PSDrive -Name "Z" -Force

        Write-Log "✅ Upload successful! File copied to NetApp."
        $UploadBtn.IsEnabled = $false
    } catch {
        Write-Log "❌ Upload failed: $($_.Exception.Message)"
        if ($UploadAttempts -eq $MaxUploadAttempts) {
            Write-Log "Max attempts reached. Saving to Desktop..."
            try {
                Copy-Item -Path $localTempPath -Destination $desktopPath -Force
                Write-Log " File saved to Desktop: $desktopPath"
            } catch {
                Write-Log "Desktop fallback failed: $($_.Exception.Message)"
            }
            $UploadBtn.IsEnabled = $false
        } else {
            Write-Log "Try again. You have $($MaxUploadAttempts - $UploadAttempts) attempt(s) left."
        }
    }
})

$ExportBtn.Add_Click({
    if (-not $UserNameBox.Text) {
        Write-Log "Please enter your name before exporting."
        return
    }

    $pc = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $desktopFolder = Join-Path ([Environment]::GetFolderPath("Desktop")) "IOQ_Export"
    $finalPath = Join-Path $desktopFolder ("{0}_{1}.csv" -f $pc, $date)

    if (-not (Test-Path $desktopFolder)) {
        New-Item -ItemType Directory -Path $desktopFolder | Out-Null
    }

    try {
        pcData | Export-Csv -Path $finalPath -NoTypeInformation -Force
        Write-Log "✅ Exported to Desktop: $finalPath"
    } catch {
        Write-Log "Export failed: $($_.Exception.Message)"
    }
})

$null = $window.ShowDialog()

if ($tempLogoPath) {
    Remove-Item -Path $tempLogoPath -Force -ErrorAction SilentlyContinue
}
