Add-Type -AssemblyName System.Windows.Forms; 
    Add-Type -AssemblyName System.Drawing; 
    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon; 
    $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information; 
    $notifyIcon.Text = 'Time counter'; 
    $notifyIcon.Visible = $true; 
    $notifyIcon.ShowBalloonTip(5000, 'your time is up!', 'fucking stop playing now!!!', [System.Windows.Forms.ToolTipIcon]::Info); 
    Start-Sleep -Seconds 5; 
    $notifyIcon.Dispose();