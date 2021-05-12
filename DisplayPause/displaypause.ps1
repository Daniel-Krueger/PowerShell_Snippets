#region setup
# labels
$formTitelText = "Select time to pause"
$playReminderText = "Play reminder"
$okButtonText = "OK"
$cancelButtonText = "Cancel"
# predefined settings
$pauseTitleText = "Pause ends"
$numberOf5MinuteOptions = 16
$reminderChecked = $false
$alarmSoundPath = "$($env:windir)\Media\alarm01.wav"
$playReminderMinutesBeforePauseEnd = 2
#endregion

#region Form setup
# Drop down sample taken from 
# https://docs.microsoft.com/de-de/powershell/scripting/samples/selecting-items-from-a-list-box?view=powershell-7.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = $formTitelText
$form.Size = New-Object System.Drawing.Size(200,365)
$form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(70,20)
$label.Text = 'Pause name'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(80,15)
$textBox.Size = New-Object System.Drawing.Size(89,20)
$textBox.Text = $pauseTitleText
$form.Controls.Add($textBox)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(11,40)
$listBox.Size = New-Object System.Drawing.Size(158,280)
$listBox.Height = 230


$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,270)
$label.Size = New-Object System.Drawing.Size(90,20)
$label.Text = $playReminderText
$form.Controls.Add($label)

$reminderBox = New-Object System.Windows.Forms.CheckBox
$reminderBox.Location = New-Object System.Drawing.Point(105,267)
$reminderBox.Size = New-Object System.Drawing.Size(20,20)
$reminderBox.Checked = $reminderChecked;
$form.Controls.Add($reminderBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10,292)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = $okButtonText
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(95,292)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = $cancelButtonText
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)
#endregion 

#region populate drop down
$currentTime = (Get-Date)
for ($i = 0; $i -lt $numberOf5MinuteOptions;$i++){
    $currentTime= $currentTime.AddMinutes(5- $currentTime.Minute %5)
    [void] $listBox.Items.Add($currentTime.ToShortTimeString())
}
$form.Controls.Add($listBox)
#endregion

$form.Topmost = $true

$result = $form.ShowDialog()
if ($result -ne [System.Windows.Forms.DialogResult]::OK)
{
    exit
}

#region calculate url
$selectedTime = $listBox.SelectedItem
$pauseUntil = [datetime]::Parse((Get-Date).ToShortDateString()+' '+$selectedTime)
$url = "https://webuhr.de/embed/timer/"
$url += "#date="+$pauseUntil.ToString("s")
$url += "&theme=0"
$url += "&title="+[uri]::EscapeDataString($textBox.Text)
$url += "&ampm="
if ((Get-Culture).DateTimeFormat.AMDesignator -eq ''){
    $url += 0
}
else{
    $url += 1
}
#endregion 
    
#region start browser
# Using a complex alternative to get a predefined sequence which browsers should be tested.
$browsers = new-object  "System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[[string],[string]]]"
$browsers.Add((New-Object "System.Collections.Generic.KeyValuePair[[string],[string]]" -ArgumentList @("msedge.exe","-InPrivate $url")))
$browsers.Add((New-Object "System.Collections.Generic.KeyValuePair[[string],[string]]" -ArgumentList @("chrome.exe","--incognito $url")))
$browsers.Add((New-Object "System.Collections.Generic.KeyValuePair[[string],[string]]" -ArgumentList @("firefox.exe","-private-window $url")))
$browsers.Add((New-Object "System.Collections.Generic.KeyValuePair[[string],[string]]" -ArgumentList @("iexplore.exe","-private $url")))

$runningProcess = $null
foreach ($item in $browsers){
    #$item = $browsers[0]
    try{
        $runningProcess = [System.Diagnostics.Process]::Start($item.Key, $item.Value)
    }
    catch {
    }   
    if ($runningProcess -ne $null){
        break;
    }
}
      
if ($runningProcess -eq $null){
    Write-Host  "Browser could not be started to load url"
    Write-Host  "$url" -ForegroundColor Green
    Read-Host "Press key to exit"
}

#endregion

#region Play pause, if requested
if ($reminderBox.Checked){   
    $totalPause = $pauseUntil.AddMinutes($playReminderMinutesBeforePauseEnd*-1).Subtract((Get-Date))   
    Write-Host  "A reminder will be sounded at $($pauseUntil.AddMinutes($playReminderMinutesBeforePauseEnd*-1).ToShortTimeString()), if the window remains open."
    if ($totalPause.TotalSeconds -gt 0){
        sleep -Seconds $totalPause.TotalSeconds
    }
    $PlayWav=New-Object System.Media.SoundPlayer
    $PlayWav.SoundLocation=$alarmSoundPath
    $PlayWav.playsync()
}

#endregion