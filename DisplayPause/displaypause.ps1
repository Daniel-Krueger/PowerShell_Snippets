
#region Form setup
# Drop down sample taken from 
# https://docs.microsoft.com/de-de/powershell/scripting/samples/selecting-items-from-a-list-box?view=powershell-7.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select time to pause'
$form.Size = New-Object System.Drawing.Size(200,350)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10,270)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(95,270)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(30,20)
$label.Text = 'Title'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(40,15)
$textBox.Size = New-Object System.Drawing.Size(129,20)
$textBox.Text = 'Pause ends'
$form.Controls.Add($textBox)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(11,40)
$listBox.Size = New-Object System.Drawing.Size(158,280)
$listBox.Height = 230
#endregion 

#region populate drop down
$currentTime = (Get-Date)
for ($i = 0; $i -lt 16;$i++){
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
$pauseUntil = [datetime]::Parse($currentTime.ToShortDateString()+' '+$selectedTime)
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