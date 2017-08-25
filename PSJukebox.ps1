#https://foxdeploy.com/2013/10/23/creating-a-gui-natively-for-your-powershell-tools-using-net-methods/
#https://msdn.microsoft.com/en-us/library/system.windows.forms.aspx
#https://foxdeploy.com/2014/02/17/two-ways-to-provide-gui-interaction-to-users/
#http://eddiejackson.net/wp/?p=9268


Add-Type -AssemblyName presentationCore
$mediaPlayer = New-Object system.windows.media.mediaplayer
$mediaPlayer2 = New-Object system.windows.media.mediaplayer

#Reserved XX# csv file
$Songs = Import-Csv ".\PSJukebox.ini"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

Add-Type -AssemblyName System.Windows.Forms
 $res = [System.Windows.Forms.Screen]::AllScreens | Where-Object {$_.Primary -eq 'True'} | Select-Object WorkingArea
 if (($res -split ',')[3].Substring(10,1) -match '}') {$heightend = 3}
 else {$heightend = 4}
 $w = ($res -split ',')[2].Substring(6)
 $h = ($res -split ',')[3].Substring(7,$heightend)
 'Screen Resolution: ' + $w + 'x' + $h

#dell venue, 1280 x 800
$w = 1280
$h = 800

$split = 8;

#begin to draw forms
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerShell Jukebox v1.0"
$Form.Size = New-Object System.Drawing.Size($w,$h)
$Form.StartPosition = "CenterScreen"
#$Form.BackColor = "MidnightBlue"
$statusBar1 = New-Object System.Windows.Forms.StatusBar

$resizeHandler = { DrawButtons }

$Form.Add_Resize( $resizeHandler )

$volDown = {
    $obj = new-object -com wscript.shell
    $obj.SendKeys([char]174)
}
$volUp = {
    $obj = new-object -com wscript.shell
    $obj.SendKeys([char]175)
}
$ff = {
    $mediaPlayer.Position = New-Object System.TimeSpan(0, 0, 0, ($mediaPlayer.Position.Seconds + 30), 0)
}
$rw = {
    $mediaPlayer.Position = New-Object System.TimeSpan(0, 0, 0, ($mediaPlayer.Position.Seconds - 10), 0)
}

$play_click =
{
    if($Songs[$this.Name].Type -eq 0){
        $statusBar1.Text = "Playing $($Songs[$this.Name].File)"
        $mediaPlayer.open([uri]($PSScriptRoot + "\" + "$($Songs[$this.Name].File)"))
        #Start-Sleep 2 # This allows the $wmplayer time to load the audio file
        #$duration = $mediaPlayer.NaturalDuration.TimeSpan.TotalSeconds
        $mediaPlayer.Play()
        #Start-Sleep $duration

    } else {
        $mediaPlayer2.open([uri]($PSScriptRoot + "\" + "$($Songs[$this.Name].File)"))
        $mediaPlayer2.Play()
    }
}

$stop_click =
{
  $statusBar1.Text = "Stopped"
  $mediaPlayer.Stop()
  $mediaPlayer2.Stop()
}

function DrawButtons{
    $Form.Controls.Clear()
    $w = ($Form.Size.Width-17)
    $h = ($Form.Size.Height-62)

    #$listbox = New-Object System.Windows.Forms.ListBox
    #$listbox.Location = New-Object System.Drawing.Size((($w/$split)*6),(($h/$split)*6))
    #$listbox.Size = New-Object System.Drawing.Size(($w/$split),(($h/$split)-50))
    #ForEach($song in $Songs){
    #    $listbox.Items.Add($song.Title)
    #}
    #$Form.Controls.Add($listbox)

    $statusBar1.Name = "statusBar1"
    $statusBar1.Text = "Stopped"
    $form.Controls.Add($statusBar1)
    $Buttons = @()

    for( $i=0; $i -lt 59; $i++ ){
        $button = New-Object System.Windows.Forms.Button
        $Buttons = $Buttons + $button
        $Buttons[$i].Name = $i;
        if($i -lt $split){
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * $i),0)
        } elseif($i -lt ($split * 2)) {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - $split)),($h/$split))
        } elseif($i -lt ($split * 3)) {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - ($split * 2))),(($h/$split) * 2))
        } elseif($i -lt ($split * 4)) {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - ($split * 3))),(($h/$split) * 3))
        } elseif($i -lt ($split * 5)) {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - ($split * 4))),(($h/$split) * 4))
        } elseif($i -lt ($split * 6)) {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - ($split * 5))),(($h/$split) * 5))
        } elseif($i -lt ($split * 7)) {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - ($split * 6))),(($h/$split) * 6))
        } else {
            $Buttons[$i].Location = New-Object System.Drawing.Size((($w/$split) * ($i - ($split * 7))),(($h/$split) * 7))
        }
        $Buttons[$i].Size = New-Object System.Drawing.Size(($w/$split),($h/$split))
        $Buttons[$i].Text = "$($Songs[$i].Title)"
        $Buttons[$i].Add_Click($play_click)
        if($Songs[$i].Type -eq 1){$Buttons[$i].BackColor="beige"}
        $Form.Controls.Add($Buttons[$i])
    }
    
    #VOL-
    $VDownButton = New-Object System.Windows.Forms.Button
    $VDownButton.Location = New-Object System.Drawing.Size((($w/$split)*3),(($h/$split)*7))
    $VDownButton.Size = New-Object System.Drawing.Size(($w/$split),(($h/$split)-0))
    $VDownButton.Text = "VOL-"
    $VDownButton.Add_Click($volDown)
    $Form.Controls.Add($VDownButton)

    #VOL+
    $VUpButton = New-Object System.Windows.Forms.Button
    $VUpButton.Location = New-Object System.Drawing.Size((($w/$split)*4),(($h/$split)*7))
    $VUpButton.Size = New-Object System.Drawing.Size(($w/$split),(($h/$split)-0))
    $VUpButton.Text = "VOL+"
    $VUpButton.Add_Click($volUp)
    $Form.Controls.Add($VUpButton)

    #RW
    $RWButton = New-Object System.Windows.Forms.Button
    $RWButton.Location = New-Object System.Drawing.Size((($w/$split)*5),(($h/$split)*7))
    $RWButton.Size = New-Object System.Drawing.Size(($w/$split),(($h/$split)-0))
    $RWButton.Text = "RW - 10"
    $RWButton.Add_Click($rw)
    $Form.Controls.Add($RWButton)

    #FF
    $FFButton = New-Object System.Windows.Forms.Button
    $FFButton.Location = New-Object System.Drawing.Size((($w/$split)*6),(($h/$split)*7))
    $FFButton.Size = New-Object System.Drawing.Size(($w/$split),(($h/$split)-0))
    $FFButton.Text = "FF +30"
    $FFButton.Add_Click($ff)
    $Form.Controls.Add($FFButton)

    #STOP
    $StopButton = New-Object System.Windows.Forms.Button
    $StopButton.Location = New-Object System.Drawing.Size((($w/$split)*7),(($h/$split)*7))
    $StopButton.Size = New-Object System.Drawing.Size(($w/$split),(($h/$split)-0))
    $StopButton.Text = "STOP"
    $StopButton.Add_Click($stop_click)
    $Form.Controls.Add($StopButton)

}

DrawButtons

#$Form.Add_KeyDown({if ($_.KeyCode -eq "F1"){&amp; $add_reservation_click}})
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape")
{$Form.Close()}})



#Show form
$Form.Topmost = $False
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()


