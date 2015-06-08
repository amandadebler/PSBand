<#
 # BasicBand.ps1 - How to connect to your Band with PowerShell and make it do things
 # Amanda Debler (mandie@mandie.net)
 # Inspired by Stefan Stranger's work on an older release of Microsoft Band (pre-async):
 # https://github.com/stefanstranger/PowerShell/blob/master/Examples/MSBand_v0.1.ps1
 # 
#>


add-type -path 'C:\Program Files (x86)\Microsoft Band Sync\Microsoft.Band.dll'
add-type -Path 'C:\Program Files (x86)\Microsoft Band Sync\Microsoft.Band.Desktop.dll'

# Gets the first band it sees, which unless you have one connected via Bluetooth 
# and another via USB, should be fine

function Get-MSBandClient {

$bands = [Microsoft.Band.BandClientManager]::Instance.GetBandsAsync()
$bandClient = [microsoft.band.BandClientManager]::Instance.ConnectAsync($bands.result[0]).Result

$bandClient

}

#Learn about your band
function Get-MSBandVersionTable {
Param ($bandClient = $MSBandClient)

$hardwareVersion = $bandClient.GetHardwareVersionAsync().Result
$firmwareVersion = $bandClient.GetFirmwareVersionAsync().Result
$bandVersionTable = New-Object -TypeName PSCustomObject
$bandVersionTable | Add-Member -Name "HardwareVersion" -MemberType NoteProperty -Value $hardwareVersion
$bandVersionTable | Add-Member -Name "FirmwareVersion" -MemberType NoteProperty -Value $firmwareVersion
$bandVersionTable
}

# What kinds of vibrations can you send?

function Get-MSBandVibrationType {
[Microsoft.Band.Notifications.VibrationType].GetEnumNames()
}


function Send-MSBandVibration {
Param(
    $bandClient = $MSBandClient
   ,[ValidateSet({[Microsoft.Band.Notifications.VibrationType].GetEnumNames()})][string]$vibrationType='NotificationOneTone'
   )

    $bandClient.VibrateAsync([Microsoft.Band.Notifications.VibrationType]::$vibrationType)

}

# Colors for your themes - BandColor objects have R(ed), G(reen) and B(lue) parameters.

$red = New-Object Microsoft.Band.BandColor -ArgumentList 150,0,0
$limeGreen = New-Object Microsoft.Band.BandColor -ArgumentList 0,255,0
$yuckyGreen = New-Object Microsoft.Band.BandColor -ArgumentList 130,178,63

function New-MSBandTile {
Param(
$tileName = 'MyAwesomeTile'
, $largeIconImage
, $smallIconImage
)
# Gotta make a tile to send messages
$tileCapacity = $MSBandClient.GetRemainingTileCapacityAsync().Result
$tileGuid = [guid]::NewGuid()


$smallIconBitmap = New-Object -TypeName System.Windows.Media.Imaging.WriteableBitmap -ArgumentList 24,24,96,96,([System.Windows.Media.PixelFormats]::Bgr32),$null
$largeIconBitmap = New-Object -TypeName System.Windows.Media.Imaging.WriteableBitmap -ArgumentList 46,46,96,96,([System.Windows.Media.PixelFormats]::Bgr32),$null

$smallBandIcon = [Microsoft.Band.WriteableBitmapExtensions]::ToBandIcon($smallIconBitmap)

$myBandTile = new-object -TypeName Microsoft.Band.Tiles.BandTile -ArgumentList $tileGuid
$MSBandClient.SendMessageAsync($tileGuid, "Test", "This is just a test.", [System.DateTimeOffset]::Now, [Microsoft.Band.Notifications.MessageFlags]::ShowDialog)
$myBandTile
}

function ConvertTo-BitMapSource {
# from C# at http://stackoverflow.com/questions/94456/load-a-wpf-bitmapimage-from-a-system-drawing-bitmap
    param($bitmapFile)
    $bitmap = [System.Drawing.Bitmap]::FromFile($bitmapFile)
    $hBitmap = $bitmap.GetHbitmap()
    $bitmapSource = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap($hBitmap, [System.IntPtr]::Zero, [System.Windows.Int32Rect]::Empty, [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
    
}