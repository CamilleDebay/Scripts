param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="PowerShell Script Path")][String] $FilePath,
    [Parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Architecture")][ValidateSet('32','64')][Int] $Arch
)

if (Test-Path -Path $FilePath )
{   
    $Base64Script = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes((Get-content -Path $FilePath -Raw)))
}
else {
    Write-Error -Message "File doesn't exist." -ErrorAction Stop
    Exit
}

if ( $Arch -eq 32)
{
    $CommandLine = "powershell -EncodedCommand " + $Base64Script
}
else {
    $CommandLine = "&`$env:SystemRoot\sysnative\WindowsPowerShell\v1.0\powershell.exe -EncodedCommand " + $Base64Script
}

[XML] $WapProfile = '<wap-provisioningdoc id="" name="customprofile"><characteristic type="com.airwatch.winrt.powershellcommand" uuid=""><parm name="PowershellCommand" value=""/></characteristic></wap-provisioningdoc>'

#Generating a new GUID for the profile
Write-Debug -Message "Generating New GUIDs"
$WapProfile.'wap-provisioningdoc'.id = [GUID]::NewGuid().ToString()
$WapProfile.'wap-provisioningdoc'.characteristic.uuid = [GUID]::NewGuid().ToString()

#Adding Command Line
Write-Debug -Message "Addnig Command Line"
$WapProfile.'wap-provisioningdoc'.characteristic.parm.value = $CommandLine

#Exporting XML to File
Write-Debug -Message "Writing file"
try {
    $FileDirectory = (Get-Item -Path $FilePath).Directory.ToString()
    $XmlFileName = "PSProfile-{0:yyyyMMdd-HHmmss}.xml" -f (Get-Date)
    $XmlFilePath = $FileDirectory + "\" + $XmlFileName
    [System.Xml.XmlWriterSettings] $XmlSettings = New-Object System.Xml.XmlWriterSettings
    #Preserve Windows formating
    $XmlSettings.Indent = $true
    $XmlSettings.OmitXmlDeclaration = $true
    #Keeping UTF-8 without BOM
    $XmlSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
    [System.Xml.XmlWriter] $XmlWriter = [System.Xml.XmlWriter]::Create($XmlFilePath, $XmlSettings)
    $WapProfile.Save($XmlWriter)
}
catch
{
    Write-Debug -Message "Error Catched"
    Write-Error "XML Writing Error"
    Write-Error -Message $Error[0]
}
finally
{
    if ($XmlWriter -ne $null)
    {
        Write-Debug -Message "Clearing"
        $XmlWriter.Dispose()
    }
}