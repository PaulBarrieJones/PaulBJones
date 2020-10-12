#######################################################################################################################


#### Paul Jones - Take-Two - Script Example 4 - LogonMessage
####
#### This solution allows for a branded message box to be presented to the end user when they login. This can have
#### many aditional features and is much richer than any built-in login script message. This demo shows the method,
#### I have created advanced toolsets using this approach. This system is easily transferable and it makes designing
#### any GUI very straightforward. The company logo was included in the script as BASE64.


#######################################################################################################################

# # Requires PowerShell Version 3 or above  >

<#

.SYNOPSIS
  Logon script to pop-up branded message box to user upon login.

.DESCRIPTION
Message box designed in Visual Studio, XML code is then pasted between the @" "@ at line 95. Any active objects from 
Visual Studio desing will be collected and processed by REGEX.

.INPUTS


.OUTPUTS


.NOTES
  Version:          1
  Author:           Paul Jones
  Creation Date:    15/03/2017
  Purpose/Change:   Project
  Change Ref:       %
  Change Type:      Normal
  
.EXAMPLE
  None

#>


#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Gather initial system varibles at PowerShell instance start
$DefaultVariables = Get-Variable | Select-Object -ExpandProperty Name

# Set Error Verbose and Warning
$ErrorActionPreference = "Continue"
$VerbosePreference = "continue"
$WarningPreference = "continue"
$ExitCode = 0

# Load Required Module/Function Libraries

# Load assembly for GUI
Add-Type -AssemblyName PresentationFramework

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script Version
$ScriptVersion = "1"
$ProjectName = "LogonMessage"
$ProjectPath = "C:\PowerShell\Projects\$ProjectName"
$LogFilePath = "$ProjectPath\Logs\$ProjectName.log"
$ScriptPath = "$ProjectPath\Scripts\"

# Log File Info
$LogFilePath = "$ProjectPath\Logs\$ProjectName.log"


#-----------------------------------------------------------[Functions]------------------------------------------------------------

### Logging function

function Write-Log {
    
    param (
        
        [Parameter(Mandatory=$False, Position=0)]
        [String]$Entry
    
    )

    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') $Entry" | Out-File -FilePath $LogFilePath -Append
}


# Main function to generate message box from XML pasted in from Visual Studio
Function Open-MainForm {

# This is the XML used to create the message box
[XML]$LoadMainForm = @"

<Window

xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="CompanyName" Height="529.306" Width="538" FontSize="14" ResizeMode="NoResize" HorizontalAlignment="Left" VerticalAlignment="Top" BorderBrush="#FF0074FF" WindowStyle="None">

    <Window.Effect>
        <DropShadowEffect/>
    </Window.Effect>

    <Grid x:Name="ButtonMove" Margin="0,0,-6,0">
        <Grid.RowDefinitions>
            <RowDefinition/>
        </Grid.RowDefinitions>

        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="700"/>
        </Grid.ColumnDefinitions>

        <Image Name="Logo" Margin="25,-1,486,406"/>
        <Label x:Name="TextBoxTitle" Content="Window Virtual Desktop&#xD;&#xA;" HorizontalAlignment="Right" Margin="0,40,197,0" VerticalAlignment="Top" Height="43" Width="256" FontSize="24" FontFamily="Calibri Light" FontWeight="Bold" RenderTransformOrigin="0.513,-0.946"/>
        <RichTextBox x:Name="TextBoxMain" HorizontalAlignment="Left" Margin="20,137,0,81" Width="483" BorderBrush="{x:Null}" SelectionBrush="{x:Null}" Background="{x:Null}" BorderThickness="0" SelectionOpacity="0">
            <FlowDocument>
                <Paragraph>
                    <Run FontWeight="Bold" Text=""/>
                </Paragraph>
                <Paragraph>
                    <Run Text=""/>
                </Paragraph>
                <Paragraph>
                    <Run Text=""/>
                </Paragraph>
                <Paragraph>
                    <Run Text=""/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button x:Name="ButtonClose" Content="Close" HorizontalAlignment="Left" Margin="404,463,0,0" VerticalAlignment="Top" Width="75" Height="27" RenderTransformOrigin="0.56,-2.741"/>
        <Label x:Name="TextBoxTitle_Copy" Content="Important Information" HorizontalAlignment="Right" Margin="0,65,221,0" VerticalAlignment="Top" Height="43" Width="231" FontSize="20" FontFamily="Calibri Light" FontWeight="Bold" RenderTransformOrigin="0.513,-0.946"/>

    </Grid>

</Window>
"@

# Load the aboce XML into an XML object
$reader=(New-Object System.Xml.XmlNodeReader $LoadMainForm)
$Window=[Windows.Markup.XamlReader]::Load($reader)

# Find objects such as textboxes and buttons and load them into variables
$LoadMainForm.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {

New-Variable -Scope global -Name $_.Name -Value $Window.FindName($_.Name) -Force

}


$base64Logo = # The company logo was stored in the script as Base64, I removed it for this demo.


# Create a streaming image by streaming the base64 string to a bitmap streamsource
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64Logo)
$bitmap.EndInit()

# Freeze() prevents memory leaks.
$bitmap.Freeze()

# Set source here. Take note in the XAML as to where the variable name was taken.
$btplogo.source = $bitmap


# Reset textboxes

# Make message box appear center of the screen when opened
$Window.WindowStartupLocation = "CenterScreen"

# Prevent typing in the text box
$TextBoxMain.IsReadOnly = $true

# Link close button with close function
$ButtonClose.add_click({Close-form})

# Prepare message box for opening
$Window.add_loaded({})

# Open GUI
$Null = $Window.ShowDialog()


}


# Fuction to close form with button
Function Close-form {

$Window.Close()

}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

### Log Start
Write-Log -Entry "Script DHCP started on $(Get-Date -Format 'dddd, MMMM dd, yyyy')."
Write-Log -Entry "Script Version: $ScriptVersion"
Write-Log -Entry "Script Path: $ScriptPath"

###################################################################################################################################

# Call function (no parameters required)

Open-MainForm




#-----------------------------------------------------------[Cleanup]------------------------------------------------------------

$UserVariables = Get-Variable | Select-Object -ExpandProperty Name | Where-Object {$DefaultVariables -notcontains $_ -and $_ -ne "ExistingVariables"}
Remove-Variable $UserVariables
 

### Log End
Write-Log -Entry "Script ended ($ExitCode)."
### 

### Script END