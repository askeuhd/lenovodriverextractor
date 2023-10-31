# user input parameters
param (
    [string]$ExtractDir,
    [string]$DriverDir
)

# function sanatize quotes from copy pasted paths
function Remove-InputPathQuotes {
    param (
        [string]$InputPath
    )
    # Trim leading and trailing quotation marks if they are present
    return ($InputPath -replace '^"(.*)"$', '$1')
}

# function to check if the script is running as an administrator
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

### UAC prompt to run as administrator ###

# Load the necessary .NET assemblies for UAC prompt
Add-Type -AssemblyName PresentationFramework

# Define the XAML for the UA prompt window
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Run as Administrator" SizeToContent="WidthAndHeight">
    <StackPanel>
        <TextBlock FontSize="14" Padding="10">
            <Run FontWeight="Bold" Text="Warning: " /><Run Text="This script is not running as an administrator." /><LineBreak />
        </TextBlock>
        <TextBlock FontSize="12" Padding="10">
            <Run Text="You may have to accept a UAC prompt for each file." /><LineBreak />
            <Run Text="If you want to avoid this you can press Yes to open a UAC prompt to run this script as an administrator," /><LineBreak />
            <Run Text="press No to continue running this script as a non-privileged user," /><LineBreak />
            <Run Text="or press Cancel to exit." />
        </TextBlock>
        <Border Padding="10">
            <WrapPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Content="Yes" x:Name="yesButton" Width="75" Margin="5" />
            <Button Content="No" x:Name="noButton" Width="75" Margin="5" />
            <Button Content="Cancel" x:Name="cancelButton" Width="75" Margin="5" />
            </WrapPanel>
        </Border>
    </StackPanel>
</Window>
"@

# Check if the script is running as an administrator, and if not, show the UAC prompt
if (-not (Test-Admin)) {
    # Load the XAML for the window
    $xamlDocument = New-Object System.Xml.XmlDocument
    $xamlDocument.LoadXml($xaml)
    $reader = New-Object System.Xml.XmlNodeReader $xamlDocument
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Get references to the buttons
    $yesButton = $window.FindName("yesButton")
    $noButton = $window.FindName("noButton")
    $cancelButton = $window.FindName("cancelButton")

    # Add event handlers for the buttons
    $yesButton.Add_Click({ $window.DialogResult = $true; $window.Close() })
    $noButton.Add_Click({ $window.DialogResult = $false; $window.Close() })

    # Add an event handler for the Cancel button, to exit the script
    $cancelButton.Add_Click({
            $window.DialogResult = $null
            $window.Close()
            # Exit the script
            exit
        })

    # Show the dialog and get the result
    $result = $window.ShowDialog()

    # If 'yes' was pressed, relaunch the script with administrative privileges
    if ($result -eq $true) {
        # Relaunch the script with administrative privileges and wait for it to finish
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`" -ExtractDir `"$ExtractDir`" -DriverDir `"$DriverDir`"" -Verb RunAs -Wait
        # Exit the original script
        exit
    }

    # If 'no' was pressed, continue running the script as a non-privileged user
    elseif ($result -eq $false) {
        # Continue running the script as a non-privileged user
    }
}

# Get the path to the script
$scriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Load the necessary assemblies for the driverdir prompt
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Define the XAML for the driver directory prompt
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Directory" Height="300" Width="500">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBlock TextWrapping="Wrap">
            <Run>The DriverDir parameter is not set.</Run>
            <Run>Therefore the script assumes that the drivers are located in the same directory as the script.</Run>
            <LineBreak />
            <Run>You may choose a different path below:</Run>
        </TextBlock>
        <TextBox x:Name="pathTextBox" Grid.Row="1" TextWrapping="Wrap" />
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="browseButton" Content="Browse..." Margin="5" />
            <Button x:Name="okButton" Content="OK" Margin="5" />
        </StackPanel>
    </Grid>
</Window>
"@

# Load the XAML and create the window for the driver directory prompt
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get references to the controls for the driver directory prompt
$pathTextBox = $window.FindName("pathTextBox")
$browseButton = $window.FindName("browseButton")
$okButton = $window.FindName("okButton")

# Set the initial path in the text box for the driver directory prompt
$pathTextBox.Text = $scriptLocation

# If browse is clicked in the driver directory prompt, open a folder browser dialog
$browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.SelectedPath = $pathTextBox.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $pathTextBox.Text = $folderBrowser.SelectedPath
        }
    })

# if OK is clicked, set the driver directory to the path in the text box, and proceed
$okButton.Add_Click({
        $window.DialogResult = $true
        $window.Close()
    })

# Show the dialog and get the result
if ($window.ShowDialog() -eq $true) {
    # sanatize quotes from copy pasted paths
    $folderPath = Remove-InputPathQuotes -InputPath $pathTextBox.Text
}
else {
    $folderPath = $scriptLocation
}

# Determine the default extraction directory
$defaultExtractionDirectory = Join-Path $folderPath "extracted"

# Load the necessary assemblies for the extraction directory prompt
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Define the XAML for the window for the extraction directory prompt
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Extraction Directory" Height="300" Width="500">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBlock TextWrapping="Wrap">
            <Run>The ExtractDir parameter is not set.</Run>
            <Run>Therefore the script assumes that the extraction directory should be placed in the same directory as the driver files.</Run>
            <LineBreak />
            <Run>You may choose a different path below:</Run>
        </TextBlock>
        <TextBox x:Name="pathTextBox" Grid.Row="1" TextWrapping="Wrap" />
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="browseButton" Content="Browse..." Margin="5" />
            <Button x:Name="okButton" Content="OK" Margin="5" />
        </StackPanel>
    </Grid>
</Window>
"@

# Load the XAML and create the window for the extraction directory prompt
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get references to the controls for the extraction directory prompt
$pathTextBox = $window.FindName("pathTextBox")
$browseButton = $window.FindName("browseButton")
$okButton = $window.FindName("okButton")

# Set the initial path in the text box for the extraction directory prompt
$pathTextBox.Text = $defaultExtractionDirectory

# If browse is clicked in the extraction directory prompt, open a folder browser dialog
$browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.SelectedPath = $pathTextBox.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $pathTextBox.Text = $folderBrowser.SelectedPath
        }
    })

# If OK is clicked, set the extraction directory to the path in the text box, and proceed
$okButton.Add_Click({
        $window.DialogResult = $true
        $window.Close()
    })

# Show the dialog and get the result
if ($window.ShowDialog() -eq $true) {
    # sanatize quotes from copy pasted paths
    $extractionDirectory = Remove-InputPathQuotes -InputPath $pathTextBox.Text
}
else {
    $extractionDirectory = $defaultExtractionDirectory
}

# Load the necessary assemblies for the script output window
Add-Type -AssemblyName PresentationFramework

# Define the XAML for the window
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Extraction output" Height="800" Width="800">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBox x:Name="outputTextBox" Grid.Row="0" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" />
        <Grid Grid.Row="1">
            <ProgressBar x:Name="progressBar" Minimum="0" Maximum="100" />
            <TextBlock x:Name="progressText" HorizontalAlignment="Center" VerticalAlignment="Center" />
        </Grid>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="stopButton" Content="Stop" Margin="5" />
            <Button x:Name="exitButton" Content="Exit" Margin="5" IsEnabled="False" />
        </StackPanel>
    </Grid>
</Window>
'@

# Load the XAML and create the window for the script output
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get references to the controls for the script output
$outputTextBox = $window.FindName("outputTextBox")
$stopButton = $window.FindName("stopButton")
$exitButton = $window.FindName("exitButton")
$progressBar = $window.FindName("progressBar")
$progressText = $window.FindName("progressText")

# Create a PowerShell object for the extraction script
$ps = [powershell]::Create()

# Open a new runspace for the PowerShell object
$ps.Runspace = [runspacefactory]::CreateRunspace()

# Open the runspace and add the script to the PowerShell object
$ps.Runspace.Open()

# Create a script scope variable to hold the process
$script:process = $null

# Create a synchronized hashtable to store shared variables between the script and the GUI
$sharedVariables = [hashtable]::Synchronized(@{})

# Add the variables to the hashtable
$sharedVariables["outputTextBox"] = $outputTextBox
$sharedVariables["exitButton"] = $exitButton
$sharedVariables["folderPath"] = $folderPath
$sharedVariables["extractionDirectory"] = $extractionDirectory
# Add a flag to indicate whether the script should continue running
$sharedVariables["continueRunning"] = $true
$sharedVariables["progressBar"] = $progressBar
$sharedVariables["progressText"] = $progressText
$sharedVariables["totalFiles"] = (Get-ChildItem -Path $folderPath -Filter *.exe).Count

# Set the initial text for the progress bar
$progressText.Text = "0 / " + $sharedVariables["totalFiles"].ToString()

# Add the hashtable to the PowerShell object
$ps.Runspace.SessionStateProxy.SetVariable("sharedVariables", $sharedVariables)

# Add the script to the PowerShell object
$ps.AddScript({
        # Get the shared variables from the hashtable and set them to script scope variables
        $outputTextBox = $sharedVariables["outputTextBox"]
        $exitButton = $sharedVariables["exitButton"]
        $folderPath = $sharedVariables["folderPath"]
        $extractionDirectory = $sharedVariables["extractionDirectory"]
        $progressBar = $sharedVariables["progressBar"]
        $progressText = $sharedVariables["progressText"]
        $totalFiles = $sharedVariables["totalFiles"]

        # Initialize the file counter
        $fileCount = 0

        # Iterate over each .exe file in the folder
        Get-ChildItem -Path $folderPath -Filter *.exe | ForEach-Object {
            # Check the flag before processing each file
            if ($sharedVariables["continueRunning"]) {
                # Increment the file counter
                $fileCount++

                # Get the full path of the file
                $filePath = $_.FullName

                # Construct the directory for the /DIR= parameter
                $dirPath = Join-Path $extractionDirectory $_.BaseName

                # Construct the command arguments
                $arguments = "/VERYSILENT /DIR=`"$dirPath`" /EXTRACT=YES"

                # Notify the user that the extraction is starting
                $outputTextBox.Dispatcher.Invoke({
                        $outputTextBox.AppendText("Extracting $filePath to $dirPath...")
                    }, 'Normal')

                try {
                    # Start the process and wait for it to finish
                    $script:process = New-Object System.Diagnostics.Process
                    $script:process.StartInfo.FileName = $filePath
                    $script:process.StartInfo.Arguments = $arguments
                    $script:process.StartInfo.UseShellExecute = $false
                    $script:process.StartInfo.RedirectStandardOutput = $true
                    $script:process.StartInfo.RedirectStandardError = $true

                    $script:process.Start()
                    $script:process.WaitForExit()

                    # Check the flag after the extraction process
                    if (!$sharedVariables["continueRunning"]) {
                        break
                    }

                    # Notify the user that the extraction is complete
                    $outputTextBox.Dispatcher.Invoke({
                            $outputTextBox.AppendText(" done.`n")
                        }, 'Normal')

                    # Update the progress bar
                    $progressBar.Dispatcher.Invoke({
                            $progressBar.Value = ($fileCount / $totalFiles) * 100
                            $progressText.Text = "$fileCount / $totalFiles"
                        }, 'Normal')
                }
                catch {
                    # Notify the user that the extraction failed
                    $outputTextBox.Dispatcher.Invoke({
                            $outputTextBox.AppendText(" failed.`nError: $_`n")
                        }, 'Normal')
                }
            }
        }
        # Notify the user that the extraction is complete
        $outputTextBox.Dispatcher.Invoke({
                $outputTextBox.AppendText("Extraction complete, extracted $fileCount files.`n")
                $exitButton.IsEnabled = $true
            }, 'Normal')
    })

# Handle the Stop button click
$stopButton.Add_Click({
        # Stop the process if it's running
        if ($null -ne $script:process) {
            $script:process.Kill()
            # Wait for the process to exit
            $script:process.WaitForExit()
        }

        # Set the flag to false in the shared hashtable
        $sharedVariables["continueRunning"] = $false

        $outputTextBox.AppendText("Script stopped manually.`n")
        $exitButton.IsEnabled = $true
    })

# Handle the Exit button click
$exitButton.Add_Click({
        $window.Close()
    })

# Start the script execution
$ps.BeginInvoke()

# Show the window
$window.ShowDialog()