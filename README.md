# Lenovo Driver Extraction Tool

This PowerShell script is designed to streamline the extraction of Lenovo Driver Packages, which can be particularly useful for tasks such as preparing drivers for integration with software like NTLite. The script automates the extraction process, ensuring efficiency and consistency, especially when dealing with a large number of driver packages.

## Disclaimer

Please note that this tool is not officially affiliated or endorsed by Lenovo or any of its subsidiaries or affiliates. This is one of my first significant PowerShell projects, and I am open to any suggestions or feedback.

## Technical Details

The script operates by executing each driver package with specific parameters to automate and control the extraction process. The general form of the command executed for each driver package is as follows:

```plaintext
[driverpackage.exe] "/VERYSILENT /DIR=[extractiondir] /EXTRACT=YES"
```

- `"/VERYSILENT"`: This parameter ensures that the extraction process runs silently, meaning that it does not display any user interface or prompts from the driver package installer.
- `"/DIR=[extractiondir]"`: This parameter specifies the directory where the driver files will be extracted. The `[extractiondir]` placeholder is dynamically replaced by the actual path determined at runtime, based on user input or default settings.
- `"/EXTRACT=YES"`: This parameter instructs the driver package to perform the extraction of files rather than proceeding with a full installation.

## Optional Parameters

- `-ExtractDir [string]`: This specifies the folder where files will be extracted. If not provided, it defaults to a subdirectory named "extracted" within the driver directory.
- `-DriverDir [string]`: This defines the directory containing the driver files. If left unspecified, the script will use its own location as the default.

## Usage

To use this script, open PowerShell, navigate to the script's directory, and enter the following command:

```powershell
.\scriptName.ps1 -ExtractDir "Path\To\Extract\Directory" -DriverDir "Path\To\Driver\Directory"
```

### 1. Administrative Privileges and UAC Prompt

Running the script first involves a check for administrative privileges, which, while not mandatory, can make your experience smoother.

- **If not running as an administrator**: A User Access Control (UAC) prompt will show up, giving you these options:
  - **Yes**: Grant administrative privileges to the script, preventing multiple UAC prompts during the extraction of driver packages.
  - **No**: Proceed without administrative privileges, though this might trigger separate UAC prompts for each driver package.
  - **Cancel**: Stop and exit the script entirely.

### 2. Driver Directory Selection

Following this, the script will check if you've set the `-DriverDir` parameter:

- **If not set**: You'll see a prompt where you can:
  - **OK**: Confirm and use the script’s current location as the driver directory.
  - **Browse**: Choose a different folder for your driver files.

### 3. Extraction Directory Selection

Next, the script checks the `-ExtractDir` parameter:

- **If not set**: Another prompt will appear, allowing you to:
  - **OK**: Use the default directory, which is where your driver packages are located.
  - **Browse**: Select a different folder to extract your files to.

### 4. Execution and Output Window

Once all necessary folders are selected, the script starts extracting the files. An output window will appear, featuring options to:

- **Stop**: Halt the extraction process at any point.
- **Exit**: Close the output window once extraction is finished or manually stopped.

## License
This project is open-sourced under the MIT License. For more details, see the [LICENSE](LICENSE) file.
