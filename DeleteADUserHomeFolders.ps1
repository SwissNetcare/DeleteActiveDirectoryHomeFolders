/*
 * Copyright (C) 2024 Swiss Netcare Solutions GmbH
 *
 * This file is part of our Tool Sets - in this case "DeleteActiveDirectoryHomeFolders".
 *
 * DeleteActiveDirectoryHomeFolders is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * any later version. 
 * You are not allowed to remove our company's name and/or its website in any forks you create.
 *
 * DeleteActiveDirectoryHomeFolders is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with DeleteActiveDirectoryHomeFolders. If not, see <http://www.gnu.org/licenses/>.
 *
 * For more information, please visit Swiss Netcare via
 * www.swissnetcare.ch
 *
 * Please notice that we do not give any support on this repository or its files.
 */



# Load required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create the WPF GUI
function Show-DirectorySelectionForm {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Folders to Delete" Height="500" Width="600">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="10">
            <TextBox Name="FolderPath" Width="400" Margin="0,0,10,0"/>
            <Button Name="BrowseButton" Content="Browse" Width="100"/>
        </StackPanel>
        <ListBox Name="FolderList" Grid.Row="1" Margin="10" SelectionMode="Extended"/>
        <ProgressBar Name="ProgressBar" Grid.Row="2" Margin="10" Height="20" Minimum="0" Maximum="100" Value="0"/>
        <Button Name="DeleteButton" Content="Delete Selected" Grid.Row="3" Margin="10" HorizontalAlignment="Right"/>
    </Grid>
</Window>
"@

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $Form = [Windows.Markup.XamlReader]::Load($reader)

    # Get the elements
    $folderPath = $Form.FindName('FolderPath')
    $browseButton = $Form.FindName('BrowseButton')
    $folderList = $Form.FindName('FolderList')
    $deleteButton = $Form.FindName('DeleteButton')
    $progressBar = $Form.FindName('ProgressBar')

    # Define the event handler for the browse button click
    $browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $result = $folderBrowser.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $folderPath.Text = $folderBrowser.SelectedPath
            Update-FolderList $folderBrowser.SelectedPath
        }
    })

    # Function to update the folder list
    function Update-FolderList($path) {
        $folderList.Items.Clear()
        Get-ChildItem -Directory -Path $path | ForEach-Object {
            $folderList.Items.Add($_.FullName)
        }
    }

    # Function to export deletion details to a text file
    function Export-DeletionDetails($deletedFolderDetails) {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        $saveFileDialog.DefaultExt = "txt"
        $saveFileDialog.AddExtension = $true
        $result = $saveFileDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePath = $saveFileDialog.FileName
            $deletedFolderDetails | ForEach-Object {
                Add-Content -Path $filePath -Value "Path: $($_.Path) - Name: $($_.Name) - Size: $($_.Size) MB"
            }
            [System.Windows.MessageBox]::Show("Deletion details exported to: $filePath")
        }
    }

    # Function to update the progress bar
    function Update-ProgressBar($progressBar, $current, $total) {
        $progressBar.Value = ($current / $total) * 100
    }

    # Function to delete folder using robocopy and remove-item
    function Delete-Folder($path) {
        $emptyFolder = Join-Path -Path $env:TEMP -ChildPath "EmptyFolder"
        New-Item -ItemType Directory -Path $emptyFolder -Force | Out-Null

        robocopy $emptyFolder $path /MIR | Out-Null
        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop

        Remove-Item -Path $emptyFolder -Recurse -Force
    }

    # Define the event handler for the delete button click
    $deleteButton.Add_Click({
        $selectedItems = $folderList.SelectedItems
        if ($selectedItems.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one folder to delete.")
        } else {
            $result = [System.Windows.MessageBox]::Show("Are you sure you want to delete the selected folders?", "Confirmation", [System.Windows.MessageBoxButton]::YesNo)
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                $deletedFolders = @()
                $deletedFolderDetails = @()
                $failedFolders = @()
                $totalItems = $selectedItems.Count
                $progressBar.Maximum = $totalItems
                $progressBar.Value = 0

                # Show the blocking popup
                [xml]$blockingXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Deletion in Progress" Height="150" Width="300" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" WindowStyle="None">
    <Grid>
        <TextBlock Text="Deletion in progress, please wait..." VerticalAlignment="Center" HorizontalAlignment="Center" FontSize="14"/>
    </Grid>
</Window>
"@
                $reader = (New-Object System.Xml.XmlNodeReader $blockingXaml)
                $blockingForm = [Windows.Markup.XamlReader]::Load($reader)
                $blockingForm.Show()
                
                # Perform the deletion process
                for ($i = 0; $i -lt $totalItems; $i++) {
                    $item = $selectedItems[$i]
                    try {
                        $folder = Get-Item -Path $item
                        $folderSize = (Get-ChildItem -Path $item -Recurse | Measure-Object -Property Length -Sum).Sum
                        $folderDetail = [PSCustomObject]@{
                            Path = $item
                            Name = $folder.Name
                            Size = [math]::Round($folderSize / 1MB, 2)
                        }
                        Delete-Folder $item
                        $deletedFolders += $item
                        $deletedFolderDetails += $folderDetail
                    } catch {
                        $failedFolders += $item
                    }
                    # Update progress
                    Update-ProgressBar $progressBar ($i + 1) $totalItems
                }

                # Close the blocking form
                $blockingForm.Close()

                # Show the summary message
                $summaryMessage = "Deletion Summary:`n"
                if ($deletedFolders.Count -le 10) {
                    $summaryMessage += "Deleted folders:`n" + ($deletedFolders -join "`n") + "`n"
                } else {
                    $summaryMessage += "$($deletedFolders.Count) folders deleted.`n"
                    $exportResult = [System.Windows.MessageBox]::Show("Do you want to export the deletion details?", "Export Details", [System.Windows.MessageBoxButton]::YesNo)
                    if ($exportResult -eq [System.Windows.MessageBoxResult]::Yes) {
                        Export-DeletionDetails $deletedFolderDetails
                    }
                }
                if ($failedFolders.Count -gt 0) {
                    $summaryMessage += "`nFailed to delete folders:`n" + ($failedFolders -join "`n")
                }
                [System.Windows.MessageBox]::Show($summaryMessage, "Summary")
                Update-FolderList $folderPath.Text  # Refresh the folder list
            }
        }
    })

    # Show the form
    $Form.ShowDialog()
}

# Ensure script is running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "You need to run this script as an administrator."
    exit
}

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 4) {
    Write-Warning "PowerShell 4.0 or higher is required to run this script."
    exit
}

# Check .NET Framework version
$netVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release | Select-Object -ExpandProperty Release
if ($netVersion -lt 378389) {
    Write-Warning ".NET Framework 4.5 or higher is required to run this script."
    exit
}

# Run the GUI function
Show-DirectorySelectionForm
