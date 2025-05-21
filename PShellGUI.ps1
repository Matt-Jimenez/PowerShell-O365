Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration ---
$ScriptDirectory = "C:\Temp" # Replace with the actual path to your scripts
$WindowTitle = "PowerShell Script Launcher"
$WindowWidth = 400
$WindowHeight = 300

# --- Create the Form ---
$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = $WindowTitle
$objForm.Width = $WindowWidth
$objForm.Height = $WindowHeight
$objForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# --- Create the ListBox to display scripts ---
$objListBox = New-Object System.Windows.Forms.ListBox
$objListBox.Location = New-Object System.Drawing.Point(10, 10)
$objListBox.Size = New-Object System.Drawing.Size(360, 200)
$objListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$objForm.Controls.Add($objListBox)

# --- Populate the ListBox with PowerShell scripts ---
Get-ChildItem -Path $ScriptDirectory -Filter "*.ps1" | ForEach-Object {
    $objListBox.Items.Add($_.Name)
}

# --- Create the Launch Button ---
$objButton = New-Object System.Windows.Forms.Button
$objButton.Location = New-Object System.Drawing.Point(10, 220)
$objButton.Size = New-Object System.Drawing.Size(100, 30)
$objButton.Text = "Launch Script"
$objButton.Add_Click({
    if ($objListBox.SelectedItem) {
        $selectedScript = Join-Path $ScriptDirectory $objListBox.SelectedItem
        Write-Host "Launching script: $selectedScript"
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", "`"$selectedScript`""
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a script to launch.", "Information")
    }
})
$objForm.Controls.Add($objButton)

# --- Create the Exit Button ---
$objExitButton = New-Object System.Windows.Forms.Button
$objExitButton.Location = New-Object System.Drawing.Point(270, 220)
$objExitButton.Size = New-Object System.Drawing.Size(100, 30)
$objExitButton.Text = "Exit"
$objExitButton.Add_Click({
    $objForm.Close()
})
$objForm.Controls.Add($objExitButton)

# --- Show the Form ---
$objForm.ShowDialog() | Out-Null