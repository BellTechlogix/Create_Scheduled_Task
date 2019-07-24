<#
	Create_Scheduled_Task.ps1
	Created by - Kristopher C. Roy
	Created on - 24 Jul 2019
	Purpose - To provide a GUI driven function to help create scheduled tasks
#>

Function Get-DateTime([Parameter(Mandatory)][ValidateSet('Date','Time')][string[]]$type)
{
	Add-Type -AssemblyName System.Windows.Forms
	IF($type -eq 'Date'){
	# Main Form
	$mainForm = New-Object System.Windows.Forms.Form
	$font = New-Object System.Drawing.Font("Courier", 13)
	$mainForm.Text = "Pick Date"
	$mainForm.Font = $font
	$mainForm.ForeColor = "White"
	$mainForm.BackColor = "DarkBlue"
	$mainForm.Width = 320
	$mainForm.Height = 130

	# DatePicker Label
	$datePickerLabel = New-Object System.Windows.Forms.Label
	$datePickerLabel.Text = "Selected Date"
	$datePickerLabel.Location = "15, 10"
	$datePickerLabel.Height = 22
	$datePickerLabel.Width = 120
	$mainForm.Controls.Add($datePickerLabel)

	# DatePicker
	$datePicker = New-Object System.Windows.Forms.DateTimePicker
	$datePicker.Location = "140, 7"
	$datePicker.Width = "150"
	$datePicker.Format = [windows.forms.datetimepickerFormat]::custom
	$datePicker.CustomFormat = "dd/MM/yyyy"
	$mainForm.Controls.Add($datePicker)

	# OD Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = "15, 50"
	$okButton.ForeColor = "Black"
	$okButton.BackColor = "White"
	$okButton.Text = "OK"
	$okButton.add_Click({$mainForm.close()})
	$mainForm.Controls.Add($okButton)

	[void] $mainForm.ShowDialog()
	$datePicker.Text
    }
    IF($type -eq 'Time'){
    # Main Form
	$mainForm = New-Object System.Windows.Forms.Form
	$font = New-Object System.Drawing.Font("Courier", 13)
	$mainForm.Text = "Pick Time"
	$mainForm.Font = $font
	$mainForm.ForeColor = "White"
	$mainForm.BackColor = "DarkBlue"
	$mainForm.Width = 320
	$mainForm.Height = 160

	# TimePicker Label
	$TimePickerLabel = New-Object System.Windows.Forms.Label
	$TimePickerLabel.Text = "Selected Time"
	$TimePickerLabel.Location = "15, 45"
	$TimePickerLabel.Height = 22
	$TimePickerLabel.Width = 120
	$mainForm.Controls.Add($TimePickerLabel)

	# TimePicker
	$TimePicker = New-Object System.Windows.Forms.DateTimePicker
	$TimePicker.Location = "140, 42"
	$TimePicker.Width = "150"
	$TimePicker.Format = [windows.forms.datetimepickerFormat]::custom
	$TimePicker.CustomFormat = "HH:mm:ss"
	$TimePicker.ShowUpDown = $TRUE
	$mainForm.Controls.Add($TimePicker)

	# OD Button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = "15, 80"
	$okButton.ForeColor = "Black"
	$okButton.BackColor = "White"
	$okButton.Text = "OK"
	$okButton.add_Click({$mainForm.close()})
	$mainForm.Controls.Add($okButton)

	[void] $mainForm.ShowDialog()
	$TimePicker.Text
    }
}

Function New-Scheduledtask($Name,[Parameter()][ValidateSet('Hourly','Daily','Weekly','Monthly','AtStartup','AtLogon')][string[]]$Frequency,
[Parameter()][ValidateSet('CMD','Powershell','Program')][string[]]$Type)
{
	$Time = New-ScheduledTaskTrigger -At ""
}