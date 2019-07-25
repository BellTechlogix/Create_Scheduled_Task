<#
	Create_Scheduled_Task.ps1
	Created by - Kristopher C. Roy
	Created on - 24 Jul 2019
	Purpose - To provide a GUI driven function to help create scheduled tasks
#>

#Function to create a scheduled task in task scheduler
Function New-Scheduledtask($Name,[Parameter()][ValidateSet('Daily','Weekly','Monthly','AtStartup','AtLogon')][string[]]$Frequency,
[Parameter()][ValidateSet('CMD','PowerShell','Program')][string[]]$Type)
{
	#Acknowledge Box Function
	Function AckBox([Parameter()][ValidateSet('Error','Question','Exclamation','Information')][string[]]$type,
	[Parameter()][ValidateSet('OK','OKCancel','YesNo','YesNoCancel')][string[]]$button,
	$title,$body
	){[System.Windows.MessageBox]::Show($body,$title,$button,$type)}

	#Typed Inputbox Function
	Function InputBox($title,$body,$default){Add-Type -AssemblyName Microsoft.VisualBasic
	Add-Type -AssemblyName PresentationCore,PresentationFramework
	[Microsoft.VisualBasic.Interaction]::InputBox($body, $title, $default)}

	#Function to grab a folder location from GUI
	function Get-Folder {
		param([string]$Description="Select Folder to place results in",[string]$RootFolder="Desktop")

	 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
		 Out-Null     

	   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
			$objForm.Rootfolder = $RootFolder
			$objForm.Description = $Description
			$Show = $objForm.ShowDialog()
			If ($Show -eq "OK")
			{
				Return $objForm.SelectedPath
			}
			Else
			{
				Write-Error "Operation cancelled by user."
			}
	}

	#select a file Function
	function Get-FileName
	{
	  param(
		  [Parameter(Mandatory=$false)]
		  [string] $Filter,
		  [Parameter(Mandatory=$false)]
		  [switch]$Obj,
		  [Parameter(Mandatory=$False)]
		  [string]$Title = "Select A File"
		)
 
		[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
		$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	  $OpenFileDialog.initialDirectory = $initialDirectory
	  $OpenFileDialog.FileName = $Title
	  #can be set to filter file types
	  IF($Filter -ne $null){
	  $FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
		$OpenFileDialog.filter = $FilterString}
	  if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
	  $OpenFileDialog.filter = $Filter
	  }
	  $OpenFileDialog.ShowDialog() | Out-Null
	  ## dont bother asking, just give back the object
	  IF($OBJ){
	  $fileobject = GI -Path $OpenFileDialog.FileName.tostring()
	  Return $fileObject
	  }
	  else{Return $OpenFileDialog.FileName}
	}

	#Function to grab date time from selection
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
		$datePicker.CustomFormat = "MM/dd/yyyy"
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
	
	#Get Credentials
	AckBox -type Information -button OK -title "Credentials" -body "The next prompt will be for the credentials that the Task will be run with"
	$User = Get-Credential
	
	#Define Triggers for scheduling types
	If($Frequency -ne 'AtStartup' -or $Frequency -ne 'AtLogon')
	{
		#initiates Get-DateTime function to grab date and time and store for use in triggers
		$time = Get-DateTime -type Time
		$date = Get-DateTime -type Date

		#If($Frequency -eq 'Hourly')
		#{
		#	$trigger = New-ScheduledTaskTrigger -At "$date $time" -Once -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration(0)
		#}
		
		If($Frequency -eq 'Daily')
		{
			$trigger = New-ScheduledTaskTrigger -At "$date $time" -Daily
		}
		If($Frequency -eq 'Weekly')
		{
			$day = (Get-Date $date).DayOfWeek
			$trigger = New-ScheduledTaskTrigger -At "$date $time" -Weekly -WeeksInterval 1 -DaysOfWeek $day
		}
		If($Frequency -eq 'Monthly')
		{
			$day = (Get-Date $date).DayOfWeek
			$trigger = New-ScheduledTaskTrigger -At "$date $time" -Weekly -WeeksInterval 4 -DaysOfWeek $day
		}
	}
	If($Frequency -eq 'AtStartup')
	{
		$trigger = New-ScheduledTaskTrigger -AtStartup
	}

		If($Frequency -eq 'AtLogon')
	{
		$trigger = New-ScheduledTaskTrigger -AtLogOn
	}

	#Define Actions for task to accomplish
	If($Type -eq 'CMD')
	{
		$argument = InputBox -title "CMD Argument" -body "Enter CMD Arguments" -default "\argument1 \argument2"
		$startin = Get-Folder
		$Action = New-ScheduledTaskAction -Execute "CMD.exe" -Argument $argument -WorkingDirectory $startin
	}
		If($Type -eq 'PowerShell')
	{
		$argument = InputBox -title "PowerShell Argument" -body "Enter PowerShell Arguments" -default "-argument1 -argument2"
		$startin = Get-Folder -Description "StartIn Folder"
		$Action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument $argument -WorkingDirectory $startin
	}
	IF($Type -eq 'Program')
	{
		$app = GI(Get-FileName)
		$Action = New-ScheduledTaskAction -Execute $app.name -Argument $argument -WorkingDirectory $app.Directory.FullName
	}

	#Create the Scheduled Task with Parameters
	Register-ScheduledTask -TaskName $Name -Trigger $trigger -Action $Action -User ($User.GetNetworkCredential().UserName) -Password ($User.GetNetworkCredential().Password) -RunLevel Highest -Force

	#Hourly Non-Functional - Possibly add at a later date
	#If($Frequency -eq 'Hourly')
	#{
    #    $task = Register-ScheduledTask -TaskName $Name -Trigger $trigger -Action $Action -User ($User.GetNetworkCredential().UserName) -Password ($User.GetNetworkCredential().Password) -RunLevel Highest -Force
    #    $task.Triggers.Repetition.Duration = "P1825D"
    #    $task.Triggers.Repetition.Interval = "PT1H"
    #    $task | Set-ScheduledTask	
	#}
}