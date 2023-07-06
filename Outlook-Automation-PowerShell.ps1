Add-Type -AssemblyName System.Windows.Forms

$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

## 
# Creating new form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Input Form'
$form.Size = New-Object System.Drawing.Size(500,500) # Changed size here
$form.StartPosition = 'CenterScreen'

# Adding subject label and textbox
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Location = New-Object System.Drawing.Point(10,20)
$subjectLabel.Size = New-Object System.Drawing.Size(460,20) 
$subjectLabel.Text = 'Subject:'
$form.Controls.Add($subjectLabel)

$subjectBox = New-Object System.Windows.Forms.TextBox
$subjectBox.Location = New-Object System.Drawing.Point(10,40)
$subjectBox.Size = New-Object System.Drawing.Size(460,20) 
$form.Controls.Add($subjectBox)

# Adding body label and textbox
$bodyLabel = New-Object System.Windows.Forms.Label
$bodyLabel.Location = New-Object System.Drawing.Point(10,70)
$bodyLabel.Size = New-Object System.Drawing.Size(460,20) 
$bodyLabel.Text = 'Body:'
$form.Controls.Add($bodyLabel)

$bodyBox = New-Object System.Windows.Forms.TextBox
$bodyBox.Location = New-Object System.Drawing.Point(10,90)
$bodyBox.Size = New-Object System.Drawing.Size(460,300)
$bodyBox.Multiline = $true 
$form.Controls.Add($bodyBox)

# Adding OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10,400)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$form.Topmost = $true

$form.Add_Shown({$form.Activate()})
$result = $form.ShowDialog()

##

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $emailSubject = $subjectBox.Text
    $emailBody = $bodyBox.Text
    # Convert body to HTML
    $emailBody = $emailBody -replace "`n", "<br />"

    $dialogResult = $folderBrowser.ShowDialog()
    if ($dialogResult -eq 'OK') {
        $filesDirectory = $folderBrowser.SelectedPath

        # Get the list of all PDF files in the directory
        $files = Get-ChildItem -Path $filesDirectory -Filter *.pdf

        foreach ($file in $files) {
            # Extract username from the file name (removing '.pdf')
            $username = [IO.Path]::GetFileNameWithoutExtension($file.Name)
            $emailTo = "$username@domain.com"

            # using the existing Outlook instance on your machine, and hence the existing user session.
            $Mail = $Outlook.CreateItem(0)
            $Mail.Recipients.Add($emailTo)
            $Mail.Subject = $emailSubject
            $Mail.Body = $emailBody
            $Mail.Attachments.Add($file.FullName)
            $Mail.Send()
        }
    }
}