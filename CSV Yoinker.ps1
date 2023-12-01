Add-Type -assembly System.Windows.Forms
$PSDefaultParameterValues['Import-Csv:Encoding'] = 'utf7'

$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
$selectedFile = "$scriptPath\*.csv"
$varStorage = @()
$varPlacement = 50
$yoink_Coords = 85
$form_height = 150
$expand_num = 1

function expand{
    $global:form_height += 40
    $global:yoink_Coords += 40
    $global:newField_Coords += 40

    $form.minimumSize = New-Object System.Drawing.Size(400,$form_height) 
    $form.maximumSize = New-Object System.Drawing.Size(400,$form_height)
    $form.Height = $form_height
    $global:yoinkButton.Location = New-Object System.Drawing.Size(200, $yoink_Coords)
    $global:newField.Location = New-Object System.Drawing.Size(100, $newField_Coords)
}

function setVar($name, $inputText){ # dynamically create text fields 
    $global:varPlacement += 40
    Set-Variable -Name $name -value (New-Object System.Windows.Forms.TextBox) -Scope Global
    $tmp = Get-Variable $name -ValueOnly
    $tmp.Location = New-Object System.Drawing.Size(10, $varPlacement)
    $tmp.Width = 150
    $tmp.Text = $inputText
    $global:form.Controls.Add($tmp)
    $global:varStorage += $name
    $global:expand_num += 1
}

function firstExpansion($name, $inputText){
    $checkCustom.Checked = $true
    # Expand GUI when checked
    $global:form_height += 40
    $global:yoink_Coords += 40
    $global:newField_Coords = 125
    $newField.Show()
    $form.minimumSize = New-Object System.Drawing.Size(400,$form_height) 
    $form.maximumSize = New-Object System.Drawing.Size(400,$form_height)
    $form.Height = $form_height
    $yoinkButton.Location = New-Object System.Drawing.Size(200, $yoink_Coords)

    # Remove checkmarks and disable other checkboxes
    $checkPhone.Enabled = $checkMail.Enabled = $checkRover.Enabled = $false
    $checkMail.Checked = $checkPhone.Checked = $checkRover.Checked = $false

    $newField.Location = New-Object System.Drawing.Size(100, $newField_Coords)
    $newField.Text = "Add field"
    $form.Controls.Add($global:newField)
    setVar "var_$global:expand_num" $inputText
}

function output{
    $string = $string -replace " ", ""

    if ($string.Length -gt 5){
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = 'CSV yoinker :)'
    $form2.Icon = '.\csv.ico'
    $form2.Width = 400
    $form2.Height = 400
    $form2.minimumSize = New-Object System.Drawing.Size(400,400) 
    $form2.maximumSize = New-Object System.Drawing.Size(400,400)
    $form2.MaximizeBox = $false

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Size(10,10)
    $textBox.Size = New-Object System.Drawing.Size(365,340)
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Vertical'
    $textBox.Text = $string
    $textBox.ReadOnly = $true
    $form2.Controls.Add($textBox)
    
    $form2.ShowDialog()
    } else {
        [System.Windows.Forms.MessageBox]::Show('Fant ikke valgt data i fil','Error :(')
    }
}

function displayHeaders($headers){

    $headerDisplayForm = New-Object System.Windows.Forms.Form
    $headerDisplayForm.Text = 'Headers :)'
    $headerDisplayForm.Icon = '.\csv.ico'
    $headerDisplayForm.Width = 300
    $headerDisplayForm.Height = 400
    $headerDisplayForm.minimumSize = New-Object System.Drawing.Size(300,400) 
    $headerDisplayForm.maximumSize = New-Object System.Drawing.Size(300,400)

    $headersListbox = New-Object System.Windows.Forms.ListBox
    $headersListbox.Location = New-Object System.Drawing.Point(10,20)
    $headersListbox.Size = New-Object System.Drawing.Size(260,320)
    $headerDisplayForm.Controls.Add($headersListbox)
    
    foreach($header in $headers){
        $headersListbox.Items.Add($header)
    }

    $headersListbox.Add_Click{
        if (-not $global:checkCustom.Checked){
            firstExpansion $global:expand_num $headersListbox.SelectedItem
            return
        }else{
            expand
            setVar "var_$global:expand_num" $headersListbox.SelectedItem
        }
    }
    $headerDisplayForm.ShowDialog()
}

$global:newField = New-Object System.Windows.Forms.Button
$yoink_Coords = 85
$form_height = 150
$form = New-Object System.Windows.Forms.Form
$form.Text = 'CSV yoinker :)'
$form.Width = 400
$form.Icon = ".\csv.ico"
$form.Height = $form_height
$form.minimumSize = New-Object System.Drawing.Size(400,$form_height) 
$form.maximumSize = New-Object System.Drawing.Size(400,$form_height)
$form.MaximizeBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(10,6)
$label.Text = "File:"
$label.Width = 50
$label.Height = 15
$form.Controls.Add($label)

$yoinkButton = New-Object System.Windows.Forms.Button
$yoinkButton.Location = New-Object System.Drawing.Size(155, $yoink_Coords)
$yoinkButton.Text = 'Yoink'
$form.Controls.Add($yoinkButton)

$fileDisplay = New-Object System.Windows.Forms.TextBox
$fileDisplay.Location = New-Object System.Drawing.Size(10,25)
$fileDisplay.Width = 275
$fileDisplay.Text =  $selectedFile
$form.Controls.Add($fileDisplay)

$fileButton = New-Object System.Windows.Forms.Button
$fileButton.Location = New-Object System.Drawing.Size(290, 23)
$fileButton.Width = 90
$fileButton.Text = 'Select'
$form.Controls.Add($fileButton)

$checkPhone = New-Object System.Windows.Forms.CheckBox
$checkPhone.Location = New-Object System.Drawing.Size(10, 50)
$checkPhone.Width = 61
$checkPhone.Text = 'Telefon'
$form.Controls.Add($checkPhone)

$checkMail = New-Object System.Windows.Forms.CheckBox
$checkMail.Location = New-Object System.Drawing.Size(80, 50)
$checkMail.Text = "Epost"
$checkMail.Width = 55
$form.Controls.Add($checkMail)

$checkRover = New-Object System.Windows.Forms.CheckBox
$checkRover.Location = New-Object System.Drawing.Size(140, 50)
$checkRover.Text = "Rover"
$checkRover.Width = 55
$form.Controls.Add($checkRover)

$checkCustom = New-Object System.Windows.Forms.CheckBox
$checkCustom.Location = New-Object System.Drawing.Size(200, 50)
$checkCustom.Text = "Custom"
$checkCustom.Width = 70
$form.Controls.Add($checkCustom)

$headerButton = New-Object System.Windows.Forms.Button
$headerButton.Location = New-Object System.Drawing.Size(290, 50)
$headerButton.Width = 90
$headerButton.Text = "View headers"
$form.Controls.Add($headerButton)


$headerButton.Add_Click({
    $file = Import-Csv -Path $selectedFile
    $headers = $file | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
    displayHeaders $headers
})


$newField.Add_Click({
    expand
    setVar "var_$global:expand_num" ""
})

$checkCustom.Add_Click({
    if($checkCustom.Checked){
        firstExpansion "var_$global:expand_num" ""
    } else {
        # Revert to original size
        $global:yoink_Coords = 85
        $global:form_height = 150

        $checkPhone.Enabled = $checkMail.Enabled = $checkRover.Enabled = $true

        $form.minimumSize = New-Object System.Drawing.Size(400,$form_height) 
        $form.maximumSize = New-Object System.Drawing.Size(400,$form_height)
        $yoinkButton.Location = New-Object System.Drawing.Size(155, $yoink_Coords)

        # Remove added fields and reset list & counter
        foreach ($var in $varStorage){
            $tmp = Get-Variable $var -ValueOnly
            $tmp.Dispose()

        $global:varStorage = @()
        $global:expand_num = 1
        $global:varPlacement = 50
        }
    }
})

$fileButton.Add_Click({
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'CSV Files (*.csv)|*.csv'}
    $FileBrowser.ShowDialog() | Out-Null

    if($fileBrowser.FileName -ne ""){
        $fileDisplay.Text = $global:selectedFile = $fileBrowser.FileName
    }
})


$yoinkButton.Add_Click({
    if (-not $checkPhone.Checked -and -not $checkMail.Checked -and -not $checkRover.Checked -and -not $checkCustom.Checked){
        [System.Windows.Forms.MessageBox]::Show('Velg noe for yoinking!','Error :(')
        return
    }

    try{
        # Importer valgt CSV-fil
        $file = Import-Csv -Path $selectedFile
        $string = ""

        if ($checkCustom.Checked){
            foreach ($var in $varStorage){
                $var = Get-Variable $var -ValueOnly
                if($var.Text.Length -gt 0){
                    foreach($item in $file){
                        $find = $var.text
                        if ($item.$find.length -gt 1){
                        $string += $item.$find + ';'
                        }
                    }
                }
            }
        }

        # Fetch number and email, separate with ;
        foreach ($item in $file)
        {
            if ($checkPhone.Checked){
                if ($item.'Telefon'.Length -gt 1){
                    $string += $item.'Telefon' + ';'
            }}

            if ($checkMail.Checked){
                if ($item.'E-post'.Length -gt 1){
                    $string += $item.'E-post' + ';'
            }}
            
            if ($checkRover.Checked){
                if($item.'phone_number'.Length -gt 1){
                    $string += $item.'phone_number' + ";"}
                }
            }
            output
        }Catch{
        [System.Windows.Forms.MessageBox]::Show('Kan kun yoinke en fil i gangen. Vennligst velg en via menyen.','Error :(')
    }}
)

$form.ShowDialog()