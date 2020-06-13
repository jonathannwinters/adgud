<#  

┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ Active Directory Group User Details powershell script                                       │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 2020.06.12                                                                  |
│   AUTHOR      : __Jonathan N. Winters, jnw25@cornell.edu__                                  |
│   DESCRIPTION : Script that Queries a Microsoft AD Server for All Groups their User Details |
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
 
#>

# CHANGE DEPARTMENT NAMES LIST and the QUERY STRING For your Specific Active Directory
$Department_Names = "[Departmnet1]", "[Department2]"

#FOLLOWING 5 lines are UNFINISHED, need some work, tested working with MANUAL string
#Create ARRAY of OUs
#Create Array of DCs


Function Get_Group_Names($Group_Name){
 
#Modify Query to be For Each OU in OUs and each DC in DCs make following string
$query = dsquery group OU="GROUPS,OU=$Group_Name,OU=[OUNAME],OU=[OUNAME2],DC=[DC1],DC=[DC2],DC=[DC3]" 


$Group_Names =  New-Object System.Collections.Generic.List[System.String]

    ForEach ($Line in $($query -split "`r`n"))
    {
        $start = $Line.IndexOf("CN=") + 3 
        $end = $Line.IndexOf(",") - 4
        $GroupName = $Line.Substring($start,$end)
        $Group_Names += $GroupName
    }

    return $Group_Names

}







Function Print_Group_Names($Group_Names){
 
    ForEach ($Group in $Group_Names )
    {
        Write-host $Group
    }

}

Function Print_Groups_Users_Details(){
    ForEach ($Group in $Group_Names )
    {
       Write-host "------------------------------------------------------"
       Write-host $Group "has members:"
       Write-host "------------------------------------------------------"
       # dsquery group -samid "DL FAS PSY WestLabBerkWeinRas" | dsget group -members -expand | dsget user -samid -display -email
       # Write-host $Group
       dsquery group -samid "$Group" | dsget group -members -expand | dsget user -samid -display -email
    }
}

Function Get_Groups_Users_Details($Group){
      $Users = dsquery group -samid "$Group" | dsget group -members -expand | dsget user -samid -display -email
    return $Users
}


Add-Type -assembly System.Windows.Forms

$main_form = New-Object System.Windows.Forms.Form

$main_form.Text ='AD Group User Details'

$main_form.Width = 600

$main_form.Height = 400

$main_form.AutoSize = $true




$Label = New-Object System.Windows.Forms.Label

$Label.Text = "Department: "

$Label.Location  = New-Object System.Drawing.Point(0,10)

$Label.AutoSize = $true

$main_form.Controls.Add($Label)




$ComboBox = New-Object System.Windows.Forms.ComboBox

$ComboBox.Width = 300



ForEach ($Department_Name in $Department_Names)
{

$ComboBox.Items.Add($Department_Name)

}

$ComboBox.Location  = New-Object System.Drawing.Point(80,10)


$main_form.Controls.Add($ComboBox)




$ComboBox2 = New-Object System.Windows.Forms.ComboBox

$ComboBox2.Width = 300

$ComboBox2.Location  = New-Object System.Drawing.Point(100,40)

$main_form.Controls.Add($ComboBox2)




$Label2 = New-Object System.Windows.Forms.Label

$Label2.Text = "Available Groups:"

$Label2.Location  = New-Object System.Drawing.Point(0,40)

$Label2.AutoSize = $true

$main_form.Controls.Add($Label2)

$Label3 = New-Object System.Windows.Forms.Label

$Label3.Text = "Results of Query: "

$Label3.Location  = New-Object System.Drawing.Point(0,62)

$Label3.AutoSize = $true

$main_form.Controls.Add($Label3)





$Button = New-Object System.Windows.Forms.Button

$Button.Location = New-Object System.Drawing.Size(410,10)

$Button.Size = New-Object System.Drawing.Size(120,23)

$Button.Text = "Get Groups"

$main_form.Controls.Add($Button)


$Button.Add_Click(
    {
    $Combobox2.items.clear()
    $Choice = $ComboBox.SelectedItem.ToString()
    $Group_Names = Get_Group_Names($Choice)
    #Print_Group_Names($Group_Names)
    ForEach ($Group_Name in $Group_Names){
        $TextBox.Text += $Group_Name +"`r`n"
        $ComboBox2.Items.Add($Group_Name)
    } 
    #$Label3.Text =   $Group_Names #$Choice  #"TEST" # [datetime]::FromFileTime((Get-ADUser -identity $ComboBox.selectedItem -Properties pwdLastSet).pwdLastSet).ToString('MM dd yy : hh ss')
    }

)







$Button2 = New-Object System.Windows.Forms.Button

$Button2.Location = New-Object System.Drawing.Size(410,40)

$Button2.Size = New-Object System.Drawing.Size(120,23)

$Button2.Text = "Get User Details"

$main_form.Controls.Add($Button2)


$Button2.Add_Click({
    $TextBox.clear()
    $Choice2 = $ComboBox2.SelectedItem.ToString()
    $TextBox.Text += "Selected Group: " + $Choice2 +"`r`n"
    $Users =Get_Groups_Users_Details($Choice2)
    ForEach ($User in $Users){
        $TextBox.Text += $User +"`r`n"
    } 
})







$TextBox = New-Object System.Windows.Forms.TextBox

$TextBox.Width = 600
$TextBox.Height = 600

$TextBox.Location  = New-Object System.Drawing.Point(10,80)
$TextBox.MultiLine = $True

$main_form.Controls.Add($TextBox)


$main_form.ShowDialog()
