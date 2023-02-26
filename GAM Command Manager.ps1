<#

GAM Command Manager | STABLE

Made by Iri5s@Iri5s.com
Licensed with Apache 2.0
Github Repo: https://github.com/Iri5s/GAM-Command-Manager

#>

<# Starting Argument #>
param([switch] $Help)


<# CHANGE OPTIONS HERE #>

    # General GAM Settings
        $global:GAMLocalLocation = "" # The local folder of GAM on the server. EG. C:\GAM\ or D:\GAM. Do not use UNC paths here
        $global:GAMServerHostName = "" # NAME OF GAM SERVER. EG. Google-Server
        $global:DomainEmail = "" # Your domain name for email. EG. google.com (Do not include @, this is done later)
    # Banned users settings
        # If you have a banned group to stop sending an email or accessing a service, add the options here
        $global:ADGroup = "" # The name of the group on AD EG. 'BannedUsers'
        $global:GoogleGroup = "" + "@$DomainEmail" # The name of the Google Group. EG 'BannedUsers@google.com'

<# END OF OPTIONS #>

<# Global Variables #>

$global:Output =""

<# Pre Initialization #>

$ScriptVersion = "STABLE 1.0"
$VerbosePreference = "Continue"

<# Help Section #>

if ($Help)
{
    write-output "******************** GAM Command Manager | Version $ScriptVersion | By Iri5s @ Iri5s.com ********************" 
    write-output "*************************************************************************************************************"
    write-output "******* This script acts as a GUI for GAM (Google Apps Manager) to help manage your Google workspace ********"
    write-output "****************************** Find GAM here: https://github.com/GAM-team/GAM *******************************"
    write-output "*************************************************************************************************************"
    write-output "*********************************************** Help Section ************************************************"
    write-output "*************************************************************************************************************"

    write-output "- Settings like 'GAMLocation' can be saved directly at the top of the script."
    write-output "- Some features are marked [WIP], so are not implemented, please bare this in mind."
    write-output "- This script is an interface with GAM's command line, your GAM must be configured correctly for this to work."
    write-output "- If you encounter any bugs or feature suggestions, please submit an issue on Github."
    exit
}

<# Menu Functions #>

function Show-Menu {
    param (
        [string]$Title = "$Title"
        )
    Clear-Host

    Write-Host "================ $Title ================" -ForegroundColor Cyan

    if ($Option1) {Write-Host "1: Press '1' for $Option1" -ForegroundColor Yellow}
    if ($Option2) {Write-Host "2: Press '2' for $Option2" -ForegroundColor Yellow}
    if ($Option3) {Write-Host "3: Press '3' for $Option3" -ForegroundColor Yellow}
    if ($Option4) {Write-Host "4: Press '4' for $Option4" -ForegroundColor Yellow}
    if ($Option5) {Write-Host "5: Press '5' for $Option5" -ForegroundColor Yellow}
    if ($Option6) {Write-Host "6: Press '6' for $Option6" -ForegroundColor Yellow}
    if ($Option7) {Write-Host "7: Press '7' for $Option7" -ForegroundColor Yellow}
    if ($Option8) {Write-Host "8: Press '8' for $Option8" -ForegroundColor Yellow}
    Write-Host "M: Press 'M' to go to menu" -ForegroundColor Cyan
    Write-Host "Q: Press 'Q' to quit." -ForegroundColor Cyan
}

function Main-Menu {
    param (
        [string]$Title="Main Menu"
        )
    PreMenu
    Clear-Variable Option*
    $Title = "Main Menu"
    $Option1 = "Google Classroom Management"
    $Option2 = "Gmail Management"
    $Option3 = "Calendar Management"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {ClassroomMainMenu} 
            '2' {GmailManagementMenu}
            '3' {CalendarManagementMenu}
            'Q' {Exit}
        }

    }
    until ($selection -eq'Q')
}


function ClassroomMainMenu {
    param (
        [string]$Title="Google Classroom Management"
        )
    PreMenu
    Clear-Variable Option*
    $Title = "Google Classroom Management"
    $Option1 = "Class Creation and Deletion"
    $Option2 = "Adding/Removing Teachers"
    $Option3 = "Adding/Removing Students"
    $Option4 = "General Info"
    $Option5 = "Course Admin"
    $Option6 = "Misc"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {CreateAndDeleteMenu} 
            '2' {AddingAndRemovingTeachersMenu}
            '3' {AddingAndRemovingStudentsMenu}
            '4' {GeneralInfoMenu}
            '5' {CourseAdminMenu}
            '6' {MiscMenu}
            'M' {Main-Menu}
            'Q' {Exit}
        }

    }
    until ($selection -eq'Q')
}

function GmailManagementMenu {
    param (
        [string]$Title="Google Mail Management"
        )
    PreMenu
    Clear-Variable Option*
    $Title = "Gmail Management Menu"
    $Option1 = "Delegate Menu"
    $Option2 = "Manage Email Bans"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {GmailDelegateMenu} 
            '2' {BannedUsersMenu}
            'M' {Main-Menu}
            'Q' {Exit}
        }
    }
    until ($selection -eq'Q')
}

function CalendarManagementMenu {
    param (
        [string]$Title="Calendar Management"
        )
    PreMenu
    Clear-Variable Option*
    $Title = "Calendar Management Menu"
    $Option1 = "Delegate Menu"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {CalendarDelegateMenu} 
            'M' {Main-Menu}
            'Q' {Exit}
        }
    }
    until ($selection -eq'Q')
}


#// GOOGLE CLASSROOM FUNCTIONS // #

function CreateAndDeleteMenu {
    param (
        [string]$Title="Create/Delete Menu"
        )
    PreMenu
    Clear-Variable Option*
    $Title = "Create/Delete Menu"
    $Option1 = "Create a class"
    $Option2 = "Delete a class"
    $Option3 = "Main Menu"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {# Create Class
                $again = "Y"

                while ($again.ToLower() -eq "y")
                {
                    $Alias = Read-Host "Please type in name of class"
                    $TeacherR = Read-Host "Please type in teacher's user name"
                    $Teacher = "$TeacherR@$DomainEmail"
                    ./GAM create course alias $Alias teacher $Teacher name $Alias
                    ./GAM update course $Alias status Active
                    $again = Read-Host "Would you like to add another class?"
                }
                ClassroomMainMenu}
                '2' {# Deletes a class
                    $again = "Y"
                    while ($again.ToLower() -eq "y")
                    {
                        $Alias = Read-Host "Please type in ID of class"
                        CourseLookup
                        Write-Host "WARNING: You are about to delete class " -ForegroundColor Red -NoNewLine
                        Write-Host "$CourseName! " -ForegroundColor Yellow -NoNewLine
                        Write-Host "are you sure?" -ForegroundColor Red
                        $Confirm = Read-Host "Y/N"

                        if (($Confirm -eq "Y") -or ($Confirm -eq "y"))
                        {
                            ./GAM delete course $Alias
                            #Write-Host "Course $Alias deleted" -ForegroundColor Yellow
                        }
                        else { Write-Host "Command Aborted" -ForegroundColor Yellow}
                        $again = Read-Host "Would you like to delete another class?"
                    }
                        ClassroomMainMenu
                    }
                    '3' {ClassroomMainMenu}
                    'M' {ClassroomMainMenu}
                    'Q' {Exit}
                }

            }
            until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

        function AddingAndRemovingTeachersMenu {
            param (
                [string]$Title="Adding And Removing Teachers"
                )
            PreMenu
            Clear-Variable Option*
            $Title = "Adding And Removing Teachers"
            $Option1 = "Add a Teacher to a class"
            $Option2 = "Remove a Teacher from a class"
            Show-Menu
            do
            {
                $selection = Read-Host "Please select an option"
                switch ($selection)
                {
                    '1' {
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            Write-Host "Type 'csv' to run from csv file"
                            $Alias = Read-Host "Type in Course ID/Name"

                            if (($Alias -eq "CSV"))
                            {
                                $CSVPath = "$GAMLocalLocation\GAM_GUI\AddITeacher.txt"
                                $Header1 = "`"~Class`""
                                $Header2 = "`"~Email`""
                                .\GAM csv $CSVPath GAM course $Header1 add teacher $Header2
                            }
                            else{
                                $TeacherR = Read-Host "User name of Teacher to add"
                                $Teacher = "$TeacherR@$DomainEmail"
                                .\GAM course $Alias add teacher $Teacher
                            }
                            $again = Read-Host "Would you like to add another teacher?"
                        }
                        ClassroomMainMenu
                    }
                    '2' {
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            Write-Host "Type 'csv' to run from csv file"
                            $Alias = Read-Host "Type in Course ID/Name"

                            if (($Alias -eq "CSV"))
                            {
                                $CSVPath = "$GAMLocalLocation\GAM_GUI\RemoveITeacher.txt"
                                $Header1 = "`"~Class`""
                                $Header2 = "`"~Email`""
                                .\GAM csv $CSVPath GAM course $Header1 remove teacher $Header2
                            }
                            else{
                                $TeacherR = Read-Host "User name of Teacher to remove"
                                $Teacher = "$TeacherR@$DomainEmail"
                                .\GAM course $Alias remove teacher $Teacher
                            }
                            $again = Read-Host "Would you like to remove another teacher?"
                        }
                        ClassroomMainMenu
                    }
                    'M' {ClassroomMainMenu}
                    'Q' {Exit}
                }
            }
            until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

        function AddingAndRemovingStudentsMenu {
            param (
                [string]$Title="Add and Remove Students menu"
                )
            PreMenu
            Clear-Variable Option*
            $Title = "Add and Remove Students menu"
            $Option1 = "Add a Student"
            $Option2 = "Remove a Student"
            Show-Menu
            do
            {
                $selection = Read-Host "Please select an option"
                switch ($selection)
                {
                    '1' {
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            Write-Host "Type 'csv' to run from csv file: GAM\GAM_GUI\AddIStudent.txt"
                            $Alias = Read-Host "Type in Course ID/Name"

                            if (($Alias -eq "CSV"))
                            {
                                $CSVPath = "$GAMLocalLocation\GAM_GUI\AddIStudent.txt"
                                $Header1 = "`"~Class`""
                                $Header2 = "`"~Email`""
                                .\GAM csv $CSVPath GAM course $Header1 add student $Header2
                            }
                            else{
                                $StudentR = Read-Host "User name of student"
                                $Student = "$StudentR@$DomainEmail"
                                .\GAM course $Alias add student $Student
                            }
                            $again = Read-Host "Would you like to add another student?"
                        }
                        ClassroomMainMenu
                    } 
                    '2' {
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            Write-Host "Type 'csv' to run from csv file: GAM\GAM_GUI\RemoveIStudent.txt"
                            $Alias = Read-Host "Type in Course ID/Name"

                            if (($Alias -eq "CSV"))
                            {
                                $CSVPath = "$GAMLocalLocation\GAM_GUI\RemoveIStudent.txt"
                                $Header1 = "`"~Class`""
                                $Header2 = "`"~Email`""
                                .\GAM csv $CSVPath GAM course $Header1 add student $Header2
                            }
                            else{
                                $StudentR = Read-Host "User name of student"
                                $Student = "$StudentR@$DomainEmail"
                                .\GAM course $Alias remove student $Student
                            }
                            $again = Read-Host "Would you like to remove another student?"
                        }
                        ClassroomMainMenu
                    } 
                    #Could have an option to remove all students from a class... 
                    'M' {ClassroomMainMenu}
                    'Q' {Exit}
                }

            }
            until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

        function GeneralInfoMenu 
        {
            param (
                [string]$Title="General Info Menu"
                )
            PreMenu
            Clear-Variable Option*
            $Title = "General Info Menu"
            $Option1 = "Get Course Info"
            $Option2 = "List Teacher Courses"
            Show-Menu
            do
            {
                $selection = Read-Host "Please select an option"
                switch ($selection)
                {
                    '1' {   # Get Course Info
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            Write-Host "Type 'csv' to run from csv file: GAM\GAM_GUI\GetCourseInfo.txt"
                            $Alias = Read-Host "Type in Course ID"
                            $OutPath = "$GAMLocalLocation\GAM_GUI\CourseInfo.txt"

                            if ($Alias -eq "CSV")
                            {
                                $CSVPath = "$GAMLocalLocation\GAM_GUI\GetCourseInfo.txt"
                                $Header = "`"~ID`""
                                .\GAM csv $CSVPath GAM info course $Header > $OutPath 
                            }
                            else{
                                .\GAM info course $Alias > $OutPath
                            }
                            Write-Host "Done! Do you want to open '$OutPath'?" -ForegroundColor Yellow
                            Write-Host "Selecting No will display results in console!" -ForegroundColor Yellow
                            $Answer = Read-Host "Yes/No"

                            if (($Answer -eq "Y") -or ($Answer -eq "Yes"))
                            {OpenOutFile}
                            else {DisplayResults}
                            $again = Read-Host "Would you like to get info of another class?"
                        }
                        ClassroomMainMenu
                    }

                    '2' {   # List Teacher Courses
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            $TeacherR = Read-Host "Please type in teacher's user name"
                            $Teacher = "$TeacherR@$DomainEmail"
                            $global:OutPath = "$GAMLocalLocation\GAM_GUI\ITeacherCourses.txt"
                            ./GAM print courses teacher $Teacher > $OutPath
                            #Write-Host "Course $Alias created" -ForegroundColor Yellow

                            Write-Host "Done! Do you want to open '$OutPath'?" -ForegroundColor Yellow
                            Write-Host "Selecting No will display results in console!" -ForegroundColor Yellow
                            $Answer = Read-Host "Yes/No"

                            if (($Answer -eq "Y") -or ($Answer -eq "Yes"))
                            {OpenOutFile}
                            else {DisplayResults}

                            $again = Read-Host "Would you like to list another teachers classes?"
                        }
                        ClassroomMainMenu
                    } 
                    'M' {ClassroomMainMenu}
                    'Q' {Exit}
                }
            }        
            until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
        }

        function CourseAdminMenu 
        {
            param (
                [string]$Title="Course Admin Menu"
                )
            PreMenu
            Clear-Variable Option*
            $Title = "Course Admin Menu"
            $Option1 = "Set Class Status"
            $Option2 = "Change Class Owner"
            Show-Menu
            do
            {
                $selection = Read-Host "Please select an option"
                switch ($selection)
                {
                    '1' {   #Set Class Status
                        $again = "Y"
                        while ($again.ToLower() -eq "y")
                        {
                            $Alias = Read-Host "Type in Course ID"
                            CourseLookup
                            Write-Host "Current course $Alias state is" -NoNewLine
                            Write-Host " $CourseStatus" -ForegroundColor Yellow
                            if ($CourseStatus.ToLower() -eq "active")
                            {
                                Write-Host "Would you like to archive this classroom?"
                                $confirm = Read-Host "y/N"
                                if ($confirm.ToLower() -eq "y")
                                {
                                    .\GAM update course $Alias status "Archived"
                                }
                            }
                            elseif ($CourseStatus.ToLower() -eq "archived")
                            {
                                Write-Host "Would you like to make this classroom active?"
                                $confirm = Read-Host "y/N"

                                if ($confirm.ToLower() -eq "y")
                                {
                                    .\GAM update course $Alias status "Active"
                                }
                            }
                            $again = Read-Host "Would you like to update another class status?"
                        }
                        ClassroomMainMenu }
                        '2' {   #Change class owner
                            $again = "Y"
                            while ($again.ToLower() -eq "y")
                            {
                                $Alias = Read-Host "Type in Course ID"
                                $TeacherR = Read-Host "Teachers user name?"
                                $Teacher = "$TeacherR@$DomainEmail" 
                                .\GAM update course $Alias owner $Teacher
                                CourseLookup
                                #Write-Host "$TeacherR is now owner of '$CourseName'"
                                $again = Read-Host "Would you like to change another class owner?"
                            }
                            ClassroomMainMenu} 
                            'M' {ClassroomMainMenu}
                            'Q' {Exit}
                        }

                    }
                    until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
                }

                function MiscMenu 
                {
                    param (
                        [string]$Title="Misc Menu"
                        )
                    PreMenu
                    Clear-Variable Option*
                    $Title = "Misc Menu"
                    $Option1 = "Archive ALL Classes"
                    $Option2 = "Run Class Sync [WIP]"
                    Show-Menu
                    do
                    {
                        $selection = Read-Host "Please select an option"
                        switch ($selection)
                        { # Come back to this and change the $ImportCSV var to be a bit better optimized. 
                            '1' {   #Archive All Classes
                                $ResultsFile = "\\server73\c$\GAM\_Results\AllCoursesRAW.csv"
                                .\GAM print courses state active > "\\server73\c$\GAM\_Results\AllCoursesRAW.csv"
                                $CSVConvert = ConvertTo-Csv -inputobject $ResultsFile -NoTypeInformation #Grabs the csv file produced and correctly interprets the commas
                                $AllCourseName = Import-Csv $ResultsFile | select name | Select-Object -ExpandProperty name # This returns a list of names of all the courses. Now to filter..
                                $FilteredCourseName = $AllCourseName | Where-Object {($_ -like "*_2020*") -or ($_ -like "SS_*") -or ($_ -like "Tutor_*")}
                                $ImportCSV = Import-Csv $ResultsFile | export-csv "\\server73\c$\GAM\_Results\AllCourses.csv" -NoTypeInformation
                                #Remove-Item $ResultsFile
                                #Start-Sleep -s 5 
                                foreach ($Names in $FilteredCourseName)
                                {
                                    Write-Host $Names
                                    #.\GAM update course $Names state archived 
                                }
                                Write-Host "Operation Completed" -ForegroundColor Yellow
                                $Choice = Read-Host "Would you like to view current active classes? Y/N"
                                if (($Choice -eq "Y") -or ($Choice -eq "Yes"))
                                {
                                    $ResultsFile = "\\server73\c$\GAM\_Results\NewAllCoursesRAW.csv"
                                    .\GAM print courses state active > "\\server73\c$\GAM\_Results\NewAllCoursesRAW.csv"
                                    $CSVConvert = ConvertTo-Csv -inputobject $ResultsFile -NoTypeInformation 
                                    $NewAllCourseName = Import-Csv $ResultsFile | select name | Select-Object -ExpandProperty name
                                    $ImportCSV = Import-Csv $ResultsFile | export-csv "\\server73\c$\GAM\_Results\NewAllCourses.csv" -NoTypeInformation
                                    Write-Host "Current active courses" -ForegroundColor Yellow
                                    foreach ($NewCourseName in $NewAllCourseName)
                                    {
                                        Write-Host $NewCourseName
                                    }
                                    Remove-Item $ResultsFile
                                    Write-Host "End of active course list" -ForegroundColor Yellow
                                    Read-Host "Press 'enter' to continue.."
                                }
                                ClassroomMainMenu } 
                                '2' {   #Run Class sync
                                    # WIP
                                    ClassroomMainMenu} 
                                    'M' {ClassroomMainMenu}
                                    'Q' {Exit}
                                }

                            }
                            until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
                        }

#// GMAIL FUNCTIONS // #

function GmailDelegateMenu {
    param (
        [string]$Title="Gmail Delegate Menu"
        )
    PreMenu
    Clear-Variable Option*
    $Title = "Gmail Delegate Menu"
    $Option1 = "See delegates"
    $Option2 = "Add delegates"
    $Option3 = "Remove delegates"
    $Option4 = "Remove all your delegates [WIP]"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {# Show Delegate
                $again = "Y"
                while ($again.ToLower() -eq "y")
                {
                    $SharedEmail = Read-Host "Email (Accepts just user name)"
                    If ($SharedEmail -notcontains "@$DomainEmail")
                    {  $SharedEmail = "$SharedEmail@$DomainEmail" }
                    Write-Host "Listing delegates for: $SharedEmail" -ForegroundColor Yellow
                    ./GAM user $SharedEmail show delegates
                    Write-Host "All delegates displayed" -ForegroundColor Yellow
                    $again = Read-Host "Would you like to list another emails delegates?"
                }
                GmailDelegateMenu
            } 
            '2' {# Add Delegate
                $again = "Y"
                while ($again.ToLower() -eq "y")
                {
                    $SharedEmail = Read-Host "Shared Email (Accepts just user name)"
                    $DelegateToAdd = Read-Host "Delegate to Add User name"
                    If ($SharedEmail -notcontains "@$DomainEmail")
                    {  $SharedEmail = "$SharedEmail@$DomainEmail" }
                    ./GAM user $SharedEmail add delegate $DelegateToAdd
                    #Write-Host "Processed." -ForegroundColor Yellow
                    $again = Read-Host "Would you like to add another emails delegates?"
                }
                GmailDelegateMenu}
                '3' {# Remove Delegate
                    $again = "Y"
                    while ($again.ToLower() -eq "y")
                    {
                        $SharedEmail = Read-Host "Shared Email (Accepts just user name)"
                        $DelegateToAdd = Read-Host "Delegate to remove User name"
                        If ($SharedEmail -notcontains "@$DomainEmail")
                        {  $SharedEmail = "$SharedEmail@$DomainEmail" }
                        ./GAM user $SharedEmail delete delegate $DelegateToAdd
                        #Write-Host "Processed." -ForegroundColor Yellow
                        $again = Read-Host "Would you like to remove another emails delegates?"
                    }
                    GmailDelegateMenu}
                    '4' {# Remove all your Delegates
                        #$SharedEmail = Read-Host "Shared Email (Accepts just user name)"
                        #$DelegateToAdd = Read-Host "Delegate to Add User name"
                        #If ($SharedEmail -notcontains "@$DomainEmail")
                        #{  $SharedEmail = "$SharedEmail@$DomainEmail" }


                        GmailDelegateMenu}
                        'M' {GmailManagementMenu}
                        'Q' {Exit}
                    }

                }
                until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
            }

            function BannedUsersMenu {
                param (
                    [string]$Title="Banned Users Management"
                    )
                #$adminUsername = Read-Host "Your User name:"
                #$adminPassword = Read-Host "Your Password:" -AsSecureString
                #-Server $DomainController -Credential adminUsername, $Password | ConvertTo-SecureString -AsPlainText -Force

                #$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
                #$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

                Clear-Variable Option*
                $Title = "Banned Users Management"
                $Option1 = "Add a user to the ban list"
                $Option2 = "Remove a user from the ban list"
                $Option3 = "See banned users"
                Show-Menu
                do
                {
                    $selection = Read-Host "Please select an option"
                    switch ($selection)
                    {
                        '1' {
                            $username = Read-Host "User name"
                            $user = Get-ADGroupMember -Identity $ADGroup | Where-Object {$_.sAMAccountName -eq $username}
                            if($user){
                                Write-Host "$username is already a banned user! - Exiting" -ForegroundColor Blue 
                            }
                            else {
                                # Add to AD Group.
                                Write-Host "Adding $username to AD group.." -ForegroundColor Yellow
                                try {
                                    Add-ADGroupMember -Identity $ADGroup -Members $username
                                    Write-Host "Added $username to AD group.." -ForegroundColor Green
                                }
                                catch {Write-Host "Something went wrong when adding $username to AD Group: $ADGroup |" $Error[0] -ForegroundColor Red}

                                # Add to Google Group.
                                Write-Host "Adding $username to Google group.." -ForegroundColor Yellow
                                try {
                                    cd $GAMLocalLocation
                                    .\GAM update group $GoogleGroup add member "$username@$DomainEmail"
                                    Write-Host "Added $username@$DomainEmail to Google group.." -ForegroundColor Green
                                }
                                catch {Write-Host "Something went wrong when adding $username to Google Group: $GoogleGroup |" $Error[0] -ForegroundColor Red}
                                Write-Host "$username is now banned." -ForegroundColor Green
                            }
                            Read-Host "Press 'enter' to continue"
                            BannedUsersMenu
                        } 
                        '2' {
                            $username = Read-Host "User name"
                            $user = Get-ADGroupMember -Identity $ADGroup -Confirm $false | Where-Object {$_.sAMAccountName -eq $username} 
                            if(!$user){
                                Write-Host "$username is not a banned user! - Exiting" -ForegroundColor Blue 
                            }
                            else {
                                # Remove from AD Group.
                                Write-Host "Removing $username from AD group.." -ForegroundColor Yellow
                                try {
                                    Remove-ADGroupMember -Identity $ADGroup -Members $username
                                    Write-Host "Removed $username from AD group.." -ForegroundColor Green
                                }
                                catch {Write-Host "Something went wrong when removing $username from AD Group: $ADGroup |" $Error[0] -ForegroundColor Red}

                                # Remove from Google Group.
                                Write-Host "Removing $username from Google group.." -ForegroundColor Yellow
                                try {
                                    cd $GAMLocalLocation
                                    .\GAM update group $GoogleGroup delete member "$username@$DomainEmail"
                                    Write-Host "Removed $username@$DomainEmail from Google group.." -ForegroundColor Green
                                }
                                catch {Write-Host "Something went wrong when removing $username from Google Group: $GoogleGroup |" $Error[0] -ForegroundColor Red}
                                Write-Host "$username is now unbanned." -ForegroundColor Green
                            }
                            Read-Host "Press 'enter' to continue"
                            BannedUsersMenu
                        }
                        '3' {
                            Import-Module ActiveDirectory -Verbose:$false
                            $BannedUsers = Get-ADGroupMember -Identity $ADGroup -Recursive | Get-ADUser -Property SAMAccountName | Select-object -ExpandProperty sAMAccountName
                            Write-Host "Found" $BannedUsers.Count "banned users on AD." -ForegroundColor Cyan
                            Write-Host "=====================================================" -ForegroundColor Yellow
                            foreach ($i in $BannedUsers)
                            {
                                Write-Host "$i@$DomainEmail"
                            }
                            Write-Host "=====================================================" -ForegroundColor Yellow
                            Read-Host "Press 'enter' to continue"
                            BannedUsersMenu
                        }
                        'M' {GmailManagementMenu}
                        'Q' {Exit}
                    }

                }
                until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
            }

#// CALENDAR FUNCTONS // #

function CalendarDelegateMenu {
    param (
        [string]$Title="Calendar Delegate Menu"
        )
    Clear-Variable Option*
    $Title = "Gmail Delegate Menu"
    $Option1 = "List users calendars"
    $Option2 = "Add to calendar"
    $Option3 = "Remove from calendar"
    Show-Menu
    do
    {
        $selection = Read-Host "Please select an option"
        switch ($selection)
        {
            '1' {# List users calendars
                $again = "Y"
                while ($again.ToLower() -eq "y")
                {
                    $UserEmail = Read-Host "Show calendars for this email (Accepts just user name)"
                    If ($UserEmail -notcontains "@$DomainEmail")
                    {  $UserEmail = "$UserEmail@$DomainEmail" }
                    Write-Host "Listing calendars for: '$UserEmail'" -ForegroundColor Yellow
                    ./GAM user $UserEmail show calendars
                    Write-Host "All calendars displayed" -ForegroundColor Yellow
                    $again = Read-Host "Would you like to list another users calendars?"
                }
                CalendarDelegateMenu
            }
            '2' {# Add to calendar
                $again = "Y"
                while ($again.ToLower() -eq "y")
                {
                    $UserEmail = Read-Host "Email to add to calendar (Accepts just user name)"
                    $CalendarName = Read-Host "Calendars email (id)"
                    Write-Host "What permissions would you like them to have?" -ForegroundColor Blue 
                    Write-Host "1 - View free/busy Access" -ForegroundColor Yellow
                    Write-Host "2 - View all Access" -ForegroundColor Yellow
                    Write-Host "3 - Editor Access" -ForegroundColor Yellow
                    Write-Host "4 - Owner Access" -ForegroundColor Yellow
                    $permissionsInput = Read-Host "Please input a selection"
                    $permissionToHave = ""
                    switch ($permissionsInput)
                    {
                        '1' {
                            $permissionToHave = "freebusy"
                        }
                        '2' {
                            $permissionToHave = "read"
                        }
                        '3'
                        {
                            $permissionToHave = "editor"
                        }
                        '4'
                        {
                            $permissionToHave = "owner"
                        }
                    }
                    If ($UserEmail -notcontains "@$DomainEmail")
                    {  $UserEmail = "$UserEmail@$DomainEmail" }
                    Write-Host "Adding '$UserEmail' to calendar '$CalendarName'" -ForegroundColor Yellow
                    ./GAM calendar $CalendarName add $permissionToHave $UserEmail sendnotifications false
                    Write-Host "Showing calendar '$CalendarName' for '$UserEmail'" -ForegroundColor Yellow
                    ./GAM user $UserEmail add calendar $CalendarName selected true
                    Write-Host "'$UserEmail' added to calendar '$CalendarName' as a '$permissionToHave'" -ForegroundColor Yellow
                    $again = Read-Host "Would you like to add another user to a calendar?"
                }
                CalendarDelegateMenu
            }
            '3' {# Remove from calendar
                $again = "Y"
                while ($again.ToLower() -eq "y")
                {
                    $UserEmail = Read-Host "Email to remove from calendar (Accepts just user name)"
                    $CalendarName = Read-Host "Calendars email (id)"
                    If ($UserEmail -notcontains "@$DomainEmail")
                    {  $UserEmail = "$UserEmail@$DomainEmail" }
                    Write-Host "De-listing calendar '$CalendarName' from '$UserEmail'" -ForegroundColor Yellow
                    ./GAM user $UserEmail delete calendar $CalendarName
                    Write-Host "Removing '$UserEmail' from calendar '$CalendarName'" -ForegroundColor Yellow
                    ./GAM calendar $CalendarName delete user $UserEmail
                    Write-Host "'$UserEmail' removed from calendar '$CalendarName'." -ForegroundColor Yellow
                    $again = Read-Host "Would you like to remove another user from a calendar?"
                }
                CalendarDelegateMenu
            }
            'M' {CalendarManagementMenu}
            'Q' {Exit}
        }
    }
    until (($selection -eq '1') -or ($selection -eq'2') -or ($selection -eq'3') -or ($selection -eq'M') -or ($selection -eq'Q'))
}

<# Other functions on the backend #>

function CourseLookup
{
    param (
        [string]$Title="Course Lookup"
        )
    $TempOutPath = "$GAMLocalLocation\GAM_GUI\Temp.txt"
    .\GAM info course $Alias > $TempOutPath

    $CourseNameR = (Get-Content -Path "$TempOutPath" -TotalCount 2)[-1]
    $global:CourseName = $CourseNameR.Trim() -replace "name: ", "" 

    $CourseStatusR = (Get-Content -Path "$TempOutPath" -TotalCount 7)[-1]
    $global:CourseStatus = $CourseStatusR.Trim() -replace "courseState: ", "" 

}

function DisplayResults
{
    param (
        [string]$Title="Display Results"
        )
    Write-Host "Displaying results" -ForegroundColor Yellow
    if ($OutPath){
        Import-Csv $OutPath | Format-List}
        else {Write-Host "ERROR: Failed to find an Out file" -ForegroundColor Red} 
    }

    function OpenOutFile
    {
        param (
            [string]$Title="OpenOutFile"
            )
        Write-Host "Opening '$OutPath'" -ForegroundColor Yellow
        Invoke-Command -ScriptBlock {start notepad.exe "$OutPath"}

    }  

    function PreMenu
    {
        param (
            [string]$Title="PreMenu"
            )
        Clear-Variable Option* -Scope Global
        Clear-Variable Command* -Scope Global
        Clear-Variable Command* -Scope Global

    }  

    function CheckSettings {
        if ($GAMLocalLocation.Length -lt 2)
        {
            Write-Host "WARNING: GAMLocation is not set, please type this in now (You can save this setting directly at the top of the script)." -ForegroundColor Yellow
            Do {
                $GAMLocalLocation = Read-Host "Enter GAM folder location. e.g. 'C:\GAM' (No UNC paths)"
                $confirm = Read-Host "Confirm GAM folder location is:"$GAMLocalLocation"? y/N"

                if ($confirm.ToLower() -eq "y")
                {
                    break
                }
            }
            While ($true)
        }

        if ($GAMServerHostName.Length -lt 2)
        {
            Write-Host "WARNING: GAMServerName is not set, please type this in now (You can save this setting directly at the top of the script)." -ForegroundColor Yellow
            Do {
                $GAMServerHostName = Read-Host "Enter GAM Server Host name. e.g. 'Server99'"
                $confirm = Read-Host "Confirm your GAM server host name is:"$GAMServerHostName"? y/N"

                if ($confirm.ToLower() -eq "y")
                {
                    break
                }
            }
            While ($true)
        }
        if ($DomainEmail.Length -lt 2)
        {
            Write-Host "WARNING: DomainEmail is not set, please type this in now (You can save this setting directly at the top of the script)." -ForegroundColor Yellow
            Do {
                $domainInput = Read-Host "Enter your domain name. e.g. '@google.com'"
                if ($domainInput -like "@*")
                { $DomainEmail = "$domainInput" }
                else { $DomainEmail = "@$domainInput"}

                $confirm = Read-Host "Confirm your domain email is:"$DomainEmail"? y/N"

                if ($confirm.ToLower() -eq "y")
                {
                    break
                }
            }
            While ($true)
        }
    }

<# Checks & Script #>

$hostname = [System.Net.Dns]::GetHostName()
$Location = Get-Location # Store location of PS script to CD back to after
if ($GAMServerHostName.Length -lt 2)
{
    Write-Host "WARNING: GAMServerName is not set, please type this in now (You can save this setting directly at the top of the script)." -ForegroundColor Yellow
    Do {
        $GAMServerHostName = Read-Host "Enter GAM Server Host name. e.g. 'Server99'"
        $confirm = Read-Host "Confirm your GAM server host name is:"$GAMServerHostName"? y/N"

        if ($confirm.ToLower() -eq "y")
        {
            break
        }
    }
    While ($true)
}

if ($hostname -eq $GAMServerHostName)
{
    CheckSettings # Check to make sure settings are correct
    New-Item -itemType Directory -Path "$GAMLocalLocation" -Name "GAM_GUI" >$null
    cd $GAMLocalLocation
}
else{ 
    Write-Host "WARNING: You are currently running this script remotely, please run this script on your GAM server or through a PSSession." -ForegroundColor Red 
    exit
}

Main-Menu

# Once main menu function returns, CD back to script.

cd "$Location"