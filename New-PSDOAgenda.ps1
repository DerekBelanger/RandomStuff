# Load YAML data once and store in variables
# Maybe add handling file and formatting exceptions
# Another idea: -ExcludePresenter and -ExcludeSubject
$YamlSchedule = Get-Content -Path .\SummitSchedule.yaml -Raw | ConvertFrom-Yaml

$Presentations =@()
foreach ($presentation in $YamlSchedule) {
    $presentations += [pscustomobject]@{
        title       = $presentation.title
        when        = [datetime]$presentation.when
        presenters  = $presentation.presenters
        keywords    = $presentation.keywords
        interest    = 0
    }
}

# Register ArgumentCompleters for New-PSDOAgenda
# tried colapsing to one call to Register-ArgumentCompleters in a paramaterized funciton, but struggled with scope issues in the scriptblcok
function Register-ArgumentCompleters {
    Register-ArgumentCompleter -CommandName New-PSDOAgenda -ParameterName PriorityPresenters -ScriptBlock {
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
        foreach ($presenter in ($Presentations.presenters | Select-Object -Unique)) {
            if ($presenter -like "$wordToComplete*") {
                [System.Management.Automation.CompletionResult]::new($presenter, $presenter, 'ParameterValue', $presenter)
            }
        }
    }
    
    Register-ArgumentCompleter -CommandName New-PSDOAgenda -ParameterName PrioritySubjects -ScriptBlock {
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
        foreach ($keyword in ($Presentations.keywords | Select-Object -Unique)) {
            if ($keyword -like "$wordToComplete*") {
                [System.Management.Automation.CompletionResult]::new($keyword, $keyword, 'ParameterValue', $keyword)
            }
        }
    }
}

Register-ArgumentCompleters

# the main user function of this script
function New-PSDOAgenda {
    param (
        [string[]]$PriorityPresenters,
        [string[]]$PrioritySubjects
    )
    $Agenda =@()
    # validate the PriorityPresenters
    foreach ($presenter in $PriorityPresenters) {
        if ($Presentations.presenters -notcontains $presenter) {
            throw "Invalid presenter: $presenter"
        }
    }

    # validate the PrioritySubjects
    foreach ($subject in $PrioritySubjects) {
        if ($Presentations.keywords -notcontains $subject) {
            throw "Invalid subject: $subject"
        }
    }
    
    # calculate the level of interest in each presentation, add this value to each presentation object
    foreach ($presentation in $Presentations) {
        foreach ($PrioritySubject in $PrioritySubjects){
            if ($presentation.keywords -contains $PrioritySubject) {$presentation.interest ++}
        }
        foreach ($PriorityPresenter in $PriorityPresenters){
            if ($presentation.presenters -contains $PriorityPresenter) {$presentation.interest ++}
        }
    }
    
    # derive timeslots 
    $timeslots = $presentations.when | Select-Object -Unique
    foreach ($timeslot in $timeslots){
        # filter presentations on the current timeslot
        $ts = $Presentations | Where-Object {$_.when -eq $timeslot}
        # if there is only one presentation in that timeslot, you're going!
        if ($ts.count -eq 1) {$Agenda += $ts}
        # else pick the presentation with highest interest
        elseif (($ts | Where-Object {$_.interest -gt 0})){
            $agenda += $ts | Sort-Object -Property interest -Descending | Select-Object -first 1
        }
        # if there is no interest, add a filler object
        else {
            $agenda += [pscustomobject]@{
                title       = "***NO MATCH*** pick something random or enjoy a break"
                when        = [datetime]$timeslot
                presenters  = @("N/A")
                keywords    = ""
                interest    = 0

            }
        }
    }
    
    $agenda | Sort-Object -Property when | Format-Table when,title,presenters,keywords,interest -AutoSize
    # presentations is at script level so interest will persist after each execution of New-PSDOAgenda... so we clear it.
    $Presentations | ForEach-Object {$_.interest = 0}
}
