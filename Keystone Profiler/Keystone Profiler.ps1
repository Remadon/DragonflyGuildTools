# Keystone Profiler
# v1.1  4/23/2023 | Hathlo of Wyrmrest Accord | Dragonfly
# Designed for World of Warcraft / Raider.IO API 8/9/2022
# A set of functions and code to pull keystone data from the raider.io API and format accordingly.

# v0.1 8/8/2022 - Initial Design
# v0.2 2/1/2023 - De-culred the script to use native powershell cmdlets.
# v1.0 4/9/2023 - Implemented the use of the ImportExcel module, including code to install the module if not found.
#			      	We now create a single report with a summary page, and all relevant data, in a single excel file. 
#				  Implemented a supporting function, ConvertTo-DataTable. This takes a PSObject and converts it to a data table.
#				  Dungeon List is now stored in configuration section rather than Get-KeystoneProfile function.
#				  Implemented some configuration variables. 
#					$ DungeonList - For configuring what dungeons to look for data on
#					$ SaveLocation - For configuring a path to save the report
#					$ SaveName - For configuring the name of the report file
#				  Implemented a Summary page in the report data. This will show relevant data about the character we got data for, such as name, realm, and other things.
# v1.1 4/23/2023 - Implemented some formatting and underlining to make links on the summary page more apperant, and added a link back to the summary page in each character page.

# Todo:
# - Split the creation of $Array objects into a sperate function and get rid of some of this spheggeti code.
# - Add sanity checking and try catch blocks. Catch and respond to common errors.
#	  Add some logic to check for 400 errors from the raiderIO API. A common resolution to this may be capitolization or lower case beginning characters in a player name. 
#	  For example, Âurore (suceeding) vs. âurore (failing) queries. We can add some hail mary logic to capitolize or lower the font of the first character, and try again.
#		  In the event of hail mary failure, we should end with an error log (player name, realm, and failure E.G. 404, 400, 501, etc.) in a sheet in the report, and we should not create a row or formulas pointing to places that do not exist.
# - Future Plans: Add support to Get-KeystoneProfile to handle each of the exteraparams.
# - Future Plans: Add support to Get-KeystoneProfile to validate data against the WoW API. Hopefully the WoW API will return better data someday.
# - Apperantly the act of hiding a column sets the data type to be correct (Looking at you, character Dungeon Score column!). Look into using this behavior to fix the issue and remove all the logic we do with Column J.
# - Set links to each sheet in the summary page. For ease of use, set a cell in each player sheet to go back to the summary page.
# - Progress Bar, and see if we can supress the text spam this script currently creates.

######################################
# CHECK, IMPORT, OR INSTALL, MODULES #
######################################
# Check if ImportExcel module is installed
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    # ImportExcel module not found, install it.
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -Verbose
    }
    catch {
        Write-Warning "Failed to install ImportExcel module. Please make sure you have an internet connection and try again."
        Exit 1
    }
}

Import-Module -Name ImportExcel

########################
# SCRIPT CONFIGURATION #
########################
# Dungeon List:
#   This needs to be set to whatever dungeons are the current season, else the script will give you old data from old seasons and timewalking keystones.
$Global:DungeonsList = @("Algeth'ar Academy", "Court of Stars", "Halls of Valor", "Ruby Life Pools", "Shadowmoon Burial Grounds", "Temple of the Jade Serpent", "The Azure Vault", "The Nokhud Offensive") 

# Save Location:
#   This is the location (path) and filename that the script will save the output file.
#   By default this is $env:USERPROFILE\Desktop\KeystoneReport.xlxs. $env:USERPROFILE is a windows defualt variable that always referrs to your user profile (C:\Users\YOUR_USERNAME_HERE).
# Save Name:
#   This is the name of the file.
$SaveLocation = Join-Path $env:USERPROFILE "Desktop"
$SaveName = Join-Path $SaveLocation "KeystoneReport.xlsx"

# Raider IO API Extra parameters:
#   When we talk to the RaiderIO API we need to tell it to get us extra data not given by default. This string controlls this, and generally should not be changed.
#   If the day comes that this needs to change, logic will need to be adjusted in Get-KeystoneProfile to account for it.
#	See RaiderIO's documentation on the subject for more details: https://raider.io/api#/character/getApiV1CharactersProfile
$RaiderIOExtraParams = "mythic_plus_scores_by_season:current,mythic_plus_best_runs,mythic_plus_alternate_runs"

#########################
# FUNCTIONS & WORKFLOWS #
#########################

# If the script is not running in the ISE, set the text encoding appropriately. Pour one out for Hathlo's sanity, this caused all sorts of problems before being found.
if ($psise -eq $null) {[Console]::OutputEncoding = [System.Text.Encoding]::UTF8}

# Get-KeystoneProfile 
# Obtain the raiderIO profile for a specified character. ExtraParams may contain any number (in comma seperated values) of extra parameters you want to get from RaiderIO.
#	Currently, this script expects the use of the following exteraparams: "mythic_plus_scores_by_season:current,mythic_plus_best_runs,mythic_plus_alternate_runs"
Function Get-KeystoneProfile {
Param(
    [parameter(Mandatory=$true)]$Region,
    [parameter(Mandatory=$true)]$RealmSlug,
    [parameter(Mandatory=$true)]$CharacterName,
    [parameter(Mandatory=$false)]$ExtraParams
)

    # Get Character General Profile. 
    # Depending on ExtraParams being specified, change the url query we use.
    if ($null -eq $ExtraParams) {$URL = "https://raider.io/api/v1/characters/profile?region=$($Region)&realm=$($RealmSlug)&name=$($CharacterName)"}
    if ($null -ne $ExtraParams) {$URL = "https://raider.io/api/v1/characters/profile?region=$($Region)&realm=$($RealmSlug)&name=$($CharacterName)&fields=$($ExtraParams)"}
    $GeneralProfile = Invoke-RestMethod -Uri $URL -Headers @{"accept" = "application/json"} -Method GET

    $Array = @()
    # For each run in best runs, get the dungeon, score, main affix, keystone level, completed time, par time. Store this in $Array.
    ForEach ($Run in $GeneralProfile.mythic_plus_best_runs) {
        $Dungeon = New-Object -TypeName PSObject
        
        # Get date, run time and par time in a good readable format.
        $ClearDate = ((Get-Date $Run.completed_at).ToShortDateString() + " " + (Get-Date $Run.completed_at).ToShortTimeString())
        $ClearTime = ([timespan]::FromMilliseconds($Run.clear_time_ms)).ToString("hh\:mm\:ss")
        $ParTime = ([timespan]::FromMilliseconds($Run.par_time_ms)).ToString("hh\:mm\:ss")
        
        # Calculate the run remaining time, showing how far behind or ahead of par the run was.
        # To do this, we need to check if the par time minus the clear time is a negative integer, as this effects how we display the result.
        if (($Run.par_time_ms - $Run.clear_time_ms) -le 0) {$SplitTime = ("-" + [timespan]::FromMilliseconds(($Run.par_time_ms - $Run.clear_time_ms)).ToString("hh\:mm\:ss"))}
        else {$SplitTime = [timespan]::FromMilliseconds(($Run.par_time_ms - $Run.clear_time_ms)).ToString("hh\:mm\:ss")}
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Dungeon' -Value $Run.dungeon
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Affix' -Value $Run.affixes[0].name
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Key Level' -Value $Run.mythic_level
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Completed' -Value $ClearDate
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Clear Time' -Value $ClearTime
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Par Time' -Value $ParTime
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Remaining Time' -Value $SplitTime
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Run Score' -Value $Run.score
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Best/Alternate' -Value "Best"
        $Array += $Dungeon
    }
    # For each run in alternate runs, get the dungeon, score, main affix, keystone level, completed time, par time. Store this in $Array.
    ForEach ($Run in $GeneralProfile.mythic_plus_alternate_runs) {
        $Dungeon = New-Object -TypeName PSObject
        
        # Get run time and par time in a good readable format.
        $ClearDate = ((Get-Date $Run.completed_at).ToShortDateString() + " " + (Get-Date $Run.completed_at).ToShortTimeString())
        $ClearTime = ([timespan]::FromMilliseconds($Run.clear_time_ms)).ToString("hh\:mm\:ss")
        $ParTime = ([timespan]::FromMilliseconds($Run.par_time_ms)).ToString("hh\:mm\:ss")
        
        # Calculate the run split, showing how far behind or ahead of par the run was.
        # To do this, we need to check if the par time minus the clear time is a negative integer, as this effects how we display the result.
        if (($Run.par_time_ms - $Run.clear_time_ms) -le 0) {$SplitTime = ("-" + [timespan]::FromMilliseconds(($Run.par_time_ms - $Run.clear_time_ms)).ToString("hh\:mm\:ss"))}
        else {$SplitTime = [timespan]::FromMilliseconds(($Run.par_time_ms - $Run.clear_time_ms)).ToString("hh\:mm\:ss")}
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Dungeon' -Value $Run.dungeon
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Affix' -Value $Run.affixes[0].name
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Key Level' -Value $Run.mythic_level
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Completed' -Value $ClearDate
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Clear Time' -Value $ClearTime
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Par Time' -Value $ParTime
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Remaining Time' -Value $SplitTime
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Run Score' -Value $Run.score
        Add-Member -InputObject $Dungeon -MemberType NoteProperty -Name 'Best/Alternate' -Value "Alternate"
        $Array += $Dungeon
    }
    # Check for missing runs. Add dummy data into the array as needed.
    # Determine if we have a run for each dungeon on Tyrannical. If not, create a record for the dungeon on Tyrannical with a score of zero.
    ForEach ($Dungeon in $Global:DungeonsList) {
        # Create a temporary list of dungeons that are only Tyrannical from $Array. We check to see if a dungeon is missing, then add a run to $Array if so.
        $Working = $Array | Where Affix -eq "Tyrannical"
        if ($Working.dungeon -contains "$($Dungeon)") {continue}
        else {
            $Object = New-Object -TypeName PSObject
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Dungeon' -Value $Dungeon
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Affix' -Value "Tyrannical"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Key Level' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Completed' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Clear Time' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Par Time' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Remaining Time' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Run Score' -Value 0
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Best/Alternate' -Value "Not Done!"
            $Array += $Object
        }
    }
    # Determine if we have a run for each dungeon on Fortified. If not, create a record for the dungeon on Fortified with a score of zero.
    ForEach ($Dungeon in $Global:DungeonsList) {
        # Create a temporary list of dungeons that are only Fortified from $Array. We check to see if a dungeon is missing, then add a run to $Array if so.
        $Working = $Array | Where Affix -eq "Fortified"
        if ($Working.dungeon -contains "$($Dungeon)") {continue}
        else {
            $Object = New-Object -TypeName PSObject
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Dungeon' -Value $Dungeon
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Affix' -Value "Fortified"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Key Level' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Completed' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Clear Time' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Par Time' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Remaining Time' -Value "N/A"
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Run Score' -Value 0
            Add-Member -InputObject $Object -MemberType NoteProperty -Name 'Best/Alternate' -Value "Not Done!"
            $Array += $Object
        }
    }
    return @{GeneralProfile = $GeneralProfile; Array = $Array}
}

Function ConvertTo-DataTable {
[CmdLetBinding(DefaultParameterSetName="None")]
	Param(
	 [Parameter(Position=0,Mandatory=$true)][System.Array]$Source,
	 [Parameter(Position=1,ParameterSetName='Like')][String]$Match=".+",
	 [Parameter(Position=2,ParameterSetName='NotLike')][String]$NotMatch=".+"
	)
	if ($NotMatch -eq ".+"){$Columns = $Source[0] | Select * | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -match "($Match)"}}
	else {$Columns = $Source[0] | Select * | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -notmatch "($NotMatch)"}}
	
	$DataTable = New-Object System.Data.DataTable
	Foreach ($Column in $Columns.Name){
		$DataTable.Columns.Add("$($Column)") | Out-Null
	}
	
	#For each row (entry) in source, build row and add to DataTable.
	Foreach ($Entry in $Source) {
		$Row = $DataTable.NewRow()
			Foreach ($Column in $Columns.Name) {
				$Row["$($Column)"] = if($Entry.$Column -ne $null){($Entry | Select-Object -ExpandProperty $Column) -join ', '}else{$null}
			}
		$DataTable.Rows.Add($Row)
	}
	
	#Validate source column and row count to DataTable
	if ($Columns.Count -ne $DataTable.Columns.Count){
		Throw "Conversion failed: Number of columns in source does not match data table number of columns"
	}
	else { 
		if($Source.Count -ne $DataTable.Rows.Count){throw "Conversion failed: Source row count not equal to data table row count"}
		#The use of "Return ," ensures the output from function is of the same data type; otherwise it's returned as an array.
		else {Return ,$DataTable}
	 }
}

######################
# BEGIN SCRIPT       #
######################

<#------ Begin test code
$Data = Get-KeystoneProfile -Region "us" -RealmSlug "wyrmrest-accord" -CharacterName "âurore" -ExtraParams "mythic_plus_scores_by_season:current,mythic_plus_best_runs,mythic_plus_alternate_runs"
------#>

# Create the report excel workbook.
$excelWorkbook = New-Object -TypeName OfficeOpenXml.ExcelPackage
$summaryWorksheet = $excelWorkbook.Workbook.Worksheets.Add("Summary")

# Hard code some cells to contain strings. In our foreach loop, we will be using some formulas in some cells to reference data in sheets and cells this script generates.
$summaryWorksheet.Cells["A1"].Value = "Player"
$summaryWorksheet.Cells["B1"].Value = "Race"
$summaryWorksheet.Cells["C1"].Value = "Class"
$summaryWorksheet.Cells["D1"].Value = "Spec"
$summaryWorksheet.Cells["E1"].Value = "Faction"
$summaryWorksheet.Cells["F1"].Value = "Realm"
$summaryWorksheet.Cells["G1"].Value = "Total Score"
$summaryWorksheet.Cells["H1"].Value = "Worst Dungeon"
$summaryWorksheet.Cells["I1"].Value = "Worst Dungeon Score"

# Populate our player data from CSV.
$Players = Import-Csv $PSScriptRoot\players.csv

# Get Data, Format Data, Output Data. 
# We need to count how many times we've iterated here in order to keep track of where we add cells in the summary page.
$Count = 1	# Always start at 1, becuase we've populated row 1 hardcoded above.
$PlayerCount = $Players.count

ForEach ($Player in $Players) {
    # Get the character's general profile and keystone data.
    $Data = Get-KeystoneProfile -Region us -RealmSlug $Player.RealmSlug -CharacterName $Player.Name -ExtraParams $RaiderIOExtraParams
	
	# Format the data, Then process data into our Excel Workbook, and create cells and formulas using ImportExcel.
	$DataTable = ConvertTo-DataTable -Source ($Data.Array | Sort-Object -Property Dungeon) -Match "Dungeon|Affix|Key Level|Completed|Clear Time|Par Time|Remaining Time|Run Score|Best/Alternate"
	
	# Create the workbook objects and import data.
	$worksheetName = $Data.GeneralProfile.name
    $worksheet = $excelWorkbook.Workbook.Worksheets.Add($worksheetName)
    $worksheet.Cells.LoadFromDataTable($DataTable,$true)
		
	# Increment the counter so that we add cell text and formulas in the correct possition.
	$Count++	
	
	# Insert Cells and Formulas in the player's sheet to hide some sins of not correctly specifying the data type as integer for run score and key level when we export to file.
	# Create another row for run score, using a formula to get a numerical value we can then use to sum dungeon scores for each dungeon.
	$worksheet.Cells["J1"].Value = "Run Score"
	$worksheet.Cells["J2"].Formula = "=VALUE(I2)"
	$worksheet.Cells["J3"].Formula = "=VALUE(I3)"
	$worksheet.Cells["J4"].Formula = "=VALUE(I4)"
	$worksheet.Cells["J5"].Formula = "=VALUE(I5)"
	$worksheet.Cells["J6"].Formula = "=VALUE(I6)"
	$worksheet.Cells["J7"].Formula = "=VALUE(I7)"
	$worksheet.Cells["J8"].Formula = "=VALUE(I8)"
	$worksheet.Cells["J9"].Formula = "=VALUE(I9)"
	$worksheet.Cells["J10"].Formula = "=VALUE(I10)"
	$worksheet.Cells["J11"].Formula = "=VALUE(I11)"
	$worksheet.Cells["J12"].Formula = "=VALUE(I12)"
	$worksheet.Cells["J13"].Formula = "=VALUE(I13)"
	$worksheet.Cells["J14"].Formula = "=VALUE(I14)"
	$worksheet.Cells["J15"].Formula = "=VALUE(I15)"
	$worksheet.Cells["J16"].Formula = "=VALUE(I16)"
	$worksheet.Cells["J17"].Formula = "=VALUE(I17)"
	# Literally hide our sins.
	Set-Column -Worksheet $worksheet -Column 9 -Hide
	# Create cells and formulas to show each dungeon score total:
	$worksheet.Cells["A19"].Value = "Totals:"
	$worksheet.Cells["A20"].Value = "Dungeon Name"
	$worksheet.Cells["B20"].Value = "Total Score"
	$worksheet.Cells["A21"].Formula = "=E2"
	$worksheet.Cells["B21"].Formula = "=SUM(J2:J3)"
	$worksheet.Cells["A22"].Formula = "=E4"
	$worksheet.Cells["B22"].Formula = "=SUM(J4:J5)"
	$worksheet.Cells["A23"].Formula = "=E6"
	$worksheet.Cells["B23"].Formula = "=SUM(J6:J7)"
	$worksheet.Cells["A24"].Formula = "=E8"
	$worksheet.Cells["B24"].Formula = "=SUM(J8:J9)"
	$worksheet.Cells["A25"].Formula = "=E10"
	$worksheet.Cells["B25"].Formula = "=SUM(J10:J11)"
	$worksheet.Cells["A26"].Formula = "=E12"
	$worksheet.Cells["B26"].Formula = "=SUM(J12:J13)"
	$worksheet.Cells["A27"].Formula = "=E14"
	$worksheet.Cells["B27"].Formula = "=SUM(J14:J15)"
	$worksheet.Cells["A28"].Formula = "=E16"
	$worksheet.Cells["B28"].Formula = "=SUM(J16:J17)"
	# Create summary page cells and formulas:
	$summaryWorksheet.Cells["A$($Count)"].Value = $Data.GeneralProfile.name
	$summaryWorksheet.Cells["A$($Count)"].HyperLink = "#$($worksheetName)!A1"
	$summaryWorksheet.Cells["A$($Count)"].Style.Font.UnderLine = "Single"
	$summaryWorksheet.Cells["A$($Count)"].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
	$summaryWorksheet.Cells["B$($Count)"].Value = $Data.GeneralProfile.race
	$summaryWorksheet.Cells["C$($Count)"].Value = $Data.GeneralProfile.'class'
	$summaryWorksheet.Cells["D$($Count)"].Value = $Data.GeneralProfile.active_spec_name
	$summaryWorksheet.Cells["E$($Count)"].Value = $Data.GeneralProfile.faction
	$summaryWorksheet.Cells["F$($Count)"].Value = $Data.GeneralProfile.realm
	# Add formulas to cells in the summary tab to display certain data at a glance on the summary tab.
	$summaryWorksheet.Cells["G$($Count)"].Formula = "=SUM($($worksheetName)!B21:B28)"	
	# This will be the worst dungeon name
	$summaryWorksheet.Cells["H$($Count)"].Formula = "=INDEX($($worksheetName)!A21:A28,MATCH(MIN($($worksheetName)!B21:B28),$($worksheetName)!B21:B28,0))"
	# This will be the worst dungeon score
	$summaryWorksheet.Cells["I$($Count)"].Formula = "=MIN($($worksheetName)!B21:B28)"
	# Create a link on the player's sheet that links back to the summary page.
	$worksheet.Cells["A30"].Value = "Back to Summary"
	$worksheet.Cells["A30"].Hyperlink = "#$($summaryWorksheet)!A$($Count)"
	$worksheet.Cells["A30"].Style.Font.UnderLine = "Single"
	$worksheet.Cells["A30"].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
}

# Save the Excel workbook to the specified file path
$excelWorkbook.SaveAs($SaveName)
