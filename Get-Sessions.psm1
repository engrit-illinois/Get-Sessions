# Documentation home: https://github.com/engrit-illinois/Get-Sessions
# By mseng3

function Get-Sessions {

	param(
		# Name of computer
		# Use "*" as wildcard
		[Parameter(Mandatory=$true,Position=0)]
		[string]$ComputerNameQuery,
		[string]$OUDN,
		[int]$PingCount = 1,
		[switch]$Loud
	)

	function log {
		param (
			[string]$msg,
			[int]$level=0,
			[switch]$nots
		)
		
		if($Loud) {
			for($i = 0; $i -lt $level; $i += 1) {
				$msg = "    $msg"
			}
			
			if(!$nots) {
				$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				$msg = "[$ts] $msg"
			}
			
			Write-Host $msg
			#$msg | Out-File $LOG -Append
		}
	}
	
	function Get-CompNames($nameQuery) {
		$compNames = @()
		log "Getting computer names..."
		if($OUDN) {
		    $compNames = (Get-ADComputer -SearchBase $OUDN -Filter { Name -like $nameQuery }).Name
		}
		else {
		    $compNames = (Get-ADComputer -Filter { Name -like $nameQuery }).Name
		}
		log "Done getting computer names..."
		$compNames
	}

	function Get-Comps($compNames) {
		log "Querying computers..."
		
		$comps = @()
		
		foreach($compName in $compNames) {
			log "Processing `"$compName`"..." -level 1
			$compObject = [PSCustomObject]@{
				Name = $compName
				Sessions = @()
			}
			$compObject.Sessions = Get-Sessions $compName
			$comps += @($compObject)
			log "Done processing `"$compName`"." -level 1
		}
		log "Done querying computers."
		$comps
	}
	
	function Get-SessionObjects($sessionStrings) {
		$sessionObjects = @()
		# Session info is returned as an array of ugly pre-formatted strings
		# So munge the strings into CSV format
		# Note that sessions with unpopulated values will cause problems,
		# This is taken care of while iterating through them
		log "Munging session info into object..." -level 2
		$sessionData = $sessionStrings | foreach { (($_.trim() -replace "\s+",","))}
		# and then turn that into a proper object
		$sessionObjects = $sessionData | ConvertFrom-Csv
		log "Done munging session info into object." -level 2
		$sessionObjects
	}
	
	function Get-Sessions($compName) {
		log "Getting session info..." -level 2
		
		$sessions = @()
		
		log "Testing if computer `"$compName`" responds..." -level 2
		if(Test-Connection -ComputerName $compName -Quiet -Count $PingCount) {
			log "Computer `"$compName`" responded." -level 3
			
			#$sessionError = $true
			try {
				# Have to dump errors to $null, otherwise they are output to the screen, with no way to catch them
				$sessionStrings = query session /server:$compName 2>$null
				#$sessionStrings
				#$sessionError = $false
			}
			catch {
				log "Error getting session info!" -level 2
				#$sessionError = $true
			}
			
			# As it turns out the query executable will not generate a terminating error, meaning catch will not catch anything
			# It will output error strings to the console, but I couldn't find any way to make it throw an error to the pipeline
			#if(!$sessionError) {
			if($sessionStrings) {
			
				$sessions = Get-SessionObjects $sessionStrings
				$sessionsSystem = @()
				$sessionsLocal = @()
				$sessionsRemote = @()
				$sessionsDisc = @()
				$sessionsVS = @()
				$sessionsUnknown = @()
				
				log "Iterating through sessions..." -level 2
				$i = 1
				foreach($session in $sessions) {
					log "Processing session #$i..." -level 3
					
					$session | Add-Member -NotePropertyName COMPUTER -NotePropertyValue $compName
					
					# Depending on the type of session, some fields will be missing.
					# Since the output is a string and not an object, that means our object interpretation of the output will be incorrect, so fix that here
					
					# There are also usually 2 other default sessions, named "services", "rdp-tcp", which have no USERNAME,
					if(
						($session.SESSIONNAME -eq "services") -or
						($session.SESSIONNAME -eq "rdp-tcp") -or
						(($session.SESSIONNAME -eq "console") -and ($session.USERNAME -match "^\d+$"))
					) {
						log "This is the `"$($session.SESSIONNAME)`" session." -level 4
						$session.STATE = $session.ID
						$session.ID = $session.USERNAME
						$session.USERNAME = $null
						$sessionsSystem += @($session)
					}
					# Apparently installations of Visual Studio cause a special session to be present
					# https://www.experts-exchange.com/questions/28484208/Query-session-shows-unknown-session-name-after-Visual-Studio-2013-install.html
					elseif($session.SESSIONNAME -eq "7a78855482a04...") {
						log "This is a weird session apparently caused by installations of Visual Studio." -level 4
						$session.STATE = $session.ID
						$session.ID = $session.USERNAME
						$session.USERNAME = $null
						$sessionsSystem += @($session)
					}
					# What we're left with, SHOULD only be active local login sessions and active or disconnected RDP sessions
					else {
						log "This appears to be either an active local session, or an active or disconnected remote session." -level 4
						
						# Disconnected RDP sessions have no SESSIONNAME
						if($session.ID -eq "Disc") {
							log "This is a disconnected session." -level 4
							$session.STATE = $session.ID
							$session.ID = $session.USERNAME
							$session.USERNAME = $session.SESSIONNAME
							$session.SESSIONNAME = $null
							$sessionsDisc += @($session)
						}
						elseif($session.SESSIONNAME -eq "console") {
							log "This is a local (`"console`") session." -level 4
							$sessionsLocal += @($session)
						}
						elseif($session.SESSIONNAME -match "rdp-tcp#.+") {
							log "This is a remote (`"rdp-tcp#`") session." -level 4
							$sessionsRemote += @($session)
						}
						else {
							log "This session is not recognized!"
							$sessionsUnknown += @($session)
						}
					}
					
					log "Computer: $($session.COMPUTER), SessionName: $($session.SESSIONNAME), User: $($session.USERNAME), ID: $($session.ID), State: $($session.STATE), Type: $($session.TYPE), Device: $($session.DEVICE)" -level 4
						
					log "Done processing session #$i." -level 3
					$i += 1
				}
				log "Done interating through sessions." -level 2
				
				# Output stats for reference
				$sessionsCount = @($sessions).count
				$sessionsSystemCount = @($sessionsSystem).count
				$sessionsLocalCount = @($sessionsLocal).count
				$sessionsRemoteCount = @($sessionsRemote).count
				$sessionsDiscCount = @($sessionsDisc).count
				$sessionsUnknownCount = @($sessionsUnknown).count
				$sessionsNonSystemCount = $sessionsLocalCount + $sessionsRemoteCount + $sessionsDiscCount + $sessionsUnknownCount
				
				# Check that our logic was correct
				if(($sessionsNonSystemCount + $sessionsSystemCount) -ne $sessionsCount) {
					throw "Error with session type logic detected."
				}
				
				log "Found $sessionsCount sessions." -level 2
				log "Found $sessionsSystemCount system sessions." -level 2
				log "Found $sessionsNonSystemCount non-system sessions." -level 2
				log "Found $sessionsLocalCount local sessions." -level 3
				log "Found $sessionsRemoteCount remote sessions." -level 3
				log "Found $sessionsDiscCount disconnected remote sessions." -level 3
				log "Found $sessionsUnknownCount unknown sessions." -level 3
			}
			else {
				log "No sessions found (not even system sessions), or failed to get session info. Skipping `"$compName`"." -level 2
			}
		}
		else {
			log "Computer `"$compName`" did not respond!" -level 3
		}

		$sessions
	}
	
	function Order-Sessions($sessions) {
		$sessions | Sort COMPUTER | Select -Property COMPUTER,USERNAME,SESSIONNAME,STATE,ID,TYPE,DEVICE
	}
	
	function Combine-Sessions($comps) {
		$sessions = @()
		foreach($comp in $comps) {
			foreach($session in $comp.Sessions) {
				$sessions += @($session)
			}
		}
		$sessions
	}
	
	function Print-Sessions($sessions) {
		log " " -nots
		# The TYPE and DEVICE fields always seem to be empty, so omit them here, and reorgnaize the columns
		$sessions = $sessions | Format-Table -Property COMPUTER,USERNAME,SESSIONNAME,STATE,ID
		$sessions = ($sessions | Out-String).Trim()
		log $sessions -nots
		log " " -nots
	}
	
	function Process {
		$compNames = Get-CompNames $ComputerNameQuery
		$comps = Get-Comps $compNames
		$sessions = Combine-Sessions $comps
		$sessions = Order-Sessions $sessions
		Print-Sessions $sessions
		$sessions
	}
	
	Process
	
	log "EOF"
	log " " -nots
}