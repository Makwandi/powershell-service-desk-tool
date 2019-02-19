

#Change view-name for attributes, later used in Format-Table
$myFormat = @{Expression={$_.givenName};Label="FirstName"},
            @{Expression={$_.surName};Label="LastName"},
			@{Expression={$_.samAccountName};Label="Username"}, 
			@{Expression={$_.userPrincipalName};Label="Mail"}, 
			@{Expression={$_.employeeID};Label="ID"},
            @{Expression={$_.pager};Label="Cost Centre"}, 
			@{Expression={$_.employeenumber};Label="Job Grade"}
    
#Setting standard properties, later used for displaying values in console      
$Properties = 
@(
        'givenName',
        'surName', 
        'samAccountName', 
        'userPrincipalName',
        'employeeID',
        'pager',  
        'employeenumber')

######################################################### Functions

#Start-txt-CFG Function
function Start-txt-CFG {

Import-Module -Name Fazer-ServiceDesk
#Connect-ExchangeServer -auto -ClientApplication:ManagementShell

Write-Host "`n##################################################"
Write-Host "`nAll Functions version 2.3 loaded" -ForegroundColor  Cyan
Write-Host "`nInstructions:" -ForegroundColor  Yellow
Write-Host "`n- Write go to list all functions" 
Write-Host "- Write go and a number to jump directly to a function, exampel: go 4"
Write-Host "- Write un to search for a username that's needed for some functions, you don't need to open the AD"
Write-Host "- Use Ctrl + C to cancel an ongoing function"
Write-Host "`nUpdate Info" -ForegroundColor Cyan
Write-Host "`nNew function available, nr 20. You now can move users accounts (NOT rest. accounts) from Win7x64 OU to Win10 OU."

Write-Host "`n##################################################"
Write-Host "`n"
}#Start-txt-CFG END

#Connect-OnPremise Function 
Function Connect-OnPremise {

    . 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell

}#Connect-OnPremise END

#un (username) Function
function un {
    
    $loop = $true
    do{
        $firstName = Read-Host "First name"
        
        $lastName = Read-Host "Last name"
   

        
        if(!$firstName){
            $lastName = "*" + $lastName + "*"
            $result = Get-ADuser -Properties $Properties  -Filter {surName -like $lastName} 
        }
        elseif(!$lastName){
            $firstName = "*" + $firstName + "*"
            $result = Get-ADuser -Properties $Properties  -Filter {givenName -like $firstName} 
        }
        else{
            $firstName = "*" + $firstName + "*"
            $lastName = "*" + $lastName + "*"
            $result = Get-ADuser -Properties $Properties  -Filter {givenName -like $firstName -and surName -like $lastName} 
        }
        
        #$result = Get-ADuser -Properties $Properties  -Filter {givenName -like $firstName -and surName -like $lastName}
        if(!$result){
            write-host "User does not exist, try again" -Foreground red
            $loop = $true
        }#if
        else{
            $loop = $false
        }#else
    }
    while($loop)
    if ($result.count -ge 2)
        {

            foreach ($user in $result){
                Write-Host $result.IndexOf($user)"."$user.givenName " " $user.surname " | " $user.samAccountName
            }#foreach
    
    $loop = $true
    do {
        $chosenUser = Read-Host "Select user from list"
        if ([int]$chosenUser -lt 0 -or [int]$chosenUser-gt $result.Length - 1) {
            Write-host "wrong Input, try again" -ForegroundColor Red
            $loop = $true
        }#if
            else {
            $loop = $false
        }#else

    }#do
    while ($loop)

        $result = $result[$chosenUser] 
}#if

    Write-Host "Is this the user:" -ForegroundColor Yellow
    $result | select $Properties  | Format-Table $myFormat -AutoSize

    $continueOrRetry = Read-Host "Is this correct? Type |1| to continue or |0| to retry"
    switch ($continueOrRetry){

        1 {
        Set-Clipboard -Value $result.samAccountName
        Write-Host "The username has been inserted to your clipboard" -ForegroundColor Green
        }
        0{
        un
        }

    }

}#Function un END

#Get-RezPut Function
function Get-RezPut{
    Param(
       
        [Parameter(Mandatory=$true)]  
        $whatToOutput
    )
    do{
        Write-Host "   `nMake a choice:   " -ForegroundColor  white -BackgroundColor Black
	    Write-Host "`n##################################################"
	    Write-Host "`n1.Export to CSV?"-ForegroundColor  DarkYellow
	    Write-Host "2.Write in Console?" -ForegroundColor  Green
	    Write-Host "`n##################################################"
	    [int]$choiceOutput = Read-Host "`nPick a number"
    }while (!$choiceOutput)
                                
    switch ($choiceOutput) { 


			                    #Export to CSV
                                #----------------------------------------------------------------------
                                1{  
                                    Write-Host "`nYou picked 1.Export to CSV"-ForegroundColor  Yellow
                                    $DriveletterExists = Test-Path -Path "h:" 
                                    If (-not ($DriveletterExists)) {
                                        Write-Host "`nH: drive not found, configure it before you continue" -ForegroundColor  Cyan
                                        Break
                                    }
                                    do{
                                        [String]$outputName = Read-Host "`nChoose a name for the CSV"
                                    }while (!$outputName)
                                    $whatToOutput |
                                    #Change view-name for attributes
					                foreach {
						         	        new-object psobject -Property @{
							                        Firstname = $_.givenName
									                Lastname = $_.surName
									                Username = $_.samAccountName
									                Mail = $_.userPrincipalName
									                EmployeeID = $_.employeeID
                                                    CostCenter = $_.pager
									                JobGrade = $_.employeenumber

						                }}|
					                Select Firstname,Lastname,Username,Mail,EmployeeID, CostCenter, JobGrade |
					                Export-Csv -Path H:\$outputName.csv -Encoding UTF8 -NoTypeInformation 
                                    Write-Host "The file is saved @ H:\$outputName.csv" -ForegroundColor  Cyan
                                    Invoke-Item H:\$outputName.csv
				                }
                                #Write in console 
                                #----------------------------------------------------------------------
				                2{ 
                                    Write-Host "`nYou picked 2.Write in console"-ForegroundColor  Yellow
			                        $whatToOutput |Format-Table $myFormat -AutoSize            
				                }
                            

                                default{
		                            Write-Host "`nSomething went wrong, try again"-ForegroundColor  Red
                                    

                                }
        }#switch

}#Get-RezPut END

######################################################### Functions End

Start-txt-CFG
#Main Go function
function Go {
    #Giving the function a int parameter so i can start the function at a number 
    #wrongInput is for redirecting back to a number when input is incorrect 
    Param([Parameter(Mandatory=$false)]
    [int]$choiceMain, [int]$wrongInput)
    
        #Give the user choices, later connected to a switch
	    try{
            
		    do{
                if(!$choiceMain){
			    $boolean = $TRUE
					    Write-Host "`nWelcome to the /(? × ?)\ hole, what do you want to do?" -ForegroundColor  Yellow
					   
                        Write-Host "`n`nMost Used" -ForegroundColor  Yellow
                        Write-Host "##################################################"
                        Write-Host "`n|Nr: 9|" -NoNewline
                        Write-Host "`t`t? Unlock account? ?"-ForegroundColor  Yellow
                        Write-Host "`n|Nr: 16|" -NoNewline
                        Write-Host "`t? Add an application to a computer? ?" -ForegroundColor  Yellow
                        Write-Host "`n|Nr: 7|" -NoNewline
                        Write-Host "`t`t? Disable an account? ?" -ForegroundColor  Yellow
                        Write-Host "`n|Nr: 8|" -NoNewline
                        Write-Host "`t`t? Generate a random password? ?" -ForegroundColor  Yellow
                        Write-Host "`n|Nr: 20|" -NoNewline
                        Write-Host "`t? Move user account (NOT rest. account) to Win10 OU in AD? ?" -ForegroundColor  Yellow
                         
                        Write-Host "`n`nUser Information" -ForegroundColor  Yellow
                        Write-Host "##################################################"
                        Write-Host "`n|Nr: 1|" -NoNewline
					    Write-Host "`t`tDoes this user already exist in the AD?" -ForegroundColor  Green
                        Write-Host "`n|Nr: 2|" -NoNewline
                        Write-Host "`t`tWhat groups are assigned to this user?"-ForegroundColor  Green
                        Write-Host "`n|Nr: 3|" -NoNewline       
					    Write-Host "`t`tWho is this user the boss over?"-ForegroundColor  Green
                        Write-Host "`n|Nr: 4|" -NoNewline
					    Write-Host "`t`tWhat job grade does this user have?"-ForegroundColor  Green
                        Write-Host "`n|Nr: 5|" -NoNewline       
					    Write-Host "`t`tFind a user by first name and cost centre?"-ForegroundColor  Green
                        Write-Host "`n|Nr: 6|" -NoNewline
                        Write-Host "`t`tWhat license does this user have?"-ForegroundColor  Green
                        Write-Host "`n|Nr: 10|" -NoNewline
                        Write-Host "`tChange mail address?"-ForegroundColor  Green
                        Write-Host "`n|Nr: 11|" -NoNewline
                        Write-Host "`tEnable mail address?" -ForegroundColor  Green 
                        Write-Host "`n|Nr: 19|" -NoNewline
                        Write-Host "`tAre all the user attributes OK?" -ForegroundColor  Green 
                        Write-Host "`n|Nr: 21|" -NoNewline
                        Write-Host "`tAdd attributes to new account?" -ForegroundColor  Green 
                        
                        Write-Host "`n`nGroup Infromation" -ForegroundColor  Yellow      
					    Write-Host "##################################################"
                        Write-Host "`n|Nr: 14|" -NoNewline
                        Write-Host "`tWhich users are in a group?" -ForegroundColor  Green
                        Write-Host "`n|Nr: 15|" -NoNewline
                        Write-Host "`tList accounts by country?" -ForegroundColor  Green                    
                        Write-Host "`n|Nr: 17|" -NoNewline
                        Write-Host "`tWhich groups is the user the owner of?" -ForegroundColor  Green
                        Write-Host "`n|Nr: 18|" -NoNewline
                        Write-Host "`tIs the user in the given group? <= (nested groups included)" -ForegroundColor  Green
                        
                        Write-Host "`n`nCreate functions" -ForegroundColor  Yellow      
                        Write-Host "##################################################"
                        Write-Host "`n|Nr: 12|" -NoNewline
                        Write-Host "`tCreate a new shared mailbox?" -ForegroundColor  Green
                        Write-Host "`n|Nr: 13|" -NoNewline
                        Write-Host "`tCreate a new DL-group?" -ForegroundColor  Green
                        
                        Write-Host "`nHow to cancel and abort" -ForegroundColor  Yellow       
                        Write-Host "##################################################"
                        Write-Host "`n|Nr: 0|" -NoNewline
                        Write-Host "`t`tTerminate the mission" -ForegroundColor Green
                        Write-Host "`nYou can also press " -NoNewline
                        Write-Host "|Ctrl + C|" -NoNewline -ForegroundColor Red -BackgroundColor Black 
                        Write-Host " anywhere in the execution to abort the mission."
                        Write-Host "In case of emergency, use " -NoNewline
                        Write-Host "|Alt + F4|" -NoNewline -ForegroundColor Red -BackgroundColor Black
                        Write-Host " ===> WHAT ?(????) " 

                        Write-Host "`n##################################################"

                        #Check if another script has sent the user back here
                        Write-Host "`n=====>" -ForegroundColor Magenta -NoNewline
                        $choiceMain = Read-Host " Pick a number"
                        }#If ChoiceMain
                if($wrongInput -eq 1){
                             Write-Host "`nComputer says nooouuh" -ForegroundColor Red
                             Write-Host "(?°?°)?? ???"
                             Write-Host "Try again"-ForegroundColor Red
                             Clear-variable wrongInput 
                        }#If Wrong Input
                #Check input-value
				if((([int]$choiceMain -ge  22 ) -or ([int]$choiceMain -lt 0))){
						    $boolean = $FALSE
                            Write-Host "`nWrong input, try again" -ForegroundColor  Red
                            Clear-Variable choiceMain
                            
			     }#If ChoiceMain count
		    #Retry until input-value is correct		
            }while ($boolean -eq $FALSE)
	    }catch{
		    Write-Host "`nWrong input (?°?°)?? ???, try again " -ForegroundColor  Red
		    
	    }#Catch
        #Switch connected to choices 
	    switch ($choiceMain) { 
        #1.Does this user already exist in the AD?
        #//////////////////////////////////////////////////////////////////////////////////////////
	        1 {
		                Write-Host "`nYou picked: 1.Does this user already exist in the AD" -ForegroundColor Yellow
                        $mail = Read-Host "`nMail"
				        $mail = "*" +  $mail + "*"
				        $user = Read-Host "Username"
				        $user = "*" +  $user + "*" 

				        $mailStatus = Get-ADUser -Filter {proxyAddresses -like $mail} -Properties $Properties
				        if(!$mailStatus){
					        Write-Host "`nMail address does not exist" -ForegroundColor  Green                    
				        }else{
					        Write-Host 'Mail address does exist'-ForegroundColor  Red
                            $mailStatus | Select $Properties |Format-Table $myFormat -AutoSize   
				        } 

				        $userStatus = Get-ADUser -Filter {samAccountName -like $user} -Properties $Properties
					    If(!$userStatus){
					        Write-Host "Username does not exist" -ForegroundColor  Green                              
				        }else{
					        Write-Host 'Username does exist'-ForegroundColor  Red
                            $userStatus | Select $Properties |Format-Table $myFormat -AutoSize    
				        } 
                         Clear-variable user, mail
	        }#1 
            #2.What groups are assigned to this user?
            #//////////////////////////////////////////////////////////////////////////////////////////
	        2 {
		                try{
                            Write-Host "`nYou picked: 2.What groups are assigned to this user" -ForegroundColor Yellow
				            $user = Read-Host "`nUsername"
					        Get-ADPrincipalGroupMembership  $user  | select -ExpandProperty distinguishedName |
					        Get-ADGroup | select -ExpandProperty Name
                            Clear-variable user
		                }catch{
                              Go 2 1
		                }
	        }#2 
            #3.Who is this user the boss over?
            #//////////////////////////////////////////////////////////////////////////////////////////
	        3 {
		                Write-Host "`nYou picked: 3.Who is this user the boss over" -ForegroundColor Yellow    
                        $user = Read-Host "`nUsername"
                        try{
                            #Gets only the OU, without the header-info
				            $OU = Get-ADUser $user | Select -ExpandProperty DistinguishedName
                            Clear-variable user

                            $MInfo = Get-ADUser -Filter {manager -eq $OU} -Properties $Properties
                        }Catch{
                            Go 3 1
                        }
				        Get-RezPut $MInfo
            
	        }#3 
            #4.What job grade does this user have?
            #//////////////////////////////////////////////////////////////////////////////////////////
	        4 {
                        Write-Host "`nYou picked: 4.What job grade does this user have" -ForegroundColor Yellow    
	                    $user = Read-Host "`nUsername"
                        $JGInfo = Get-ADUser $user -Properties $Properties | 
                        Select $Properties  
                        return $JGInfo| Format-Table $myFormat -AutoSize
                        Clear-variable user
	        }#4
            #5.Find a user by cost centre and first name
            #////////////////////////////////////////////////////////////////////////////////////////// 
	        5 {
		                Write-Host "`nYou picked: 5.Find a user by cost centre and first name" -ForegroundColor Yellow    
                        do{
                            $name =  Read-host "`nName"
                            $pager = Read-host "Cost centre"
                            $pager = "*" +  $pager + "*"
                            $user = Get-ADUser -Properties $Properties -Filter {(Pager -like $pager ) -and (GivenName -like $name)}
                            if(!$user){
                                Write-Host "`nUsername does not exist ¯\_(?)_/¯, try again"-ForegroundColor Red
                                }
                        }while(!$user)

                        $user  | select $Properties  | Format-Table $myFormat -AutoSize
                        Clear-variable user  

	        }#5
            #6.What license does this user have?
            #//////////////////////////////////////////////////////////////////////////////////////////
	        6 {
		            
                            Write-Host "`nYou picked: 6.What license does this user have" -ForegroundColor Yellow    
                            Write-Host "`nLogon to MsolService"-ForegroundColor Green
                            Write-Host "Example: admin@domain.com"-ForegroundColor Yellow
                            Connect-MsolService
                            try{
                                $mail = Read-host "`nMail"
                                $lic = (Get-MsolUser -userprincipalname $mail).licenses.AccountSkuID
                            }catch{
                                 Go 6 1  
                            }
                            switch ($lic)
                            {
                                        'fazergroup:ENTERPRISEPACK' {Write-Host "`nThe user has E3"-ForegroundColor Green
                                                                     Break   }
                                        'fazergroup:DESKLESS'       {Write-Host "`nThe user has K1"-ForegroundColor Green
                                                                     Break   }
                                        default {Write-Host "Info not found, user might not have a license" -ForegroundColor Red}
                                       
                            }
                           
                            Clear-variable mail
	                    
 
             }#6
            #7.Disable User
            #//////////////////////////////////////////////////////////////////////////////////////////  
            7 {
		                Write-Host "`nYou picked: 7.Disable an account" -ForegroundColor Yellow          
                        do{
                            $booleanCase6 = $FALSE
                            $user = Read-Host "`nUsername"
                            $ticketNumber = Read-Host "ticketNumber"
                            Write-Host "`n############################"
                            Write-Host '      Is this correct?' -ForegroundColor  Green
                            Write-Host "############################"
                            $userInfo = Get-ADUser -Properties $Properties -Filter {(samAccountName -like $user)}
                            Write-Host "`nYou want to disable the follwing user" -ForegroundColor  Red
                            $userInfo  | select $Properties  | Format-Table $myFormat -AutoSize
                            Write-Host "And update the following ticketNumber: $ticketNumber" -ForegroundColor  Red
                            $ChoiceCase6 = Read-Host "`nPress |1| to continue`nPress |0| to try again"
                            if($ChoiceCase6 -eq '1'){
                                $booleanCase6 = $TRUE
                            }
                        }while($booleanCase6 -eq $FALSE)

                        Disable-FazerUser -sAMAccountName $user -ticketNumber $ticketNumber -Verbose
                        Write-Host "User has been disabled" -ForegroundColor  Green
                        Clear-variable user
	        }#7
            #8.Generate a random password?
            #//////////////////////////////////////////////////////////////////////////////////////////
            8 {
                    
                    
    
                     #removed due to safty reasons   
                    
           
            }#8
            #9.Unlock account
            #//////////////////////////////////////////////////////////////////////////////////////////  
	        9 {
                        Write-Host "`nYou picked: 9.Unlock account" -ForegroundColor Yellow
                        $user = Read-Host "Username"
                        $lockOutStatus = Get-ADUser $user -Properties * | Select-Object -ExpandProperty LockedOut
                        if($lockOutStatus -eq $false){
                                Write-Host "`nThe account is not locked ¯\_(?)_/¯"-ForegroundColor Cyan
                         }
                         else{
                                Write-Host "`nThe account is locked, unlocking.."-ForegroundColor Cyan
                                Write-Host "Barabim, barabom"-ForegroundColor Cyan
                                Unlock-ADAccount $user 
                                Write-Host "`nAccount is unlocked" -ForegroundColor  Green
                         }
                         $expDate = Get-ADUser -Identity $user -Properties PasswordLastSet | select -ExpandProperty PasswordLastSet
                         $currentDate = Get-Date
                         $ts = NEW-TIMESPAN –Start $expDate –End $currentDate
                         $daysLeft = 89-$ts.Days

                         if ($ts.Days -ge 89){
                                Write-Host "`nPassword has expired, you need to change the users password" -ForegroundColor  Red
                           }else{
                                Write-Host "`nPassword has not expired, $daysLeft days until experation date" -ForegroundColor  Yellow
                           }
	        }#9
            #10.Change mail adress
            #////////////////////////////////////////////////////////////////////////////////////////// 
	        10 {
		                Write-Host "`nYou picked: 10.Change mail adress" -ForegroundColor Yellow
                        Write-Host "`n#### Examples ####" -ForegroundColor Green
                        help Set-FazerNewUserPrincipalName -Examples
                        Write-Host "`n#### Examples ####" -ForegroundColor Green
                        Set-FazerNewUserPrincipalName
                        Write-Host "Mail address changed" -ForegroundColor  Green

            }#10
            #11.Enable mail address?
            #//////////////////////////////////////////////////////////////////////////////////////////
            11 {
                        
                        Write-Host "`nYou picked: 11.Enable mail address" -ForegroundColor Yellow 
                        Write-Host "`nConnecting to OnPremise" -ForegroundColor Green 
                        Connect-OnPremise
                        Write-Host "`n#### Connected to OnPremise ####" -ForegroundColor Green 
                        Write-Host "`n#### Examples ####" -ForegroundColor Green
                        help Enable-RemoteUserMailbox -Examples
                        Write-Host "`n#### Examples ####" -ForegroundColor Green          
                        Enable-RemoteUserMailbox
            }#11
            #12.Create a new shared mailbox?
            #//////////////////////////////////////////////////////////////////////////////////////////
            12 {
                        Write-Host "`nYou picked: 12.Create a new shared mailbox" -ForegroundColor Yellow
                        Write-Host "`nConnecting to OnPremise" -ForegroundColor Green 
                        Connect-OnPremise
                        Write-Host "`n#### Connected to OnPremise ####" -ForegroundColor Green 
                        Write-Host "`n#### Examples ####" -ForegroundColor Green
                        help New-FazerSharedMailbox -Examples
                        Write-Host "`n#### Examples ####" -ForegroundColor Green           
                        New-FazerSharedMailbox
            }#12
            #13.Create a new DL-group?
            #//////////////////////////////////////////////////////////////////////////////////////////
            13 {
                        Write-Host "`nYou picked: 13.Create a new DL-group" -ForegroundColor Yellow
                        Write-Host "`n#### Connecting to OnPremise ####" -ForegroundColor Green 
                        Connect-OnPremise
                        Write-Host "`n#### Connected to OnPremise ####" -ForegroundColor Green 
                        Write-Host "`n#### Examples ####" -ForegroundColor Green 
                        help New-FazerDistributionList -Examples
                        Write-Host "`n#### Examples ####" -ForegroundColor Green
                        New-FazerDistributionList
            }#13
            #14.Users in a group?
            #//////////////////////////////////////////////////////////////////////////////////////////
            14 {
                        
                        Write-Host "`nYou picked: 14.Which users are in a group?" -ForegroundColor Yellow   
                        do{
                            $Group = Read-Host "`nGroup?"
                            $users = Get-ADGroupMember $Group -Recursive | select -ExpandProperty SamAccountName   
                        }while (!$Group)
                        foreach ($user in $users){
                               #To add values into an array
                               $outPut += @(Get-ADUser $user -Properties $Properties  | select $Properties) 
                        }
                        
                        Get-RezPut $outPut
            
            }#14
            #15.List accounts by country?
            #//////////////////////////////////////////////////////////////////////////////////////////
            15 {
                        Write-Host "`nYou picked: 15.List accounts by country" -ForegroundColor Yellow
                        #Start with a standard OU-path
                        $OUPath = ""#OU path removed due to safty reasons
                        do{
                            Write-Host "`nWhich users do you want to list"
                            Write-Host "`n1.Swedish users?"-ForegroundColor  Green
				            Write-Host "2.Danish users?" -ForegroundColor  Green
				            Write-Host "3.Norwegian users?"-ForegroundColor  Green

				            [int]$choiceCase14Country = Read-Host "`nPick a number"
                      
                        }while (!$choiceCase14Country)
                        #Depending on user input, change country
                        switch ($choiceCase14Country){

                            1{
                                 Write-Host "You picked 1.Swedish users" -ForegroundColor  Yellow
                            } 
                            2{
                                 $OUPath = $OUPath -replace "Sweden","Denmark"
                                 Write-Host "You picked 2.Danish users"-ForegroundColor  Yellow
                            }
                            3{
                                 $OUPath = $OUPath -replace "Sweden","Norway"
                                 Write-Host "You picked 3.Norwegian users"-ForegroundColor  Yellow
                            }

                        }
                         do{
                            Write-Host "`nWhich accounts do you want to list"
                            Write-Host "`n1.STD accounts ?"-ForegroundColor  Green
				            Write-Host "2.Personal accounts?" -ForegroundColor  Green
				            
				            [int]$choiceCase14Account = Read-Host "`nPick a number"
                      
                        }while (!$choiceCase14Account)
                        #Depending on userinput, change account-type
                        switch ($choiceCase14Account){

                            1{
                                 Write-Host "You picked 1.STD accounts" -ForegroundColor  Yellow
                            } 
                            2{
                                 $OUPath = $OUPath -replace "removed due to safty reasons","Win7x64"
                                 Write-Host "You picked 2.Personal accounts"-ForegroundColor  Yellow
                            }
                        }
                        $result = Get-ADUser -SearchBase $OUPath -Filter * -Properties $Properties 
                        Clear-variable OUPath

                         #Get the output function
                        Get-RezPut $result
                        
            
            }#15
            #16.Add an application to a computer
            #//////////////////////////////////////////////////////////////////////////////////////////
            16 {
                        Write-Host "`nYou picked: 16.Add an application to a computer" -ForegroundColor Yellow
                        $loop = $true
                        do{
                            $appName = Read-Host "Application"
                            $appName = "*" +  $appName + "*"
                            $Application = Get-ADGroup -SearchBase "OU path removed due to safty reasons"  -Filter {name -like $appName} | select -ExpandProperty name
                            if(!$Application){
                                Write-Host "Application does not exist, try again" -ForegroundColor Red
                                $loop = $true
                            }#if
                            else{
                                $loop = $false
                            }#if

                        }#do
                        while($loop)


                        if ($Application.count -ge 2)
                        {

                            foreach ($app in $Application){
                                Write-Host $Application.IndexOf($app) . $app 
                            }#foreach
                            $loop = $true
                            do {
                                $chosenApp = Read-Host "Select application from list"
                                if ([int]$chosenApp -lt 0 -or [int]$chosenApp -gt $Application.Length - 1) {
                                    Write-host "wrong Input, try again" -ForegroundColor Red
                                    $loop = $true
                                }#if
                                    else {
                                    $loop = $false
                                }#else

                            }#do
                            while ($loop)

                                $Application = $Application[$chosenApp]  
                        }#if
                        Write-Host "`nChosen application: $Application" -ForegroundColor Green
                        $continueOrNot = Read-Host "Is this correct? Type |1| to continue or |0| to retry"
                        switch ($continueOrNot){
 
                            1 {
                                $loop = $true
                                do{
                                    try{
                                        $computerName = Read-Host "`nWhich computer do you want to add?"

                                        Write-Host "`n####################### Computer info #######################" -ForegroundColor Green
                                        Get-ADComputer $computerName
                                        Write-Host "####################### Computer info #######################" -ForegroundColor Green
                
                                        $loop = $false
                                    }#try
                                    catch {
                                        Write-Host "`nCan't find the computer, try again" -ForegroundColor Red
                                        $loop = $true
                                    }#catch
                                }#do
                                while($loop)

                                Write-Host "`nAdding $computerName to $Application" -ForegroundColor Yellow
                                $applyChange = Read-Host "Continue? Type |1| to continue or |0| to retry"
                                switch($applyChange){
        
                                    1{
                                        $computerName = Get-ADComputer $computerName
                                        Add-ADGroupMember -identity $Application -Members $computerName.SamAccountName
                                        Write-Host "`n#####" $computerName.samAccountName.toupper() "(don't worry about the '$') added to $Application #####" -ForegroundColor Cyan  
 

        
                                    }#switch 1
                                    0 {
                                        go 16
                                      }#switch 2    
                                } #$applyChange switch
                            }#switch 1
                            0 {   
                                go 16         
                            }#switch 2
                        }# continueOrNot switch            
                }#16
            #17.Which groups is the user owner of?
            #//////////////////////////////////////////////////////////////////////////////////////////
            17 {
                        Write-Host "`nYou picked: 17. Which groups is the user owner of" -ForegroundColor Yellow
                        $loop = $true
                        do{
                            $userName = Read-Host "`nUsername"
                            $getUser = Get-ADUser $userName -Properties *
                            if(!$getUser){
                                $loop = $true
                                Write-Host "Can't find the user, try again" -ForegroundColor Red
                                
                            }#if
                            else{
                                $loop = $false
                            }#else
                        }#do
                        while($loop)
                        $getUser =  "*" + $getUser.GivenName + " " + $getUser.Surname + "*"
                        $groups = Get-ADGroup  -Properties * -Filter{Info -like $getUser }
                        Write-Host "The user is owner of the following groups:" 
                        $groups.name
            }#17
            #18.Is the user in the given group (nested groups included)?
            #//////////////////////////////////////////////////////////////////////////////////////////
            18 {
                        Write-Host "`nYou picked: 18. Is the user in the given group (nested groups included)?" -ForegroundColor Yellow
                         $Found = $false                
                         do{
                            $GroupName = Read-Host "`nGroup?"
                            $userToFind = Read-Host "Username?"
                            $users = Get-ADGroupMember $GroupName -Recursive | select -ExpandProperty SamAccountName   
                        }while (!$GroupName)
                        foreach ($user in $users){
                               #To add values into an array
                               if($user -eq $userToFind) {
                                  $Found = $true
                                  
                               }
                        }
                        if($Found){
                            Write-Host  "`n$userToFind found in this group or a nested group" -ForegroundColor Green  
                        }else{
                        
                            Write-Host  "`n$userToFind not found in this group or nested groups" -ForegroundColor Red  
                        
                        }

            }#18
            #19.Are all the user attributes OK?
            #//////////////////////////////////////////////////////////////////////////////////////////
            19 {
                        Write-Host "`nYou picked: 19. Are all the user attributes OK?" -ForegroundColor Yellow
                        
                        do{
                            $user = Read-Host "`nUsername"
                            if(!$user){
                            Write-Host "Wrong username, try again" -ForegroundColor Red
                            }
                            }while (!$user)
                            $userInfo = Get-ADUser $user -Properties * | select proxyAddresses, userprincipalname, mail, mailNickname, targetAddress


                            if(!$userInfo.proxyAddresses){
                                Write-Host "`nProxyaddress is blank, please check this" -ForegroundColor Red
                            }else{
                                Write-Host "`nProxy" -ForegroundColor Yellow
                                $userInfo.proxyAddresses
                            }
                            if(!$userInfo.targetAddress){
                                Write-Host "`nTargetaddress is blank, please check this" -ForegroundColor Red
                            }else{
                                Write-Host "`nTarget" -ForegroundColor Yellow
                                $userInfo.targetAddress
                            }
                            if(!$userInfo.userprincipalname){
                                Write-Host "`nUserPrincipalName is blank, please check this" -ForegroundColor Red
                            }else{   
                                Write-Host "`nUserPrincipalName" -ForegroundColor Yellow
                                $userInfo.userprincipalname
                            }
                            if(!$userInfo.mail){
                                Write-Host "`nMail is blank, please check this" -ForegroundColor Red
                            }else{ 
                                Write-Host "`nMail" -ForegroundColor Yellow
                                $userInfo.mail
                            }
                            if(!$userInfo.mailNickname){
                                Write-Host "`nMailNickname is blank, please check this" -ForegroundColor Red
                            }else{ 
                                Write-Host "`nMailNickname" -ForegroundColor Yellow
                                $userInfo.mailNickname
                            }


                            [ValidateSet('Yes','No')]$Answer = Read-Host  "`nDo you want to check the Office license aswell? Yes or no?"
                            if($Answer -eq "Yes")
                            {
                                    Write-Host "`nLogon to MsolService with your admin account, "-ForegroundColor Green -NoNewline
                                    Write-Host "example: admin@domain.com"-ForegroundColor Yellow
                                    Connect-MsolService
                                    try{
                                        $mail = $userInfo.userprincipalname
                                        $lic = (Get-MsolUser -userprincipalname $mail).licenses.AccountSkuID
                                    }catch{
                                         Go 6 1  
                                    }
                                    switch ($lic)
                                    {
                                        'fazergroup:ENTERPRISEPACK' {Write-Host "`nThe user has E3"-ForegroundColor Green
                                                                     Break   }
                                        'fazergroup:DESKLESS'       {Write-Host "`nThe user has K1"-ForegroundColor Green
                                                                     Break   }
                                        default {Write-Host "Info not found, user might not have a license" -ForegroundColor Red}
                                       
                                        
                                    }
                           
                                    Clear-variable mail
                            }
                            
                            
            }#19
            #20.Move user to win 10 OU in AD?
            #//////////////////////////////////////////////////////////////////////////////////////////
            20 {
                        Write-Host "`nYou picked: 20. Move user account (NOT rest. account) to win10 OU in AD" -ForegroundColor Yellow
                        
                        do{
                            $userInut= Read-Host "`nUsername"
                            $adUser = Get-ADUser $userInut
                            if(!$adUser){
                                Write-Host "`nUsername does not exist ¯\_(?)_/¯, try again"-ForegroundColor Red
                                }
                        }while(!$adUser)
                        
                        Move-ADObject  -Identity $adUser.distinguishedName -TargetPath "OU path removed due to safty reasons"
                        Write-Host "`n$userInut has been moved to Win10 OU" -ForegroundColor Green
                        Clear-variable adUser
                        $adUser = Get-ADUser $userInut
                        Write-Host "$adUser.distinguishedName" -ForegroundColor Cyan
                        Clear-variable adUser

            }#20
            #21.Add attributes to new account?
            #//////////////////////////////////////////////////////////////////////////////////////////
            21 {
                        Write-Host "`nYou picked: 21.Add attributes to new account" -ForegroundColor Yellow    
                        do{
                            $user =  Read-host "`nUsername"
                            $employeeID = Read-host "EmployeeID"
                           
                            $userAttributes = Get-ADUser -Identity $user -Properties EmployeeID,Pager,Description, givenName, surName 
                            if(!$userAttributes){
                                Write-Host "`nUsername does not exist ¯\_(?)_/¯, try again"-ForegroundColor Red
                                }
                        }while(!$user)

                        if( !$userAttributes.employeeID){
                            Write-Host "`nUser does not have and EmployeeID"-ForegroundColor Red
                            Set-ADUser -Identity $userAttributes.SamAccountName -EmployeeID $employeeID
                            $userAttributes = Get-ADUser -Identity $userAttributes.SamAccountName -Properties EmployeeID
                            if(!$userAttributes.employeeID){
                                Write-Host "`nSomething went wrong, could not add the EmployeeID ¯\_(?)_/¯" -ForegroundColor Green
                            }Else{
                                Write-Host "`nEmployeeID added" -ForegroundColor Green
                                $userAttributes  | select $Properties  | Format-Table $myFormat -AutoSize                          
                            }

                        }Else{
                            Write-Host "`nUser EmployeeID is: "$UserAttributes.employeeID -ForegroundColor Cyan
                            Write-Host "User EmployeeID is: "$UserAttributes.employeeID -ForegroundColor Cyan
                            Write-Host "No need to add EmployeeID" -ForegroundColor Green
                            
                        }
                        [ValidateSet('1','2')]$Answer = Read-Host  "`nDo you want to add pager aswell? |1|= YES or |2| = NO?"
                            if($Answer -eq "1")
                            {
                                 do{
                                    $pager =  Read-host "`nCost Center"
                                     
                                    if(!$pager){
                                        Write-Host "`nTry again"-ForegroundColor Red
                                    }
                                }while(!$pager)
                                
                                $description = $pager + " - " + $userAttributes.givenName + " " + $userAttributes.surName                                      
                                Set-ADUser -Identity $userAttributes.SamAccountName -Replace @{pager=$pager}
                                Set-ADUser -Identity $userAttributes.SamAccountName -description $description

                                Write-Host "`nCost Center added and description changed, see below"-ForegroundColor Cyan
                                $userAttributes  | select pager, description
                                    
                            }
                       
                        

             }#21
            
	    }#Main Switch
   
	
    Clear-variable choiceMain
}#Go END

