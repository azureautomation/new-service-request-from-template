<#
.SYNOPSIS 
     Creates a new Service Request with activity prefixes.

.DESCRIPTION
    This runbook will create a new Service Request and apply the template.

.PARAMETER SMCreds
    Credentials to connect to Service Manager.

.PARAMETER SMServer
    Service Manager Management Server with SMLets installed.

.PARAMETER TemplateID
    Template GUID of the Template you want to apply to the New Service Request.
    

.EXAMPLE
    $SMCreds = Get-AutomationPSCredential -Name 'SCSMAdmin'
    SM-CreateSRFromTemplate `
        -SMCreds $SMCreds `
        -SMServer SM01 `
        -TemplateID 574d7d8e-44c6-47a8-1834-291247e4531c `
        
#>
workflow New-SRFromTemplate
{
    [OutputType([object])]
    
    Param(
        [Parameter(Mandatory=$True)]
        [string]$TemplateID,
        [Parameter(Mandatory=$True)]
        [string]$SMServer,
        [Parameter(Mandatory=$false)]
        [string]$Title,
        [Parameter(Mandatory=$True)]
        [PSCredential]$SMCreds
    )
    
    $SR = InlineScript
    {
        Import-Module smlets
        function Get-SCSMObjectPrefix
        {
		  Param ([string]$ClassName =$(throw "Please provide a classname"))
    		# Function Started
    		##LogTraceMessage "*** Function Get-SCSMObjectPrefix Started ***"
    		switch ($ClassName)
    		{
    		default
    		{
    		##LogTraceMessage "Get prefix from Activity Settings"
    		#Get prefix from Activity Settings
    		if ($ClassName.StartsWith("System.WorkItem.Activity") -or $ClassName.Equals("Microsoft.SystemCenter.Orchestrator.RunbookAutomationActivity"))
    		{
    			$ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Id "5e04a50d-01d1-6fce-7946-15580aa8681d") -ComputerName $using:SMServer
     
    			if ($ClassName.Equals("System.WorkItem.Activity.ReviewActivity")) {$prefix = $ActivitySettingsObj.SystemWorkItemActivityReviewActivityIdPrefix}
    			if ($ClassName.Equals("System.WorkItem.Activity.ManualActivity")) {$prefix = $ActivitySettingsObj.SystemWorkItemActivityManualActivityIdPrefix}
    			if ($ClassName.Equals("System.WorkItem.Activity.ParallelActivity")) {$prefix = $ActivitySettingsObj.SystemWorkItemActivityParallelActivityIdPrefix}
    			if ($ClassName.Equals("System.WorkItem.Activity.SequentialActivity")) {$prefix = $ActivitySettingsObj.SystemWorkItemActivitySequentialActivityIdPrefix}
    			if ($ClassName.Equals("System.WorkItem.Activity.DependentActivity")) {$prefix = $ActivitySettingsObj.SystemWorkItemActivityDependentActivityIdPrefix}
    			if ($ClassName.Equals("Microsoft.SystemCenter.Orchestrator.RunbookAutomationActivity")) {$prefix = $ActivitySettingsObj.MicrosoftSystemCenterOrchestratorRunbookAutomationActivityBaseIdPrefix
    			}
    		}
    			else {throw "Class Name $ClassName is not supported"}
    		}
    	}
    	return $prefix
    	
    	}
        
        function Update-SCSMPropertyCollection
    	{
    		Param ([Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateObject]$Object =$(throw "Please provide a valid template object"))
     
    		##LogTraceMessage "********** Function Update-SCSMPropertyCollection Started***********"
    
    		#Regex - Find class from template object property between ! and ']
    		$pattern = '(?<=!)[^!]+?(?=''\])'
    		if (($Object.Path) -match $pattern -and ($Matches[0].StartsWith("System.WorkItem.Activity") -or $Matches[0].StartsWith("Microsoft.SystemCenter.Orchestrator")))
    		{
    			##LogTraceMessage "Set prefix from activity class"
    			#Set prefix from activity class
    			$prefix = Get-SCSMObjectPrefix -ClassName $Matches[0] -ComputerName $using:SMServer
    
    			##LogTraceMessage "Create Template propertiy object" 
    			#Create template property object
    			$propClass = [Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateProperty]
    			$propObject = New-Object $propClass
     
    			##LogTraceMessage "Add New item to property object"
    			#Add new item to property object
    			$propObject.Path = "`$Context/Property[Type='$alias!System.WorkItem']/Id$"
    			$propObject.MixedValue = "$prefix{0}"
     
    			##LogTraceMessage "Add property to template"
    			#Add property to template
    			$Object.PropertyCollection.Add($propObject)
     
    			##LogTraceMessage "recursively update activities in activities"
    			#recursively update activities in activities
    			if ($Object.ObjectCollection.Count -ne 0)
    			{
    				foreach ($obj in $Object.ObjectCollection)
    				{ 
    					Update-SCSMPropertyCollection -Object $obj
    				}       
    			}
    		}
    		##LogTraceMessage "********** Function Update-SCSMPropertyCollection Finished***********"
    	}
        
        function Apply-SCSMTemplate
    	{
    		Param ([Microsoft.EnterpriseManagement.Common.EnterpriseManagementObjectProjection]$Projection =$(throw "Please provide a valid projection object"),
    		[Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplate]$Template = $(throw 'Please provide an template object, ex. -template template'))
    
    		##LogTraceMessage "********** Apply-SCCMTemplate Function Started***********"
    
    		##LogTraceMessage "Get alias from system.workitem.library managementpack to set id property" 
    		#Get alias from system.workitem.library managementpack to set id property
    		$templateMP = $Template.GetManagementPack()
    		$alias = $templateMP.References.GetAlias((Get-SCSMManagementPack system.workitem.library))
    		 ##LogTraceMessage "Update Activites in Template"
    		#Update Activities in template
    		foreach ($TemplateObject in $Template.ObjectCollection)
    		{
    			Update-SCSMPropertyCollection -Object $TemplateObject
    		}
    		##LogTraceMessage "Apply update Tempalte"
    		#Apply update template
    		Set-SCSMObjectTemplate -Projection $Projection -Template $Template -ErrorAction Stop
    		#Write-Host "Successfully applied template:`n"$template.DisplayName "`nTo:`n"$Projection.Object
    		##LogTraceMessage "********** Apply-SCCMTemplate Function Finished***********"
    	}
        
        Function Create-SR
    	{
    		##LogTraceMessage "********* Create-SR Funtion Started***********"
    		
    		$SRClass = Get-SCSMClass -name System.WorkItem.ServiceRequest$ -ComputerName $using:SMServer
    		##LogTraceMessage "Variable SRClass set to $SRClass"
    		$Params = @{ID="SR{0}"
    		Title = $using:Title
    		Description = $using:Title
    		Status = "New"
    		}
    		##LogTraceMessage "Variable Params set to $Params"
    		$emo = New-SCSMObject -class $SRClass -PropertyHashtable $Params -pass -ComputerName $using:SMServer
    		##LogTraceMessage "Created WI $emo.ID"
    		If($emo -ne $null)
    		{
    			$SRID = $emo.ID
    			##LogTraceMessage "Work Item $script:SRID was created"
    			#determine projection according to workitem type
    			switch ($emo.GetLeastDerivedNonAbstractClass().Name)
    			{
    				"System.workitem.Incident" {$projName = "System.WorkItem.Incident.ProjectionType" }
    				"System.workitem.ServiceRequest" {$projName = "System.WorkItem.ServiceRequestProjection"}
    				"System.workitem.ChangeRequest" {$projName = "System.WorkItem.ChangeRequestProjection"}
    				"System.workitem.Problem" {$projName = "System.WorkItem.Problem.ProjectionType"}
    				"System.workitem.ReleaseRecord" {$projName = "System.WorkItem.ReleaseRecordProjection"}
     
    				default {throw "$emo is not a supported workitem type"}
    			}
    			##LogTraceMessage "Get the new SR ID"
    			#Get object projection
    			$emoID = $emo.id
    			##LogTraceMessage "Variable emoID set to $emoID"
    			$WIproj = Get-SCSMObjectProjection -ProjectionName $projName -Filter "Id -eq $emoID" -ComputerName $using:SMServer
    			##LogTraceMessage "Variable WIproj set to $WIproj"
    
    			##LogTraceMessage "get the Template" 
    			#Get template from displayname or id
    			#if ($TemplateDisplayName)
    			#{
    			#$template = Get-SCSMObjectTemplate -DisplayName $TemplateDisplayName
    			#}
    			#else
    			if ($using:templateId)
    			{
    				$template = Get-SCSMObjectTemplate -id $using:templateId -ComputerName $using:SMServer
    				LogTraceMessage "Variable template set to $template"
    			}
    			else
    			{
    				LogTraceMessage "Please provide either a template id or a template displayname to apply"
    			}
    
    			LogTraceMessage "Apply Template to the New SR" 
    			#Execute apply-template function if id and 1 template exists
    			if (@($template).count -eq 1)
    			{
     
    				if ($WIProj)
    				{
    					Apply-SCSMTemplate -Projection $WIproj -Template $template
                        return $SRID
    				}
    				else
    					{throw "Id $Id cannot be found";}
    			}
    			else{throw "Template cannot be found or there was more than one result"}
    			}
    		else
    		{
    			throw "Work Item was not created sucessfull"
    		}
    		##LogTraceMessage "********* Create-SR Funtion Finished***********"
    	}
        Create-SR
        
    } -PSCredential $SMCreds -PScomputerName $SMServer
    $SR 
}