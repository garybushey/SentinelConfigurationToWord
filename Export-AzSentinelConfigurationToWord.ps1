#requires -version 6.2
<#
    .SYNOPSIS
        This command will generate a Word document containing the information about all the Azure Sentinel
        Solutions.  
    .DESCRIPTION
       This command will generate a Word document containing the information about all the Azure Sentinel
        Solutions.  
    .PARAMETER WorkspaceName
        Enter the workspace name to use. 
    .PARAMETER ResourceGroupName
        Enter the Resource Group name to use.  
    .PARAMETER FileName
        Enter the file name to use.  Defaults to "MicrosoftSentinelSolutions.docx"  ".docx" will be appended to all filenames if needed
    .NOTES
        AUTHOR: Gary Bushey
        LASTEDIT: 19 October 2023
    .EXAMPLE
        Export-AAzSentinelConfigurationToWork -WorkspaceName testwg -ResourceGroupName rgName
        In this example you will get the file named "MicrosoftSentinelReport.docx" generated containing all configuration for the Sentinel instance
    .EXAMPLE
        Export-AzSentinelSolutionsToWord -WorkspaceName testwg -ResourceGroupName rgName -fileName "test"
        In this example you will get the file named "test.docx" generated containing all configuration for the Sentinel instance
   
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$WorkSpaceName ,

    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [string]$FileName = "MicrosoftSentinelReport.docx"
)

Function Export-AzSentinelConfigurationToWord($fileName) {

    #Setup the Authentication header needed for the REST calls
    $context = Get-AzContext
    $myProfile = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile
    $profileClient = New-Object -TypeName Microsoft.Azure.Commands.ResourceManager.Common.RMProfileClient -ArgumentList ($myProfile)
    $token = $profileClient.AcquireAccessToken($context.Subscription.TenantId)
    $authHeader = @{
        'Content-Type'  = 'application/json' 
        'Authorization' = 'Bearer ' + $token.AccessToken 
    }
    $SubscriptionId = (Get-AzContext).Subscription.Id
    $baseUrl = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$workspaceName/providers/Microsoft.SecurityInsights/"
    $apiVersion = "?api-version=2023-09-01-preview"
    #$WorkspaceID = (Get-AzOperationalInsightsWorkspace -Name $workspaceName -ResourceGroupName $resourceGroupName).CustomerID
        

    try {

        #Load the list of all  solutions
        $url = $baseUrl + "contentProductPackages" + $apiVersion
        $allSolutions = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
        # We should be able to get the deployed solutions using the "ContentPackages" call, but doesn't quite work right
        # so I determine the installed solutions but those that have an installedVersion field.
        $solutions = $allSolutions | Where-Object { $null -ne $_.properties.installedVersion }

        #Load the list of all the installed  templates
        $url = $baseUrl + "contentProductTemplates" + $apiVersion
        $solutionTemplates = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

        #Load all the metadata entries
        $url = $baseUrl + "metadata" + $apiVersion
        $metadata = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

        #Load all the alert rules entries
        $url = $baseUrl + "alertrules" + $apiVersion
        $alertRules = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value


        #Setup Word
        $word = New-Object -ComObject word.application
        $word.Visible = $false      #We don't want to see word while this is running
        $doc = $word.documents.add()
        $selection = $word.Selection        #This is where we will add all the text and formatting information

        #Add the Title page
        Add-TitlePage $selection
        $selection.InsertBreak(7)   #page break

        #Leaving the creation of the TOC here so we can update it at the end 
        #Tables of Contents.  Will need to update when everything else is added
        $range = $selection.Range
        #Note that you need to reference tableSofcontents (Note the "S" in tables, rather than tableofcontents)
        #Not going to say how long it took me to get this right ;)
        $toc = $doc.TablesOfContents.add($range, $true, 1, 3)
        $selection.TypeParagraph()
        $selection.InsertBreak(7)   #page break

        #Add all the solutions
        Add-AllSolutions $solutions $selection 
        $selection.InsertBreak(7)   #page break

        #The REST API to get the templates will only show those templates that were installed using the solution
        #installation.  If they were created prior to moving to the solution model, they don't show up so we need to look
        #at the intalled solutions.
        Add-AllWorkbooks $solutions $solutionTemplates $metadata $selection 
        $selection.InsertBreak(7)   #page break

        #Add all the Hunts that have been created
        Add-AllHunts $selection
        $selection.InsertBreak(7)   #page break

        #Add the Threat Intelligence Metrics
        # ***  THE REST API IS NOT WORKING AS ADVERTISED  ***
        # When it is working, I will add the code.

        #Add the Repository information
        Add-AllRepositories $selection
        $selection.InsertBreak(7)   #page break

        #Add the Data Connectors
        Add-AllDataConnectors $solutions $solutionTemplates $selection
        $selection.InsertBreak(7)   #page break

        #Add the Analytic Rules
        Add-AllAnalyticRules  $alertRules $selection
        $selection.InsertBreak(7)   #page break

        #Add the Watchlists
        Add-AllWatchlists $selection
        $selection.InsertBreak(7)   #page break

        #Add the Automation rules
        Add-AllAutomation $alertRules $selection
        $selection.InsertBreak(7)   #page break

        #Add the Settings
        Add-AllSettings $selection
        $selection.InsertBreak(7)   #page break

        $toc.Update()   #Update the Tables of contents
        $outputPath = Join-Path $PWD.Path $fileName
        $doc.SaveAs($outputPath)    #Save the document
        $doc.Close()                #Close the document
        $word.Quit()                #Quit word
        #NOTE:  If for some reason the Powershell quits before it can quit work, go into task manager
        #to close manually, otherwise you can get some weird results
    }
    catch {
        $errorReturn = $_
        Write-Error $errorReturn
        $doc.Close()
        $word.Quit()
    }
}

Function Add-AllSettings ($selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Settings")
    $selection.TypeParagraph()

    Write-Host "Settings: " -NoNewline

    $url = $baseUrl + "settings" + $apiVersion
    $settings = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

    $selection.Style = "Heading 2"
    $selection.TypeText("User Entity Behavior Analytics")
    $selection.TypeParagraph()
    $selection.TypeText("Enabled: ")
    $entityAnalytics = $settings | Where-Object -Property Name -eq "EntityAnalytics"
    if ($null -ne $entityAnalytics) {
        $selection.TypeText("True")
        $selection.TypeParagraph()
        $selection.TypeText("Directory services:")
        $selection.TypeParagraph()
        foreach ($entityProvider in $entityAnalytics.properties.entityProviders) {
            $text = TranslateServices $entityProvider
            $selection.TypeText("`t" + $text)
            $selection.TypeParagraph()
        }
        $selection.TypeParagraph()
        $euba = $settings | Where-Object -Property Name -eq "Ueba"
        $selection.TypeText("Directory sources:")
        $selection.TypeParagraph()
        foreach ($dataSource in $euba.properties.dataSources) {
            $text = TranslateServices $dataSource
            $selection.TypeText("`t" + $text)
            $selection.TypeParagraph()
        }
    }
    else {
        $selection.TypeText("False")
    }
    $selection.TypeParagraph()

    $selection.Style = "Heading 2"
    $selection.TypeText("Anomalies")
    $selection.TypeParagraph()
    $selection.TypeText("Enabled: ")
    $anomalies = $settings | Where-Object -Property Name -eq "Anomalies"
    if ($null -ne $anomalies) {
        $selection.TypeText("True")
    }
    else {
        $selection.TypeText("False")
    }
    $selection.TypeParagraph()

    $url = $baseUrl + "workspaceManagerConfigurations" + $apiVersion
    $settings = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

    $selection.Style = "Heading 2"
    $selection.TypeText("Is workspace a central workspace?")
    $selection.TypeParagraph()
    if ($settings.properties.mode -eq "Enabled") {
        $selection.TypeText("True")
    }
    else {
        $selection.TypeText("False")
    }
    $selection.TypeParagraph()

    $selection.Style = "Heading 2"
    $selection.TypeText("Allow Microsoft Sentinel engineers to access your data?")
    $selection.TypeParagraph()
    $selection.TypeText("Enabled: ")
    $anomalies = $settings | Where-Object -Property Name -eq "EyesOn"
    if ($null -ne $anomalies) {
        $selection.TypeText("True")
    }
    else {
        $selection.TypeText("False")
    }
    $selection.TypeParagraph()

    #Check to see which Resource Groups are set to run playbooks
    $selection.Style = "Heading 2"
    $selection.TypeText("Resource Groups that can run playbooks")
    $selection.TypeParagraph()
    $url = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups?api-version=2020-10-01" 
    $resourceGroups = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
    foreach ($resourceGroup in $resourceGroups) {
        $rgName = $resourceGroup.name
        $url = "https://management.azure.com/subscriptions/$subScriptionId/resourceGroups/$rgName/providers/Microsoft.Authorization/" +
        "roleAssignments?api-version=2015-07-01&%24filter=atScope()%20and%20assignedTo('4fae0573-3b67-4b20-98d1-5847c5faa905')"
        $foundRG = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
        if ($null -ne $foundRg.properties) {
            $selection.TypeText("`t" + $rgName)
            $selection.TypeParagraph()
        }
    }

    #Audit and HEalth monitoring
    $selection.Style = "Heading 2"
    $selection.TypeText("Auditing and Health Monitoring")
    $selection.TypeParagraph()
    $url = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/" +
    "workspaces/$workspaceName/providers/Microsoft.SecurityInsights/settings/SentinelHealth/providers/microsoft.insights/diagnosticSettings?api-version=2021-05-01-preview"
    $audit = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
    $auditLogs = $audit.properties.logs
    $allLogs = $auditLogs | Where-Object -Property categoryGroup -eq "allLogs"
    if ($null -ne $allLogs) {
        $selection.TypeText("`tAll Logs")
    }
    else {
        foreach ($log in $auditLogs) {
            $name = TranslateServices $log.Category
            $selection.TypeText("`t" + $name)
            $selection.TypeParagraph()
        }
    }
    $selection.TypeParagraph()
    Write-Host "Done"
}

Function TranslateServices ($text) {
    $returnText = $text
    switch ($text) {
        "AzureActiveDirectory" { $returnText = "Azure Active Directory" }
        "ActiveDirectory" { $returnText = "Active Directory" }
        "AuditLogs" { $returnText = "Audit Logs" }
        "AzureActivity" { $returnText = "Azure Activity" }
        "SecurityEvent" { $returnText = "Security Events" }
        "SignInLogs" { $returnText = "Sign In Logs" }
        "DataConnectors" { $returnText = "Data Collection - Connectors" }
    }
    return $returnText
}

Function Add-AllAutomation ($analyticRules, $selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Automation Rules")
    $selection.TypeParagraph()

    Write-Host "Automation Rules: " -NoNewline

    $url = $baseUrl + "automationRules" + $apiVersion
    $automationRules = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

    $Table = $doc.Tables.add($selection.Range, $automationRules.count + 1, 6) 
    $Table.cell(1, 1).range.Bold = 1
    $Table.cell(1, 1).range.text = "Order"
    $Table.cell(1, 2).range.Bold = 1
    $Table.cell(1, 2).range.text = "Display Name"
    $Table.cell(1, 3).range.Bold = 1
    $Table.cell(1, 3).range.text = "Trigger"
    $Table.cell(1, 4).range.Bold = 1
    $Table.cell(1, 4).range.text = "Analytic Rule Name"
    $Table.cell(1, 5).range.Bold = 1
    $Table.cell(1, 5).range.text = "Actions"
    $Table.cell(1, 6).range.Bold = 1
    $Table.cell(1, 6).range.text = "Expiration Date"

    $count = 1
    foreach ($automationRule in $automationRules | Sort-Object { $_.properties.order }) {
        $count++
        $Table.cell($count, 1).range.Bold = 0
        $Table.cell($count, 1).range.text = $automationRule.properties.order.toString()
        $Table.cell($count, 2).range.Bold = 0
        $Table.cell($count, 2).range.text = $automationRule.properties.displayName
        $Table.cell($count, 3).range.Bold = 0
        $trigger = $automationRule.properties.triggeringLogic.triggersOn + " " + $automationRule.properties.triggeringLogic.triggersWhen
        $Table.cell($count, 3).range.text = $trigger
        $Table.cell($count, 4).range.Bold = 0
        if ($automationRule.properties.triggeringLogic.conditions.conditionProperties.propertyName -eq "AlertAnalyticRuleIds") {
            $ruleNames = ""
            foreach ($rule in $automationRule.properties.triggeringLogic.conditions.conditionProperties.propertyValues) {
                $ruleDescription = ($analyticRules | Where-Object -Property Id -eq $rule).properties.displayName
                $ruleNames += $ruleDescription + "`n"
            }
            $Table.cell($count, 4).range.text = $ruleNAmes
        }
        else {
            $Table.cell($count, 4).range.text = "All"
        }
        Write-Host "." -NoNewline
        
        $Table.cell($count, 5).range.Bold = 0
        $Table.cell($count, 5).range.text = $automationRule.properties.displayNAme
        $Table.cell($count, 6).range.Bold = 0
        $expiration = $automationRule.properties.expirationTimeUtc
        if ($null -eq $expiration) {
            $Table.cell($count, 6).range.text = "Indefinite"
        }
        else {
            $Table.cell($count, 6).range.text = $expiration.toString()
        }
    }
    $selection.endKey(6) | Out-Null
    $selection.TypeParagraph()
    Write-Host ""
}

Function Add-AllWatchlists ($selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Watchlists")
    $selection.TypeParagraph()

    Write-Host "Watchlists: " -NoNewline

    $url = $baseUrl + "watchlists" + $apiVersion
    $watchlists = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

    $Table = $doc.Tables.add($selection.Range, $watchlists.count + 1, 5) 
    $Table.cell(1, 1).range.Bold = 1
    $Table.cell(1, 1).range.text = "Name"
    $Table.cell(1, 2).range.Bold = 1
    $Table.cell(1, 2).range.text = "Alias"
    $Table.cell(1, 3).range.Bold = 1
    $Table.cell(1, 3).range.text = "Source"
    $Table.cell(1, 4).range.Bold = 1
    $Table.cell(1, 4).range.text = "Created Time"
    $Table.cell(1, 5).range.Bold = 1
    $Table.cell(1, 5).range.text = "Last Updated"

    $count = 1
    foreach ($watchlist in $watchlists | Sort-Object { $_.properties.displayName }) {
        $count++
        $Table.cell($count, 1).range.Bold = 0
        $Table.cell($count, 1).range.text = $watchlist.properties.displayName
        $Table.cell($count, 2).range.Bold = 0
        $Table.cell($count, 2).range.text = $watchlist.properties.watchlistAlias
        $Table.cell($count, 3).range.Bold = 0
        $Table.cell($count, 3).range.text = $watchlist.properties.source
        $Table.cell($count, 4).range.Bold = 0
        $Table.cell($count, 4).range.text = $watchlist.properties.created.toString()
        $Table.cell($count, 5).range.Bold = 0
        $Table.cell($count, 5).range.text = $watchlist.properties.updated.toString()
        Write-Host "." -NoNewline
    }
    $selection.endKey(6) | Out-Null
    $selection.TypeParagraph()
    Write-Host " "
}

Function Add-AllAnalyticRules ($analyticRules, $selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Analytic Rules")
    $selection.TypeParagraph()

    Write-Host "Analytic Rules: " -NoNewline

    #$url = $baseUrl + "alertRules" + $apiVersion
    #$analyticRules = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value

    $Table = $doc.Tables.add($selection.Range, $analyticRules.count + 1, 6)
    $Table.cell(1, 1).range.Bold = 1
    $Table.cell(1, 1).range.text = "Severity"
    $Table.cell(1, 2).range.Bold = 1
    $Table.cell(1, 2).range.text = "Name"
    $Table.cell(1, 3).range.Bold = 1
    $Table.cell(1, 3).range.text = "Rule Type"
    $Table.cell(1, 4).range.Bold = 1
    $Table.cell(1, 4).range.text = "Status"
    $Table.cell(1, 5).range.Bold = 1
    $Table.cell(1, 5).range.text = "Tactics"
    $Table.cell(1, 6).range.Bold = 1
    $Table.cell(1, 6).range.text = "Techniques"
    $count = 1
    foreach ($analyticRule in $analyticRules) {
        $count++
        $Table.cell($count, 1).range.Bold = 0
        $Table.cell($count, 1).range.text = $analyticRule.properties.Severity
        $Table.cell($count, 2).range.Bold = 0
        $Table.cell($count, 2).range.text = $analyticRule.properties.displayName
        $Table.cell($count, 3).range.Bold = 0
        $Table.cell($count, 3).range.text = $analyticRule.kind
        $Table.cell($count, 4).range.Bold = 0
        if ($analyticRule.properties.enabled -eq "True") {
            $Table.cell($count, 4).range.text = "Enabled"
        }
        else {
            $Table.cell($count, 4).range.text = "Disabled"
        }
        $tactics = ""
        foreach ($tactic in $analyticRule.properties.tactics) {
            $tactics = $tactics + $tactic + "`n"
        }
        $Table.cell($count, 5).range.Bold = 0
        $Table.cell($count, 5).range.text = $tactics
        $techniques = ""
        foreach ($technique in $analyticRule.properties.Techniques) {
            $techniques = $techniques + $technique + "`n"
        }
        $Table.cell($count, 6).range.Bold = 0
        $Table.cell($count, 6).range.text = $techniques
        Write-Host "." -NoNewline
    }
    $selection.endKey(6) | Out-Null
    $selection.TypeParagraph()
    Write-Host " "
}

#For some reason the SAP data connector will not work with this scenario.
Function Add-AllDataConnectors ($solutions, $solutionTemplates, $selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Data Connectors")
    $selection.TypeParagraph()

    Write-Host "Data Connectors: " -NoNewline

    $url = $baseUrl + "dataConnectors" + $apiVersion
    $dataConnectors = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
    #Doing some translations
    foreach ($dataConnector in $dataConnectors) {
        if ($dataConnector.properties.connectorUiConfig.title -eq "Office 365") {
            $dataConnector.properties.connectorUiConfig.title = "Microsoft 365 (formerly, Office 365)"
        }
        if ($dataConnector.properties.connectorUiConfig.title -eq "SAP") {
            $dataConnector.properties.connectorUiConfig.title = "Microsoft Sentinel for SAP"
        }

    }
    foreach ($dataConnector in $dataConnectors | Sort-Object { $_.properties.connectorUiConfig.title }) {
        if ($null -ne $dataConnector.Properties.connectorUiConfig.title) {
            $selection.Style = "Heading 2"
            $selection.TypeText($dataConnector.Properties.connectorUiConfig.title)
            # $connected = Is-Connected($dataConnector.properties.connectorUiConfig)
            # $selection.TypeText("`t`t" + $connected)
            $selection.TypeParagraph()
            Write-Host "." -NoNewline
        }
    }
    Write-Host ""
}

<# Function Is-Connected($connectorUiConfig) {
    $isConnected = $false
    $queryTable = $connectorUiConfig.graphQueriesTableName
    $queries = $connectorUiConfig.connectivityCriterias.value
    foreach ($query in $queries) {
        if ($query.contains("{{")) {
            $query = $query.Replace("{{graphQueriesTableName}}", $queryTable)
        }
        if (!$query.contains("summarize LastLogReceived")) {
            $query += "| summarize LastLogReceived = max(TimeGenerated)| project IsConnected = LastLogReceived > ago(30d)"
        }
        try {
            $kqlQuery = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceID.GUID -Query $query 
        }
        catch 
        {}
        if ($kqlQuery.Results.IsConnected -eq "true") {
            $isConnected = $true
            break
        }
    }
    return $isConnected
} #>

Function Add-AllRepositories ($selection) {
    $url = $baseUrl + "sourcecontrols" + $apiVersion
    $repos = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
    $selection.Style = "Heading 1"
    $selection.TypeText("Repositories")
    $selection.TypeParagraph()

    Write-Host "Repositories: " -NoNewline

    foreach ($repo in $repos) {
        $selection.Style = "Heading 2"
        $selection.TypeText($repo.properties.displayName)
        $selection.TypeParagraph()
        $selection.Font.Bold = $true
        $selection.TypeText("Description: ");
        $selection.Font.Bold = $false
        $selection.TypeText($repo.properties.description)
        $selection.TypeParagraph()
        $selection.Font.Bold = $true
        $selection.TypeText("Source Control: ");
        $selection.Font.Bold = $false
        $selection.TypeText($repo.properties.repoType)
        $selection.TypeParagraph()
        $selection.Font.Bold = $true
        $selection.TypeText("Repository URL: ");
        $selection.Font.Bold = $false
        $selection.TypeText($repo.properties.repository.url)
        $selection.TypeParagraph()
        $selection.Font.Bold = $true
        $selection.TypeText("Content Types: ")
        $selection.Font.Bold = $false
        $selection.TypeParagraph()
        foreach ($type in $repo.properties.contentTypes) {
            $selection.TypeText("`t" + $type)
            $selection.TypeParagraph()
        }
        $selection.TypeParagraph()
        Write-Host "." -NoNewline
    }
    Write-Host " "
}

Function Add-AllHunts($selection) {
    $url = $baseUrl + "hunts" + $apiVersion
    $hunts = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
    $selection.Style = "Heading 1"
    $selection.TypeText("Hunts")
    $selection.TypeParagraph()

    Write-Host "Hunts: " -NoNewline

    foreach ($hunt in $hunts | Sort-Object { $_.properties.displayName }) {
        $selection.Style = "Heading 2"
        $selection.TypeText($hunt.properties.displayName)
        $selection.TypeParagraph()
        $selection.TypeText($hunt.properties.description)
        $selection.TypeParagraph()
        $selection.TypeText("Status: " + $hunt.properties.status);
        $selection.TypeText("`tHypothesis: " + $hunt.properties.hypothesisStatus);
        #$selection.TypeParagraph()
        if ($null -ne $hunt.properties.owner.assignedTo) {
            $selection.TypeText("`tAssigned To: " + $hunt.properties.owner.assignedTo);
        }
        else {
            $selection.TypeText("`tAssigned To: <unassigned>");
        }
        $selection.TypeParagraph()
        Write-Host "." -NoNewline
    }
    Write-Host ""
}
Function Add-TitlePage ($selection) {
    #Title Page
    $selection.Style = "Title"
    $selection.ParagraphFormat.Alignment = 1  #Center
    $selection.TypeText("Microsoft Sentinel Documentation")   #Add the text
    $selection.TypeParagraph()                                          #create a new paragraph
    $selection.TypeParagraph()
    $selection.Style = "Normal"
    $selection.ParagraphFormat.Alignment = 1  #Center
    $text = Get-Date
    $selection.TypeText("Created: " + $text)
    $selection.TypeParagraph()
    $selection.TypeParagraph()
    $selection.TypeText("Resource Group: " + $resourceGroupName)
    $selection.TypeParagraph()
    $selection.TypeText("Workspace Name: " + $workspaceName)
    $selection.TypeParagraph()

    $Table = $doc.Tables.add($selection.Range, 1, 1)
    $Table.cell(1, 1).range.Bold = 1
    $Table.cell(1, 1).range.text = ""
$selection.endKey(6) | Out-Null
    $selection.TypeParagraph()
}

Function Add-AllWorkbooks ($solutions, $solutionTemplates, $metadata, $selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Custom Workbooks")
    $selection.TypeParagraph()

    Write-Host "Custom Workbooks: " -NoNewline

    #Show "MyWorkbooks"
    $url = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Insights/workbooks?api-version=2018-06-17-preview&canFetchContent=false&%24filter=sourceId%20eq%20'%2Fsubscriptions%2F$SubscriptionId%2Fresourcegroups%2F$ResourceGroupName%2Fproviders%2Fmicrosoft.operationalinsights%2Fworkspaces%2F$WorkspaceName'&category=sentinel"
    $customWorkbooks = (Invoke-RestMethod -Method "Get" -Uri $url -Headers $authHeader ).value
   
    $Table = $doc.Tables.add($selection.Range, $customWorkbooks.count + 1, 3)
    $Table.cell(1, 1).range.Bold = 1
    $Table.cell(1, 1).range.text = "Name"
    $Table.cell(1, 2).range.Bold = 1
    $Table.cell(1, 2).range.text = "Content Source"
    $Table.cell(1, 3).range.Bold = 1
    $Table.cell(1, 3).range.text = "Content Name"
    $count = 1
    foreach ($workbook in $customWorkbooks | Sort-Object { $_.properties.displayName }) {
        $count += 1
        $Table.Cell($count, 1).range.text = $workbook.properties.displayName
        $foundTemplate = $metadata | Where-Object { $_.name -eq "workbook-" + $workbook.name }
        if ($null -ne $foundTemplate) {
            $Table.Cell($count, 2).range.text = "ContentHub"
            $Table.Cell($count, 3).range.text = $foundTemplate.properties.source.name
        }
        else {
            $Table.Cell($count, 2).range.text = "Custom"
            $Table.Cell($count, 3).range.text = "--"
        }
        $selection.TypeParagraph()
        Write-Host "." -NoNewline
    }
    $selection.endKey(6) | Out-Null
    $selection.TypeParagraph()

    $selection.Style = "Heading 1"
    $selection.TypeText("Workbook Templates")
    $selection.TypeParagraph()
    $selection.TypeText(" ")
    $selection.TypeParagraph()
    $sortedTemplates = @()

    Write-Host " " 
    Write-Host "Workbook Templates: " -NoNewline


    #Show Templates
    #Filter to get only the workbook tempalates
    $workbookTemplates = $solutionTemplates | Where-Object { $_.properties.contentKind -eq "Workbook" }

    #For all the installed oslutions, see if any of the solutions template's contentId matches the workbook's packageId
    #If so, add the information to an array so we can display the entries alphabetically 
    foreach ($solution in $solutions) {
        $foundTemplates = $workbookTemplates | Where-Object { $_.properties.packageId -eq $solution.properties.contentId }
        foreach ($foundTemplate in $foundTemplates) {
            $sortedTemplates += $foundTemplate.properties.displayName
        }
    }

    $Table = $doc.Tables.add($selection.Range, $sortedTemplates.count + 1, 2)
    $Table.cell(1, 1).range.Bold = 1
    $Table.cell(1, 1).range.text = "Name"
    $Table.cell(1, 2).range.Bold = 1
    $Table.cell(1, 2).range.text = "Content Source"
    $count = 1

    #Go through all the workbook templates in alphabetical order
    foreach ($singleTemplate in $sortedTemplates | Sort-Object) {
        $count++
        $foundTemplate = $solutionTemplates | Where-Object { $_.properties.displayName -eq $singleTemplate }
        if ($foundTemplate.properties.displayName.count -gt 1) {
            $Table.Cell($count, 1).range.text = $foundTemplate.properties.displayName[0]
        }
        else {
            $Table.Cell($count, 1).range.text = $foundTemplate.properties.displayName
        }
        #See if there is a solution that matches the workbook (there should be)
        $foundSolution = $solutions | Where-Object { $_.properties.displayName -eq $foundTemplate.properties.source.name }
        if ($null -ne $foundSolution) {
            $Table.Cell($count, 2).range.text = $foundSolution.properties.displayName
        }
        else {
            $Table.Cell($count, 2).range.text = "Standalone"
        }
        Write-Host "." -NoNewline
    }
    $selection.endKey(6) | Out-Null
    $selection.TypeParagraph()
    Write-Host ""
}

Function Add-AllSolutions($solutions, $selection) {
    $selection.Style = "Heading 1"
    $selection.TypeText("Installed Solutions")
    $selection.TypeParagraph()
    $selection.TypeText(" ")
    $selection.TypeParagraph()

    Write-Host "Solutions: " -NoNewline
        
    #Just used for testing and to show how many solutions we have worked on
    $count = 1
    #Go through and get each solution alphabetically
    foreach ($solution in $solutions | Sort-Object { $_.properties.displayName }) {
        # if ($count -le 70) {
        #Write-Host $count  $solution.properties.displayName
        Write-Host "." -NoNewline
        Add-SingleSolution $solution $solutionTemplates $selection     #Load one solution
        $selection.TypeText(" ")
        $selection.TypeParagraph()
        #$selection.InsertBreak(7)   #page break
        $count = $count + 1
        <# }
    else {
        break;
    }  #>
    }
    Write-Host ""
}

#Work with a single solutions
Function Add-SingleSolution ($solution, $solutionTemplates, $selection) {

    try {
        #We need to load the solution's template information, which is stored in a separate file.  Note that
        #some solutions have multiple plans associated with them and we will only work with the first one.
        # $uri = $baseUrl + "contentPackages/" + $solution.contentId + $apiVersion
        # $solutionData = (Invoke-RestMethod -Method "Get" -Uri $uri -Headers $authHeader ).value
        #Load the solution's data

        $singleSolutionTemplates = $solutionTemplates | Where-Object { $_.properties.packageId -eq $solution.name }
        #Output the Solution information into the Word Document
        $selection.Style = "Heading 2"
        $selection.TypeText($solution.properties.displayName)
        $selection.TypeParagraph()

        $selection.Font.Bold = $true
        $selection.TypeText("Version: ");
        $selection.Font.Bold = $false
        $selection.TypeText($solution.properties.installedVersion);
        $selection.TypeParagraph()
        #We are using the description rather than the Htmldescription since the Htmldescription can contain HTML formatting that
        #I cannot determine who to translate into what word can understand
        $selection.Font.Bold = $true
        $selection.TypeText("Short Description: ");
        $selection.Font.Bold = $false
        if ($null -eq $solution.properties.description) {
            $selection.TypeText("<None provided>");
        }
        else {
            $selection.TypeText($solution.properties.description);
        }
        $selection.TypeParagraph()
    
        #The hardest part here was determining how each of the various elements were stored in the solutions
        #Load the dataconnectors
        $dataConnectors = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "DataConnector" }
        #Load the workbooks
        $workbooks = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "Workbook" }
        #Load the Analytic Rules
        $ruleTemplates = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "AnalyticsRule" }
        #Load the Hunting Queries
        $huntingQueries = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "HuntingQuery" }
        #Load the Watchlists
        $azureFunctions = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "AzureFunction" }
        #Load the Playbooks
        $playBooks = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "Playbook" }
        #Load the Parsers
        $parsers = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "Parser" }
        #Load the Custom Connectors
        $customConnectors = $singleSolutionTemplates | Where-Object { $_.properties.contentKind -eq "Parser" }

        #Output the summary line to the word document.  I know this is in a lot of solution's descriptions already
        #but it is in HTML and I cannot figure out how to easily translate it to something Word can understand.
        $selection.Font.Bold = $true
        $selection.TypeText("Data Connectors: ");
        $selection.Font.Bold = $false
        $selection.TypeText($dataConnectors.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Workbooks: ");
        $selection.Font.Bold = $false
        $selection.TypeText($workbooks.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Analytic Rules: ");
        $selection.Font.Bold = $false
        $selection.TypeText($ruleTemplates.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Hunting Queries: ");
        $selection.Font.Bold = $false
        $selection.TypeText($huntingQueries.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Azure Functions: ");
        $selection.Font.Bold = $false
        $selection.TypeText($azureFunctions.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Playbooks: ");
        $selection.Font.Bold = $false
        $selection.TypeText($playBooks.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Parsers: ");
        $selection.Font.Bold = $false
        $selection.TypeText($parsers.count);
        $selection.Font.Bold = $true
        $selection.TypeText(", Custom Logic App Connectors: ");
        $selection.Font.Bold = $false
        $selection.TypeText($customConnectors.count);
    
        $selection.TypeParagraph()
    }
    catch {
        $errorReturn = $_
        Write-Error $errorReturn
    }
}



#Execute the code
if (! $Filename.EndsWith(".docx")) {
    $FileName += ".docx"
}
Export-AzSentinelConfigurationToWord  $FileName