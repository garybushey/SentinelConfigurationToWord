# SentinelConfigurationToWord
This command will generate a Word document containing the information about all the Azure Sentinel
        Solutions.  

    PARAMETERS
    
    WorkspaceName
        Enter the workspace name to use. 
        
    ResourceGroupName
        Enter the Resource Group name to use.  
    
     FileName
        Enter the file name to use.  Defaults to "MicrosoftSentinelSolutions.docx"  ".docx" will be appended to all filenames if needed
    
    
    EXAMPLES
        Export-AzSentinelConfigurationToWork -WorkspaceName testwg -ResourceGroupName rgName
        In this example you will get the file named "MicrosoftSentinelReport.docx" generated containing all configuration for the Sentinel instance

        Export-AzSentinelSolutionsToWord -WorkspaceName testwg -ResourceGroupName rgName -fileName "test"
        In this example you will get the file named "test.docx" generated containing all configuration for the Sentinel instance
