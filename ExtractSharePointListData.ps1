Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Function ExtractList2Table($ListName,$DBServer,$DBName,$MainTable,$StagingTable,$WebURL)
{
    #Configuration Variables
    $DatabaseServer = $DBServer
    $DatabaseName= $DBName
    function sendmail([string]$Subject, [string]$Body)
    {
        Send-MailMessage -SmtpServer "<smtp address>" -To "<To address>" -From "<From Address>" -Subject $Subject -Body $Body
    }

    #Get Web, List and Fields
    $Web= Get-SPWeb $WebURL
    $List= $Web.Lists[$ListName]
    #Get all required fields from the lists
    $ListFields = $List.Fields | 
                                Where-Object{ ($_.Hidden -ne $true ) -and
                                   ($_.InternalName -ne "Attachments") -and
                                   ($_.InternalName -ne "ContentType") -and
                                   ($_.ReadOnlyField -ne $true -or $_.InternalName -eq 'ID' -or $_.InternalName -eq 'Modified' -or 
                                    $_.InternalName -eq 'Created' -or $_.InternalName -eq 'Author' -or $_.InternalName -eq 'Editor')
                                }

    #Get SQL column Definition for SharePoint List Field
    Function Get-ColumnDefinition([Microsoft.SharePoint.SPField]$Field){
        $ColumnDefinition=""

        Switch($Field.Type)
        {
            "Boolean" { $ColumnDefinition = '['+ $Field.Title +'] [bit] NULL '}
            "Choice" { $ColumnDefinition = '['+ $Field.Title +'] [nvarchar](MAX) NULL '}
            "Currency" { $ColumnDefinition = '['+ $Field.Title +'] [decimal](18, 2) NULL '}
            "DateTime" { $ColumnDefinition = '['+ $Field.Title +'] [datetime] NULL '}
            "Guid" { $ColumnDefinition = '['+ $Field.Title +'] [uniqueidentifier] NULL '}
            "Integer" { $ColumnDefinition = '['+ $Field.Title +'] [int] NULL '}
            "Lookup" { $ColumnDefinition = '['+ $Field.Title +'] [nvarchar] (500) NULL '}
            "MultiChoice" { $ColumnDefinition = '['+ $Field.Title +'] [nText] (MAX) NULL '}
            "Note" { $ColumnDefinition = '['+ $Field.Title +'] [nText] NULL '}
            "Number" { $ColumnDefinition = '['+ $Field.Title +'] [decimal](18, 2) NULL '}
            "Text" { $ColumnDefinition = '['+ $Field.Title +'] [nVarchar] (MAX) NULL '}
            "URL" { $ColumnDefinition = '['+ $Field.Title +'] [nvarchar] (500) NULL '}
            "User" { $ColumnDefinition = '['+ $Field.Title +'] [nvarchar] (255) NULL '}
            default { $ColumnDefinition = '['+ $Field.Title +'] [nvarchar] (MAX) NULL '}
        }
        return $ColumnDefinition
        }
    ################ Format Column Value Functions ######################
    Function Format-UserValue([object] $ValueToFormat)
    {
        if([String]::IsNullOrEmpty($ValueToFormat) -eq $false) {
            $Users = $ValueToFormat.Substring($ValueToFormat.IndexOf("#") + 1)
            return "'" + $Users + "'"
        } else {
            write-host $ValueToFormatreturn "'NULL'";
        }
    }
    Function Format-LookupValue([Microsoft.SharePoint.SPFieldLookupValueCollection] $ValueToFormat){
        $LookupValue = [string]::join("; ",( $ValueToFormat | Select-Object -expandproperty LookupValue))
        $LookupValue = $LookupValue -replace "'", "''"
        return "'" + $LookupValue + "'"
    }
    Function Format-DateValue([string]$ValueToFormat){
        [datetime] $dt = $ValueToFormat
        return "'" + $dt + "'"
    }
    Function Format-CurrencyValue([string]$ValueToFormat){
        [decimal] $dc = $ValueToFormat
        return "'" + $dc + "'"
    }
    Function Format-MMSValue([Object]$ValueToFormat){
        return "'" + $ValueToFormat.Label + "'"
    }
    Function Format-BooleanValue([string]$ValueToFormat){
        if($ValueToFormat -eq "Yes") {return 1} else { return 0}
    }
    Function Format-StringValue([object]$ValueToFormat)
    {
        [string]$result = $ValueToFormat -replace "'", "''"
        return "'" + $result + "'"
    }
    #Function to get the value of given field of the List item
    Function Get-ColumnValue([Microsoft.SharePoint.SPListItem] $ListItem, [Microsoft.SharePoint.SPField]$Field){
        $FieldValue= $ListItem[$Field.Title]
        #Check for NULL
        if([string]::IsNullOrEmpty($FieldValue)) { return 'NULL'}
        $FormattedValue = ""

        Switch($Field.Type)
        {
        "Boolean"  {$FormattedValue =  Format-BooleanValue($FieldValue)}
        "Choice"  {$FormattedValue = Format-StringValue($FieldValue)}
        "Currency"  {$FormattedValue = Format-CurrencyValue($FieldValue) }
        "DateTime"  {$FormattedValue = Format-DateValue($FieldValue)}
        "Guid" { $FormattedValue = Format-StringValue($FieldValue)}
        "Integer"  {$FormattedValue = $FieldValue}"Lookup"  {$FormattedValue = Format-LookupValue($FieldValue) }
        "MultiChoice" {$FormattedValue = Format-StringValue($FieldValue)}
        "Note"  {$FormattedValue = Format-StringValue($Field.GetFieldValueAsText($ListItem[$Field.Title]))}
        "Number"    {$FormattedValue = $FieldValue}"Text"  {$FormattedValue = Format-StringValue($Field.GetFieldValueAsText($ListItem[$Field.Title]))}
        "URL"  {$FormattedValue =  Format-StringValue($FieldValue)}
        "User"  {$FormattedValue = Format-UserValue($FieldValue) }
        #Check MMS Field
        "Invalid" { if($Field.TypeDisplayName -eq "Managed Metadata") { $FormattedValue = Format-MMSValue($FieldValue) } else { $FormattedValue =Format-StringValue($FieldValue)}  }
        default  {$FormattedValue = Format-StringValue($FieldValue)}
    }
    Return $FormattedValue
    }

    #Create SQL Server table for SharePoint List
    Function CreateMainTable([Microsoft.SharePoint.SPList]$List)
    {
        #Check if the table exists already
        $TableCheckQuery = "Select OBJECT_ID('[dbo].[$($MainTable)]','U')"
        $Result = Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query $TableCheckQuery -querytimeout 300

        if([String]::IsNullOrEmpty($Result.Column1.ToString()))
        {
            Write-Host "Creating table"
            #Create the table
            $Query="CREATE TABLE [dbo].[$($MainTable)]("
            foreach ($Field in $ListFields)
            {
                $Query += Get-ColumnDefinition($Field)
                $Query += ","
            }
            $Query += ")"

            #Run the Query
            Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query $Query -querytimeout 300
        }
    }
    Function CreateStagingTable([Microsoft.SharePoint.SPList]$List)
    {
        #Check if the table exists already
        $TableCheckQuery = "Select OBJECT_ID('[dbo].[$($StagingTable)]','U')"
        $Result = Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query $TableCheckQuery -querytimeout 300

        if([String]::IsNullOrEmpty($Result.Column1.ToString()))
        {
            Write-Host "Creating table"
            #Create the table
            $Query="CREATE TABLE [dbo].[$($StagingTable)]("
            foreach ($Field in $ListFields)
            {
                $Query += Get-ColumnDefinition($Field)
                $Query += ","
            }
            $Query += ")"

            #Run the Query
            Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query $Query -querytimeout 300
        }
      
    }

    #Insert Data from SharePoint List to SQL Table
    Function InsertData([Microsoft.SharePoint.SPList]$List)
    {
        #clear Staging table
        Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query "Delete from [dbo].[$($StagingTable)]" -querytimeout 300
        $endDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        #Get the last timestamp when data was collected from the SharePoint List
        $MaxDate = Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query "select max(Modified) FROM [dbo].[$($MainTable)]"
        if([String]::IsNullOrEmpty($MaxDate.Column1.ToString()))
        {
            $MaxDate = "2010-10-01T00:00:00Z"
        }
        else
        {
            $MaxDate = $MaxDate.Column1.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        #Run caml query to select the items from SharePoint since the last run
        Write-host "Executing caml"
        $spQuery = New-Object Microsoft.SharePoint.SPQuery
        $spQuery.ViewAttributes = "Scope='Recursive'";
        $caml =
        '<Where>
            <And>
            <And>
            <Gt>
                <FieldRef Name="Modified" />
                <Value Type="DateTime" IncludeTimeValue="True">'+$MaxDate+'</Value>
            </Gt>
            <Leq>
                <FieldRef Name="Modified" />
                <Value Type="DateTime" IncludeTimeValue="True">'+$endDate+'</Value>
            </Leq>
            </And>
            <Eq>
                <FieldRef Name="FSObjType" />
                <Value Type="int">0</Value>
            </Eq>
            </And>
        </Where>'
        $spQuery.Query = $caml

        $ListItems=$List.GetItems($spQuery)

        #Progress bar counter
        $Counter=0
        $ListItemCount=$ListItems.Count

        Write-Host "Executing Inserting"
        Write-host "Total SharePoint List Items to Copy:" $ListItemCount

        foreach ($Item in $ListItems)
        {
            Write-Progress -Activity "Copying SharePoint List Items. Please wait...`n`n" -status "Processing List Item: $($Item['ID'])" -percentComplete ($Counter/$ListItemCount*100)

            $sql = new-object System.Text.StringBuilder
            [void]$sql.Append("INSERT INTO [dbo].[$($StagingTable)] (")
            $vals = new-object System.Text.StringBuilder
            [void]$vals.Append("VALUES (")

            $loop = 0
            foreach ($Field in $ListFields)
            {
                if($loop -gt 0)
                {
                    [void]$sql.Append(",")
                    [void]$vals.Append(",")
                }
                [void]$sql.Append("[$($Field.Title)]")
                $ColumnValue =  Get-ColumnValue $Item $Field
                [void]$vals.Append($ColumnValue)

                $loop += 1
            }

        [void]$sql.Append(") ")
        [void]$vals.Append(") ")

        #Combine Field and Values
        $SQLStatement = $sql.ToString() + $vals.ToString()

        #Run the Query$SQLStatement
        try {
            Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query $SQLStatement -querytimeout 300
            }
        catch
            {
                $body =  $QueryStr + "`n" + $error[0].Exception
                sendmail "Insert - LoadSPData: Load from SharePoint Error" $body
            }
        $Counter += 1;
        }
        "Total SharePoint List Items Copied: $($ListItemCount)"  
         if($ListItemCount -gt 0) {      
            MergeData
         }
    }

    #Merge Data from taging table
    Function MergeData(){
        $data = Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query "Select * from dbo.[$($StagingTable)] nolock" -querytimeout 300
        $TableColumns = $data | Get-Member | Where-Object {$_.membertype -eq 'property'} | Select-Object name

            $sql = new-object System.Text.StringBuilder
            [void]$sql.Append("MERGE [dbo].[$($MainTable)]  AS Main ")
            [void]$sql.Append("Using [dbo].[$($StagingTable)] AS Staging ")
            [void]$sql.Append("ON (Main.[ID] = Staging.[ID]) ")
            [void]$sql.Append("WHEN MATCHED THEN UPDATE SET ")

            $loop = 0
            foreach ($column in $TableColumns.name)
            {
                if($loop -gt 0)
                {
                    [void]$sql.Append(",")
                }
                [void]$sql.Append("Main.[$($column)] = Staging.[$($column)]")
                $loop += 1
            }

        [void]$sql.Append(" WHEN NOT MATCHED THEN ")
        [void]$sql.Append("INSERT( ")
            $loop = 0
            foreach ($column in $TableColumns.name)
            {
                if($loop -gt 0)
                {
                    [void]$sql.Append(",")
                }
                [void]$sql.Append("[$($column)]")
                $loop += 1
            }
        [void]$sql.Append(") ")
        [void]$sql.Append(" VALUES( ")
        $loop = 0
        foreach ($column in $TableColumns.name)
            {
                if($loop -gt 0)
                {
                    [void]$sql.Append(",")
                }
                [void]$sql.Append("Staging.[$($column)]")
                $loop += 1
            }
        [void]$sql.Append("); ")

        #Combine Field and Values
        $SQLStatement = $sql.ToString()

        #Run the Query# $SQLStatement

        try {
            Invoke-Sqlcmd -ServerInstance $DatabaseServer -Database $DatabaseName -Query $SQLStatement -querytimeout 300
            }
        catch
            {
                $body =  $QueryStr + "`n" + $error[0].Exception
                sendmail "Insert - LoadSPData: Load from SharePoint Error" $body
            }
    }

        #Call functions to export-import SharePoint list to SQL table
        <#  Drop-Table $MainTable #>

        CreateMainTable $List
        CreateStagingTable $List
        InsertData $List
}       

#Call the function to Create SQL Server Table from SharePoint List
ExtractList2Table -ListName "<ListName>" -DBServer "<DB Server>" -DBName "<DB Name>" -MainTable "<Main Table>" -StagingTable "<Staging Table>" -WebURL "<SP Web Url>"
