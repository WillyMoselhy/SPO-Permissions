# Input bindings are passed in via param block.
param($QueueItem, $TriggerMetadata)

# Write out the queue message and insertion time to the information log.
Write-PSFMessage "PowerShell queue trigger function processed work item: $QueueItem"
Write-PSFMessage "Queue item insertion time: $($TriggerMetadata.InsertionTime)"

$permissionsScanMeasure = Measure-Command -Expression {
    .\ScanSiteCollection.ps1 -SiteCollectionURL $QueueItem
}
Write-PSFMessage -Level Host -Message "Finished permissions scanning for URL: $QueueItem - Time (seconds): $($permissionsScanMeasure.TotalSeconds)"