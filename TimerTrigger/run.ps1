# Input bindings are passed in via param block.
param($Timer)

# Write to the Azure Functions log stream.
Write-Host "Starting scheduled Teams to SharePoint synchronization..."

Invoke-Synchronization
