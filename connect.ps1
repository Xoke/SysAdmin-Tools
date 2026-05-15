# Connect to Exchange Online and Microsoft Graph
# Requires: ExchangeOnlineManagement v3+, Microsoft.Graph

# Exchange Online
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline
}

# Microsoft Graph (replaces MSOnline / Connect-MsolService)
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Organization.Read.All"
}
