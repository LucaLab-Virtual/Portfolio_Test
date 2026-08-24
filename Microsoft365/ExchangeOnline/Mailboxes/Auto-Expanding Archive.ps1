# Enabling Auto-Expanding Archive for a user

# Enable Auto-Expanding Archive
Enable-Mailbox sample@sample.com -AutoexpandingArchive
# Confirm it's enabled
Get-Mailbox sample@sample.com | fl AutoexpandingArchiveEnabled
# End