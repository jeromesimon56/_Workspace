param(
    [string]$Action = "all",
    [string]$Folder = "Inbox",
    [string]$EmailsOutput = "./outlook_emails",
    [string]$AttachmentsOutput = "./attachments",
    [string]$Rules = "./config/settings.json",
    [int]$MaxMessages = 0,
    [string[]]$ForceDeleteDomains = @(),
    [string]$ImapUsername = "",
    [string]$ImapPassword = "",
    [string]$ImapAuthMethod = "auto",
    [string]$ImapOAuth2Token = "",
    [switch]$StandardCleanup,
    [switch]$DryRun,
    [switch]$AutoConfirm
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

switch ($Action.ToLower()) {
    "fetch" {
        & (Join-Path $scriptRoot "scripts\outlook_fetch.ps1") -Folder $Folder -Output $EmailsOutput -MaxMessages $MaxMessages
    }
    "extract" {
        & (Join-Path $scriptRoot "scripts\extract_attachments.ps1") -Folder $Folder -Output $AttachmentsOutput -MaxMessages $MaxMessages
    }
    "imap-extract" {
        if (-not $ImapUsername) {
            Write-Error "Action 'imap-extract' requires -ImapUsername"
            exit 1
        }

        if ($ImapAuthMethod -ne 'oauth2' -and -not $ImapPassword) {
            Write-Error "Action 'imap-extract' requires -ImapPassword pour AuthMethod '$ImapAuthMethod'"
            exit 1
        }

        & (Join-Path $scriptRoot "scripts\imap_extract_attachments.ps1") -Username $ImapUsername -Password $ImapPassword -AuthMethod $ImapAuthMethod -OAuth2Token $ImapOAuth2Token -Output $AttachmentsOutput
    }
    "clean" {
        & (Join-Path $scriptRoot "scripts\clean_promotions.ps1") -Folder $Folder -MaxMessages $MaxMessages -ForceDeleteDomains $ForceDeleteDomains @([System.Management.Automation.SwitchParameter]::new($StandardCleanup)) @([System.Management.Automation.SwitchParameter]::new($DryRun)) @([System.Management.Automation.SwitchParameter]::new($AutoConfirm))
    }
    "sort" {
        . (Join-Path $scriptRoot "scripts\outlook_utils.ps1")
        Sort-OutlookEmails -Source $EmailsOutput -RulesFile $Rules
    }
    "all" {
        & (Join-Path $scriptRoot "scripts\outlook_fetch.ps1") -Folder $Folder -Output $EmailsOutput -MaxMessages $MaxMessages
        & (Join-Path $scriptRoot "scripts\extract_attachments.ps1") -Folder $Folder -Output $AttachmentsOutput -MaxMessages $MaxMessages
        . (Join-Path $scriptRoot "scripts\outlook_utils.ps1")
        Sort-OutlookEmails -Source $EmailsOutput -RulesFile $Rules
    }
    default {
        Write-Host "Unknown action: $Action. Use fetch, extract, imap-extract, clean, sort, or all."
    }
}
