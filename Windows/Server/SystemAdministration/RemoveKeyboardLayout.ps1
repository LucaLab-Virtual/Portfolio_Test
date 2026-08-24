<#
.SYNOPSIS
    Rebuilds the Windows user language profile and removes unwanted keyboard layouts.

.DESCRIPTION
    Windows 11 may continue displaying keyboard layouts (such as "English (United States) - US")
    in the language switcher even after they have been removed through Settings.

    This script creates a new language profile for English (United States), clears all existing
    keyboard layouts, adds only the desired layouts, and applies the new configuration.

    In this example, the following keyboard layouts are kept:

        - English (United States) - Latin American Keyboard
        - English (United States) - Spanish Keyboard

    The default US keyboard layout (KBDUS.dll / 00000409) is removed.

.NOTES
    Tested on:
        - Windows 11 Pro

    Requires:
        - Windows PowerShell 5.1 or PowerShell 7+
        - No reboot required (logging out/in may be necessary)

.AUTHOR
    Luis Carlos Ramirez IT Support Specialist / Cloud & Systems Administrator
#>

# Creates a new language list containing only English (United States).
# This replaces the existing language configuration stored in memory.
$LangList = New-WinUserLanguageList "en-US"

# Removes every keyboard layout currently associated with the language.
# This helps eliminate "ghost" or unwanted keyboard layouts that may
# still appear in the Windows language switcher.
$LangList[0].InputMethodTips.Clear()

# Adds the Latin American keyboard layout.
#
# 0409 = English (United States) language
# 0000080A = Latin American keyboard layout
$LangList[0].InputMethodTips.Add("0409:0000080A")

# Adds the Spanish keyboard layout.
#
# 0409 = English (United States) language
# 0000040A = Spanish keyboard layout
$LangList[0].InputMethodTips.Add("0409:0000040A")

# Applies the new language profile.
#
# -Force suppresses confirmation prompts.
# This overwrites the current user language configuration.
Set-WinUserLanguageList $LangList -Force

Write-Host ""
Write-Host "Language profile updated successfully." -ForegroundColor Green
Write-Host "If the unwanted keyboard still appears, sign out and sign back in." -ForegroundColor Yellow
