#Requires -Version 5.1
<#
.SYNOPSIS
    pwshMimecast - PowerShell module collection for Mimecast tenant management.

.DESCRIPTION
    Root module that loads all Mimecast sub-modules. Import this module once to
    get access to all pwshMimecast commands.

    Sub-modules included:
      Mimecast-Delegates.psm1  - Delegate mailbox management (API 2.0 and API 1.0)

    To add a new sub-module, create the .psm1 in this folder and add a
    dot-source line below in the Sub-module Loading region.

.NOTES
    Author  : Bill Kindle (with AI assistance)
    Version : 1.0
    Created : 2026-04-02
#>

# Sub-modules are loaded via NestedModules in pwshMimecast.psd1.
# Do not dot-source sub-modules here; dot-sourcing a .psm1 that contains
# module-level comment-based help can silently prevent function definitions
# in PowerShell 5.1 and 7.x. NestedModules is the supported pattern.

