@{
    # v2.0.0: lint everything, but do not fail the build on the two
    # pre-existing style choices that are intentional in this codebase:
    # - Write-Host is used for interactive operator output in the
    #   interactive wizards (Sort-Inbox, Export-SharedCalendar browse).
    # - Unused variables appear in the Graph diagnostics helpers, where
    #   results are collected for troubleshooting output.
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'
    )
    Severity = @('Error', 'Warning')
}
