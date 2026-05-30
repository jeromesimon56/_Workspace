function Sort-OutlookEmails {
    param(
        [string]$Source,
        [string]$RulesFile
    )

    $rules = Get-Content $RulesFile | ConvertFrom-Json

    foreach ($rule in $rules.rules) {
        $keyword = $rule.keyword
        $target = $rule.target

        Get-ChildItem $Source -Filter *.eml | Where-Object {
            $_.Name -match $keyword
        } | Move-Item -Destination $target -Force
    }
}
