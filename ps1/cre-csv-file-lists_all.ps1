Get-ChildItem $args[0] -Recurse|
ConvertTo-Csv|
Set-Content $args[1]
