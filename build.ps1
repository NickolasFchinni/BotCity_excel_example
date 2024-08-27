$exclude = @("venv", "excel_example.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "excel_example.zip" -Force