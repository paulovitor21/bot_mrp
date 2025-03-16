$exclude = @("venv", "bot_mrp.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "bot_mrp.zip" -Force