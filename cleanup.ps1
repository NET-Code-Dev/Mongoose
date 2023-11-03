# The path to the mongoose directory
$mongooseDirectory = [System.IO.Path]::Combine($env:APPDATA, "Mongoose")

# Delete the directory if it exists
if (Test-Path $mongooseDirectory) {
    Remove-Item $mongooseDirectory -Force -Recurse
}