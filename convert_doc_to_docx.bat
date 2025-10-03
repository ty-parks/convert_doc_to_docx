@echo off
setlocal

:: Create a temporary PowerShell script
set "psScript=%TEMP%\convert_doc_to_docx.ps1"
> "%psScript%" echo $word = New-Object -ComObject Word.Application
>> "%psScript%" echo $word.Visible = $false
>> "%psScript%" echo $folder = Get-Location
>> "%psScript%" echo $files = Get-ChildItem -Path $folder -Filter *.doc
>> "%psScript%" echo foreach ($file in $files) {
>> "%psScript%" echo     $docPath = $file.FullName
>> "%psScript%" echo     $newName = [System.IO.Path]::ChangeExtension($docPath, ".docx")
>> "%psScript%" echo     $retry = 0
>> "%psScript%" echo     do {
>> "%psScript%" echo         try {
>> "%psScript%" echo             $doc = $word.Documents.Open($docPath, [ref] $false, [ref] $true)
>> "%psScript%" echo             Start-Sleep -Milliseconds 500
>> "%psScript%" echo             $doc.SaveAs([ref] $newName, [ref] 12)
>> "%psScript%" echo             $doc.Close()
>> "%psScript%" echo             $retry = 5
>> "%psScript%" echo         } catch {
>> "%psScript%" echo             Start-Sleep -Seconds 1
>> "%psScript%" echo             $retry++
>> "%psScript%" echo         }
>> "%psScript%" echo     } while ($retry -lt 5)
>> "%psScript%" echo }
>> "%psScript%" echo $word.Quit()

:: Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "%psScript%"

:: Clean up
del "%psScript%"
endlocal
