# Fix DOCX Files
# Created by: Adem Kouki
# https://github.com/Ademking/docx-fixer/
# ---------------
# How to run:
# ./fix.ps1 "C:\folder\corrupted-document.docx" "C:\folder\fixed-document.docx"
# ---------------

Try {
        $Doc = $args[0]
        $word = New-Object -ComObject word.application
        $word.Visible = $false

        #                            1      2       3       4     5   6    7     8   9 10   11     12     13   14    15        16
        [void]$word.documents.Open($Doc, $false, $false, $false, "", "", $true, "", "", 0, $null, $true, $true, 0, $false) #, $null)
        
        <# Open Parameters:
        Document Open(
	         1 string FileName, 
	         2 bool ConfirmConversions,
	         3 bool ReadOnly,
	         4 bool AddToRecentFiles,
	         5 string PasswordDocument,
	         6 string PasswordTemplate,
	         7 bool Revert,
 	         8 string WritePasswordDocument,
	         9 string WritePasswordTemplate,
	        10 int Format,
	        11 Object Encoding,
	        12 bool Visible,
	        13 bool OpenAndRepair, <- this
	        14 int DocumentDirection,
	        15 bool NoEncodingDialog,
	        16 Object XMLTransform
        )#>

        $outputFile = $args[1]
        $saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
        $word.ActiveDocument.SaveAs([ref][system.object]$outputFile, [ref]$saveFormat)
        $word.ActiveDocument.Close();
        $word.Quit()

        # Clean up Com object
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
        Remove-Variable word
}

Catch{
  $logpath = Join-Path -Path $pwd -ChildPath "logs.txt"
  $_ | Out-File $logpath
}
