Add-Type -AssemblyName System.Windows.Forms
if(-not (Get-Module -ListAvailable -Name PSWritePDF)){
    [System.Windows.Forms.MessageBox]::Show("Dependency Not Found, Please Install PSWritePDF...")
    Exit
}
$Functionnality = Read-Host "Please select functionnality: [1] Batch convert PowerPoint to PDF [2] Batch convert Word to PDF [3] Merge PDF"
switch ($Functionnality) {
    1 {
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            Title = "Select PowerPoint files to convert"
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'PowerPoint Documents (*.pptx)|*.pptx'
            Multiselect = $true
        }
        $FileBrowser.ShowDialog()
        $ppt = New-Object -com powerpoint.application
        $opt = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
        $FileBrowser.FileNames | ForEach-Object {
            $ifile = $_
            $pres = $null
            try {
                $pres = $ppt.Presentations.Open($ifile)
                $pathname = split-path $ifile
                $filename = Split-Path $ifile -Leaf
                $file = $filename.split(".")[0]
                $ofile = $pathname + "\" + $file + ".pdf"
                $pres.SaveAs($ofile, $opt)
            }
            finally {$pres.Close()}
        }
        [System.Windows.Forms.MessageBox]::Show("PowerPoint to PDF conversion done")
    }       
    2 {
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            Title = "Select Word files to convert"
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'Word Documents (*.docx)|*.docx'
            Multiselect = $true
        }
        $FileBrowser.ShowDialog()
        $word = New-Object -ComObject Word.Application
        $opt = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
        $FileBrowser.FileNames | ForEach-Object {
            $ifile = $_
            $doc = $null
            try {
                $doc = $word.Documents.Open($ifile)
                $pathname = split-path $ifile
                $filename = Split-Path $ifile -Leaf
                $file = $filename.split(".")[0]
                $ofile = $pathname + "\" + $file + ".pdf"
                $doc.SaveAs($ofile, $opt)
            } 
            finally {$doc.Close()}
        }
        [System.Windows.Forms.MessageBox]::Show("Word to PDF conversion done")

    }
    3 {
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            Title = "Select PDF documents to merge"
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'PDF Documents (*.pdf)|*.pdf'
            Multiselect = $true
        }
        $FileBrowser.ShowDialog()
        $SaveFile = New-Object System.Windows.Forms.SaveFileDialog -Property @{ 
            Title = "Save merged PDF document"
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'PDF Documents (*.pdf)|*.pdf'
        }
        $SaveFile.ShowDialog()
        try {
            Merge-PDF -InputFile $FileBrowser.FileNames $SaveFile.FileName
            [System.Windows.Forms.MessageBox]::Show("PDF merge done")
        }
        catch [System.Exception] {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message)
        }
    }
    Default { 
        [System.Windows.Forms.MessageBox]::Show("Invalid option")
        Exit
    }
}


