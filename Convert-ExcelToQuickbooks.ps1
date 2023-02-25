# Set configuration variables
$SourceFolder = "$env:userprofile\Downloads"
$DestinationFolder = "$env:userprofile\Downloads"
Write-Host "Source folder: $SourceFolder"
Write-Host "Destination folder: $DestinationFolder"

# Get last xlsx file in the source folder
Write-Host "Looking for latest *.xlxs file in the source folder"
$SourceFile = Get-ChildItem -Path $SourceFolder -Filter *.xlsx | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name
if (!$SourceFile) {
    Write-Host "No *.xlxs files found in $SourceFolder"
    Read-Host "Press 'Enter' to exit"
    Return
}
$DestinationFile = $SourceFile -replace "xlsx", "iif"

# Create value mapping
$ClassMapping = @{
    "10937"   = "WARWICK"
    "10938"   = "NORTH PROVIDENCE"
    "10939"   = "NARRAGANSETT"
    "10941"   = "WOONSOCKET"
    "10942"   = "WESTERLY"
    "10949"   = "SOUTH DENNIS"
    "10950"   = "HYANNIS"
    "10951"   = "BROCKTON"
    "10952"   = "CHELMSFORD"
    "10953"   = "MANCHESTER"
    "10954"   = "NORWAY"
    "10959"   = "TRUMBULL"
    "10964"   = "DAYVILLE"
    "35098"   = "SCARBOROUGH"
    "35970"   = "WAREHAM"
    "35971"   = "PLYMOUTH"
    "36045"   = "SOUTH ATTLEBORO"
    "36532"   = "WILLIMANTIC"
    "G127089" = "CRANSTON"
    "G127120" = "ORANGE"
    "G127353" = "BRISTOL"
    "G127354" = "BROOKFIELD"
    "G127355" = "LEWISTON"
    "G128021" = "SALEM"
    "G128917" = "NORTH DARTMOUTH"
    "G128926" = "BANGOR"
    "G129165" = "ROCKPORT"
    "G129544" = "BARRINGTON"
    "G129607" = "SOUTH BURLINGTON"
    "G131356" = "WATERVILLE"
    "G131390" = "MIDDLETOWN CT"
    "G131969" = "LEXINGTON"
    "G132381" = "NASHUA"
    "G132427" = "NORWALK"
    "G132494" = "WEST SPRINGFIELD"
    "G133010" = "BRUNSWICK"
    "G150030" = "AUBURN"
    "G150691" = "PRESQUE ISLE"
    "G151095" = "SAUGUS"
    "G151625" = "AUGUSTA"
    "G161417" = "GREENVILLE"
    "G152464" = "MONTPELIER"
    "G152465" = "MORRISVILLE"
    "G152779" = "NEWCASTLE"
    "G152862" = "MIDDLETON"
    "G153019" = "WARWICK"
    "G153051" = "GROTON"
    "G153052" = "NORWICH"
    "G153547" = "FARMINGTON"
    "G155196" = "WEST HARTFORD"
    "G156513" = "LEOMINSTER"
    "G129656" = "MOORESVILLE"
    "G129709" = "WILMINGTON"
    "G129813" = "ROCK HILL"
    "G129814" = "JACKSONVILLE"
    "G130128" = "ABERDEEN"
    "G130499" = "GLEN ALLEN"
    "G130500" = "RICHMOND"
    "G130501" = "COLONIAL HEIGHTS"
    "G130502" = "MECHANICSVILLE"
    "G131572" = "CHARLOTTE"
    "G131573" = "MATTHEWS"
    "G132592" = "INDIAN LAND"
    "G134416" = "MORGANTON"
    "G150472" = "LANCASTER"
    "G151344" = "LENOIR"
    "G153761" = "SPARTANBURG"
    "G152629" = "HICKORY"
    "G161374" = "WATERBURY"
    "G162113" = "PORTLAND"
    "G162111" = "NEWBURYPORT"
    "G162112" = "BRADFORD"
}

$NameMapping = @{
    "G156888" = @{
        Name  = "Juan MANAGED CARE"
        Terms = "Net 90"
    }
    "G156887" = @{
        Name  = "Juan MANAGED CARE"
        Terms = "Net 90"
    }
    "10937"   = @{
        Name  = "Juan"
        Terms = "Net 60"
    }
    "G129656" = @{
        Name  = "Juan"
        Terms = "Net 30"
    }
}

# Check or install ImportExcel module
# https://www.powershellgallery.com/packages/ImportExcel/
if ($null -eq (Get-Module -ListAvailable ImportExcel)) {
    try {
        Write-Output "ImportExcel module not found. Trying to install"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
        Install-Module ImportExcel -Force -ErrorAction Stop | Out-Null
        Write-Output "ImportExcel module was installed"
    }
    catch {
        Write-Output "Unable to install ImportExcel module. Try running the script 'As Administrator' or install module manually. $($Error[0].Exception.Message)"
        Read-Host "Press 'Enter' to exit"
        Return
    }
}

# Import Excel file
Write-Host "Importing file: $SourceFolder\$SourceFile"
try {
    $ExcelData = Import-Excel -Path $SourceFolder\$SourceFile
    $AllDocuments = $ExcelData | Group-Object -Property DocumentNo
}
catch {
    # Catch errors
    Write-Output "Faile to import $SourceFile file. $($Error[0].Exception.Message)"
    Read-Host "Press 'Enter' to exit"
    Return
}

try {
    # Create new empty array for final results
    $FinalDocument = @()

    # Add headers to the final result
    $TRNSHeader = [PSCustomObject]@{
        TRNS_SPL            = "!TRNS"
        TRNSID_SPLID        = "TRNSID"
        TRNSTYPE            = "TRNSTYPE"
        DATE                = "DATE"
        ACCNT               = "ACCNT"
        NAME                = "NAME"
        CLASS               = "CLASS"
        AMOUNT              = "AMOUNT"
        DOCNUM              = "DOCNUM"
        MEMO                = "MEMO"
        CLEAR               = "CLEAR"
        TOPRINT_QNTY        = "TOPRINT"
        ADDR5_REIMBEXP      = "ADDR5"
        DUEDATE_SERVICEDATE = "DUEDATE"
        TERMS_OTHER2        = "TERMS"
    }

    $SPLHeader = [PSCustomObject]@{
        TRNS_SPL            = "!SPL"
        TRNSID_SPLID        = "SPLID"
        TRNSTYPE            = "TRNSTYPE"
        DATE                = "DATE"
        ACCNT               = "ACCNT"
        NAME                = "NAME"
        CLASS               = "CLASS"
        AMOUNT              = "AMOUNT"
        DOCNUM              = "DOCNUM"
        MEMO                = "MEMO"
        CLEAR               = "CLEAR"
        TOPRINT_QNTY        = "QNTY"
        ADDR5_REIMBEXP      = "REIMBEXP"
        DUEDATE_SERVICEDATE = "SERVICEDATE"
        TERMS_OTHER2        = "OTHER2"
    }

    $TRNSENDHeader = [PSCustomObject]@{
        TRNS_SPL            = "!ENDTRNS"
        TRNSID_SPLID        = ""
        TRNSTYPE            = ""
        DATE                = ""
        ACCNT               = ""
        NAME                = ""
        CLASS               = ""
        AMOUNT              = ""
        DOCNUM              = ""
        MEMO                = ""
        CLEAR               = ""
        TOPRINT_QNTY        = ""
        ADDR5_REIMBEXP      = ""
        DUEDATE_SERVICEDATE = ""
        TERMS_OTHER2        = ""
    }

    $TRNSEND = [PSCustomObject]@{
        TRNS_SPL            = "ENDTRNS"
        TRNSID_SPLID        = ""
        TRNSTYPE            = ""
        DATE                = ""
        ACCNT               = ""
        NAME                = ""
        CLASS               = ""
        AMOUNT              = ""
        DOCNUM              = ""
        MEMO                = ""
        CLEAR               = ""
        TOPRINT_QNTY        = ""
        ADDR5_REIMBEXP      = ""
        DUEDATE_SERVICEDATE = ""
        TERMS_OTHER2        = ""
    }

    $FinalDocument += $TRNSHeader
    $FinalDocument += $SPLHeader
    $FinalDocument += $TRNSENDHeader

    # Add transaction to the final result
    ForEach ($SingleDocument in $AllDocuments) {
        $AllRecordsInTheDocument = $SingleDocument.Group
        $FirstRecordInTheDocument = $AllRecordsInTheDocument | Select-Object -First 1
        
        $TRNSRecord = [PSCustomObject]@{
            TRNS_SPL            = "TRNS"
            TRNSID_SPLID        = ""
            TRNSTYPE            = if ($FirstRecordInTheDocument.DocumentType -eq 'Invoice') { "BILL" } elseif ($FirstRecordInTheDocument.DocumentType -eq 'Credit Memo') { "CREDIT MEMO" } else { "" }
            DATE                = ([Datetime]$FirstRecordInTheDocument.OrderDate).toString("M\/d\/yyyy")
            ACCNT               = "2000"
            NAME                = $NameMapping.$($FirstRecordInTheDocument.BilltoCustomerNo).Name
            CLASS               = $ClassMapping.$($FirstRecordInTheDocument.SelltoCustomerNo)
            AMOUNT              = "-$($FirstRecordInTheDocument.TotalDocumentAmount)"
            DOCNUM              = $SingleDocument.Name
            MEMO                = ""
            CLEAR               = "N"
            TOPRINT_QNTY        = "N"
            ADDR5_REIMBEXP      = ""
            DUEDATE_SERVICEDATE = ([Datetime]$FirstRecordInTheDocument.DueDate).toString("M\/d\/yyyy")
            TERMS_OTHER2        = $NameMapping.$($FirstRecordInTheDocument.BilltoCustomerNo).Terms
        }
        $FinalDocument += $TRNSRecord

        $AccountGroups = @{
            "4010" = @()
            "4040" = @()
            "5245" = @()
            "6040" = @()
        }
        ForEach ($SingleRecord in $AllRecordsInTheDocument) {

            if ($SingleRecord.ProductType -in @("BTE Hearing Aid", "ITE Hearing Aid")) {
                $AccountGroups.'4010' += $SingleRecord
            }
            elseif ($SingleRecord.ProductType -in @("Chargers", "Parts & Accessories", "Wireless Accessories")) {
                $AccountGroups.'4040' += $SingleRecord
            }
            elseif ($SingleRecord.ItemNo -in @("BEC MARKETING FUND")) {
                $AccountGroups.'5245' += $SingleRecord
            }
            elseif ($SingleRecord.ItemDescription -in @("Shipping Fees", "Shipping Charge")) {
                $AccountGroups.'6040' += $SingleRecord
            }
        }

        if ($AccountGroups.'4010') {

            $AccountValues = $AccountGroups.'4010'
            $SPLRecord = [PSCustomObject]@{
                TRNS_SPL            = "SPL"
                TRNSID_SPLID        = ""
                TRNSTYPE            = if ($AccountValues[0].DocumentType -eq 'Invoice') { "BILL" } elseif ($AccountValues[0].DocumentType -eq 'Credit Memo') { "CREDIT MEMO" } else { "" }
                DATE                = ([Datetime]$AccountValues[0].OrderDate).toString("M\/d\/yyyy")
                ACCNT               = "4010"
                NAME                = ""
                CLASS               = $ClassMapping.$($AccountValues[0].SelltoCustomerNo)
                AMOUNT              = $AccountValues | ForEach-Object { [decimal]$_.UnitPrice * [int]$_.Quantity } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                DOCNUM              = $AccountValues[0].DocumentNo
                MEMO                = "`"$($AccountValues[0].PatientLastName), $($AccountValues[0].PatientFirstName) $($AccountValues.SerialNumber -join ', ')`""
                CLEAR               = "N"
                TOPRINT_QNTY        = ""
                ADDR5_REIMBEXP      = "NOTHING"
                DUEDATE_SERVICEDATE = "0/0/0"
                TERMS_OTHER2        = ""
            }

            $FinalDocument += $SPLRecord
        }

        if ($AccountGroups.'4040') {

            $AccountValues = $AccountGroups.'4040'
            $SPLRecord = [PSCustomObject]@{
                TRNS_SPL            = "SPL"
                TRNSID_SPLID        = ""
                TRNSTYPE            = if ($AccountValues[0].DocumentType -eq 'Invoice') { "BILL" } elseif ($AccountValues[0].DocumentType -eq 'Credit Memo') { "CREDIT MEMO" } else { "" }
                DATE                = ([Datetime]$AccountValues[0].OrderDate).toString("M\/d\/yyyy")
                ACCNT               = "4040"
                NAME                = ""
                CLASS               = $ClassMapping.$($AccountValues[0].SelltoCustomerNo)
                AMOUNT              = $AccountValues | ForEach-Object { [decimal]$_.UnitPrice * [int]$_.Quantity } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                DOCNUM              = $AccountValues[0].DocumentNo
                MEMO                = "`"$($AccountValues[0].PatientLastName), $($AccountValues[0].PatientFirstName) $($AccountValues.SerialNumber -join ', ')`""
                CLEAR               = ""
                TOPRINT_QNTY        = ""
                ADDR5_REIMBEXP      = ""
                DUEDATE_SERVICEDATE = ""
                TERMS_OTHER2        = ""
            }

            $FinalDocument += $SPLRecord
        }

        
        if ($AccountGroups.'5245') {

            $AccountValues = $AccountGroups.'5245'
            $SPLRecord = [PSCustomObject]@{
                TRNS_SPL            = "SPL"
                TRNSID_SPLID        = ""
                TRNSTYPE            = if ($AccountValues[0].DocumentType -eq 'Invoice') { "BILL" } elseif ($AccountValues[0].DocumentType -eq 'Credit Memo') { "CREDIT MEMO" } else { "" }
                DATE                = ([Datetime]$AccountValues[0].OrderDate).toString("M\/d\/yyyy")
                ACCNT               = "5245"
                NAME                = ""
                CLASS               = $ClassMapping.$($AccountValues[0].SelltoCustomerNo)
                AMOUNT              = $AccountValues | ForEach-Object { [decimal]$_.UnitPrice * [int]$_.Quantity } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                DOCNUM              = $AccountValues[0].DocumentNo
                MEMO                = ""
                CLEAR               = ""
                TOPRINT_QNTY        = ""
                ADDR5_REIMBEXP      = ""
                DUEDATE_SERVICEDATE = ""
                TERMS_OTHER2        = ""
            }

            $FinalDocument += $SPLRecord
        }

        if ($AccountGroups.'6040') {

            $AccountValues = $AccountGroups.'6040'
            $SPLRecord = [PSCustomObject]@{
                TRNS_SPL            = "SPL"
                TRNSID_SPLID        = ""
                TRNSTYPE            = if ($AccountValues[0].DocumentType -eq 'Invoice') { "BILL" } elseif ($AccountValues[0].DocumentType -eq 'Credit Memo') { "CREDIT MEMO" } else { "" }
                DATE                = ([Datetime]$AccountValues[0].OrderDate).toString("M\/d\/yyyy")
                ACCNT               = "6040"
                NAME                = ""
                CLASS               = $ClassMapping.$($AccountValues[0].SelltoCustomerNo)
                AMOUNT              = $AccountValues | ForEach-Object { [decimal]$_.UnitPrice * [int]$_.Quantity } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                DOCNUM              = $AccountValues[0].DocumentNo
                MEMO                = ""
                CLEAR               = ""
                TOPRINT_QNTY        = ""
                ADDR5_REIMBEXP      = ""
                DUEDATE_SERVICEDATE = ""
                TERMS_OTHER2        = ""
            }

            $FinalDocument += $SPLRecord
        }
        $FinalDocument += $TRNSEND
    }
}
catch {
    # Catch errors
    Write-Output "Failed to convert the data. $($Error[0].Exception.Message)"
    Read-Host "Press 'Enter' to exit"
    Return
}

try {
    # Export result to tab delimited file
    Write-Host "Processing successfully completed."
    $FinalDocument | ForEach-Object { "$($_.psobject.properties.value -join "`t")" } | Where-Object { $_.trim() -ne "" } | Out-File "$DestinationFolder\$DestinationFile" -Encoding ascii
    Write-Host "Result is written to $DestinationFolder\$DestinationFile"
    Read-Host "Press 'Enter' to exit"
}
catch {
    # Catch errors
    Write-Output "Failed to output results to $DestinationFile. $($Error[0].Exception.Message)"
    Read-Host "Press 'Enter' to exit"
    Return
}
