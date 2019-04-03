$processFilterList ="*Python*", "*sql*", "*conhost*"
$diskFilterList = "*_Total*"
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$xlChart = [Microsoft.Office.Interop.Excel.XlChartType]

$counters = @{
    "Process time"              = @{
        "xpath"            = "\\*\Process(*)\% Processor Time";
        "chartType"        = $xlChart::xlLine;
        "valueDivsor"      = 100;
        "chartTitle"       = "Process %";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Processor Usage % (All Cores)";
        "dataNumberFormat" = "0.00%";
        "headerFilters"    = $processFilterList;
    };
    "Process Memory Nonpaged"   = @{
        "xPath"            = "\\*\Process(*)\Pool Nonpaged Bytes";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "Process NonPaged Memory";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Processor Usage % (All Cores)";
        "dataNumberFormat" = '[<500000000]#,##0.00,," MB";#,##0.00,,," GB"'
        "headerFilters"    = $processFilterList;
    };
    "Process Memory WorkingSet" = @{
        "xPath"            = "\\*\Process(*)\Working Set";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "Processs Memory Working Set";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Memory Working Set";
        "dataNumberFormat" = '[<500000000]#,##0.00,," MB";#,##0.00,,," GB"'
        "headerFilters"    = $processFilterList;
    };
    "Process IOBytes"           = @{
        "xPath"            = "\\*\Process(*)\IO Data Bytes/sec";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "Total Process IO Bytes/sec";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "IO Bytes/Sec";
        "dataNumberFormat" = '[<500000000]#,##0.00,," MBps";#,##0.00,,," GBps"'
        "headerFilters"    = $processFilterList;
    };
    "Process IOps"              = @{
        "xPath"            = "\\*\Process(*)\IO Data Operations/sec";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "Total Process IO Operations/sec";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "IO Operations/sec";
        "dataNumberFormat" = '0.00'
        "headerFilters"    = $processFilterList;
    };
    "System Processor Time"     = @{
        "xpath"            = "\\*\Processor(*)\% Processor Time";
        "chartType"        = $xlChart::xlLine;
        "valueDivsor"      = 100;
        "chartTitle"       = "Process %";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Processor Usage % (Average of All Cores)";
        "dataNumberFormat" = "0.00%"
        "headerFilters"    = $diskFilterList;
    };
    "System MemoryFreeMbytes"   = @{
        "xPath"            = "\\*\Memory\Available MBytes";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "System Available Memory";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Available Memory MB";
        "dataNumberFormat" = '0.00" MB"'
    };
    "System AvgDiskBytes"       = @{
        "xPath"            = "\\*\PhysicalDisk(*)\Avg. Disk Bytes/Transfer";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "Average Disk Bytes/Transfer"
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Disk Byte/Transfer";
        "dataNumberFormat" = '[<500000000]#,##0.00,," MB";#,##0.00,,," GB"';
        "headerFilters"    = $diskFilterList;
    };
    "System AvgDiskQueueLength" = @{
        "xPath"            = "\\*\PhysicalDisk(*)\Avg. Disk Queue Length";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "System Average Disk Queue Length";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Disk Queue Length";
        "dataNumberFormat" = "0.00";
        "headerFilters"    = $diskFilterList;
    };
    "System DiskTransfers"      = @{
        "xPath"            = "\\*\PhysicalDisk(*)\Disk Transfers/sec";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "System Disk Transfers/sec";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Disk Transfers/sec";
        "dataNumberFormat" = "0.00";
        "headerFilters"    = $diskFilterList;
    };
    "System DiskBytes"          = @{
        "xPath"            = "\\*\PhysicalDisk(*)\Disk Bytes/sec";
        "chartType"        = $xlChart::xlLine;
        "chartTitle"       = "System Disk Bytes/sec";
        "xAxisTitle"       = "Elasped Time";
        "yAxisTitle"       = "Disk Bytes/sec";
        "dataNumberFormat" = '[<500000000]#,##0.00,," MBps";#,##0.00,,," GBps"';
        "headerFilters"    = $diskFilterList;
    };
}

function GenerateChart($ws, $counter) {
    $sizeMultipler = 2.5
    $chart = $ws.Shapes.AddChart().Chart
    $chart.ChartType = $counter.chartType
    
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $counter.chartTitle

    $yAxis = $chart.Axes([Microsoft.Office.Interop.Excel.XLAxisType]::xlValue, 
        [Microsoft.Office.Interop.Excel.XLAxisGroup]::xlPrimary)

    $xAxis = $chart.Axes([Microsoft.Office.Interop.Excel.XLAxisType]::xlCategory, 
        [Microsoft.Office.Interop.Excel.XLAxisGroup]::xlPrimary)


    $xAxis.HasTitle = $true
    $yAxis.HasTitle = $true
    $xAxis.AxisTitle.Text = $counter.xAxisTitle
    $yAxis.AxisTitle.Text = $counter.yAxisTitle

    $chart.ChartArea.Height = $chart.ChartArea.Height * $sizeMultipler
    $chart.ChartArea.Width = $chart.ChartArea.Width * $sizeMultipler
    $chart.ChartArea.Top = 20
    $chart.ChartArea.Left = 100    
}

function ConvertToExcelProcess($csv_file, $counter) {
    Write-Host "Converting $($csv_file) using $($counter.xPath)"
    $xlWorkbookDefault = 51
    $sampleTime = 1
    $xlsx = $csv_file -replace '\.csv$', '.xlsx'

    [Microsoft.Office.Interop.Excel.ApplicationClass]$excel = New-Object -ComObject "Excel.Application"
    $excel.Workbooks.OpenText($csv_file)
    $wb = $excel.Workbooks(1)
    $wb.SaveAs($xlsx, $xlWorkbookDefault)
    $ws = $wb.Worksheets.Item(1)
    $nws = $wb.Worksheets.Add()
    $used = $ws.UsedRange
    $firstColumn = $ws.Columns(1).Value2
    
    $nwsColumnIndex = 1
    for ($j = 2; $j -le $used.Columns.Count; $j++) {
        $column = $used.Columns($j)
        $header = $column.Rows(1)
        if (!$counter.headerFilters -or ($counter.headerFilters | Where-Object { $header.Text -like $_ })) {
            $nwsColumnIndex++
            $head = $header.Text
            if ($counter.xpath -like "*(*)*") {
                $head = $head.Split("(")[1].Split(")")[0]
            } else {
                $head =  $head.Split("\")[-1]
            }
            $nws.Columns($nwsColumnIndex).Rows(1) = $head
            for ($i = 2; $i -le $used.Rows.Count; $i++) {
                $val = $column.Rows($i).Text
                if ($counter.valueDivsor) {
                    $val = $val / $counter.valueDivsor
                }
                $nws.Columns($nwsColumnIndex).Rows($i) = $val
            }
        }
    }
    $nws.UsedRange.NumberFormat = [string]$counter.dataNumberFormat
    $nws.Columns(1).Value2 = $firstColumn
    
    if ($nws.Columns(2).Rows(2).Text.Trim() -eq "") {
        Write-Host $nws.Columns(2).Rows(2).Text
        $nws.Rows(2).Delete()
    }

    $timeCount = 0
    for ($i = 2; $i -le $nws.UsedRange.Rows.Count; $i++) {      
        $nws.UsedRange.Columns(1).Rows($i).Value2 = "=$($timeCount)/86400"
        $timeCount += $sampleTime;
    }
    $ws.Delete()
    $nws.Columns(1).NumberFormat = "h:mm:ss"

    $ws = $nws
    GenerateChart $ws $counter
    $wb.Save()
    $wb.Close()
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Remove-Variable excel
    Remove-Variable ws
    Remove-Variable nws
    Remove-Variable wb
    Start-Sleep -Seconds 5
    Stop-Process -Name excel -Force 
}

function GenerateCounters([System.IO.FileInfo]$perf_file) {
    foreach ($counter in $counters.GetEnumerator()) {
        Write-Host("relog $($perf_file.FullName) -c '$($counter.Value.xPath)' -f csv -o '$($perf_file.DirectoryName)\$($counter.Name).csv' -y")
        relog.exe "$($perf_file.FullName)" -c "$($counter.Value.xPath)" -f csv -o "$($perf_file.DirectoryName)\$($counter.Name).csv" -y       
        $csv_file = "$($perf_file.DirectoryName)\$($counter.Name).csv"
        ConvertToExcelProcess $csv_file $counter.Value
        (Get-Item $csv_file).Delete()
    }
}


foreach ($perf_file in Get-ChildItem -Recurse -Include *.blg) {
    GenerateCounters($perf_file)
}

