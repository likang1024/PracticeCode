
Function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

#打开一个Excel文档
Function OpenExcel([string]$ExFileName) 
{
    if (!(Test-Path $ExFileName)) {
        Write-Host "$ExFileName is not exist."
        return;
    }

    #
    $Excel = new-object -comobject excel.application
    if (!$Excel)
    {
        Write-Host "Create Excel Failed."
        return;
    }
    $Excel.Visible = $true
    $ExFile = $Excel.Workbooks.Open($ExFileName)
    if (!$ExFile)
    {
		Write-Host "Open $ExFileName Failed."
		return;
    }
    
    return $Excel,$ExFile

}

#只有OpenExcel 成功是， 才需要CloseExcel
Function CloseExcel($Excel, $ExFile)
{
	if (!$Excel) {
        Write-Host "Parameter Excel is null."
        return
    }
	$Excel.quit()
    
    Release-Ref $Excel
	return
}

Function SaveExFile($ExFile){
	if (!$ExFile) {
		Write-Host "Parameter ExFile is null."
        return
	}
	$ExFile.save()
    $ExFile.close()
	Release-Ref $ExFile
}

Function CloseExFile($ExFile){
	if (!$ExFile) {
		Write-Host "Parameter ExFile is null."
        return
	}
	$ExFile.close()
	Release-Ref $ExFile
}

Function OpenSheet($Excel,$SheetName)
{
    if (!$Excel)
    {
    Write-Host "Err:Excel object is null!"
    return
    }
    $Exsheet = $Excel.worksheets | where {$_.name -eq $SheetName}
    if (!$Exsheet)
    {
    Write-Host "Err:Fail to open sheet $SheetName"
    return
    }
    $Exsheet.activate()
    return $Exsheet
}

Function ConvertAtoi($Str)
{
    
    if ($Str -match "\d+")
    {
        Write-Host "Info:ConvertAtoi $Str already a number"
        return $Str
    }
    
    if (!($str -match "[a-z]+"))
    {
        Write-Host "Err:ConvertAtoi $Str is not a valid input!"
        return 0
    }
       
    [int]$iValue =0;
    $Str = $Str.toUpper()
    for ($i=0; $i -lt $Str.Length; $i++)
    {
        $iValue = $iValue*26
        $Cha = $Str.SubString($i, 1)
        switch ($Cha)
        {
            "A" {$iValue += 1}
            "B" {$iValue += 2}
            "C" {$iValue += 3}
            "D" {$iValue += 4}
            "E" {$iValue += 5}
            "F" {$iValue += 6}
            "G" {$iValue += 7}
            "H" {$iValue += 8}
            "I" {$iValue += 9}
            "J" {$iValue += 10}
            "K" {$iValue += 11}
            "L" {$iValue += 12}
            "M" {$iValue += 13}
            "N" {$iValue += 14}
            "O" {$iValue += 15}
            "P" {$iValue += 16}
            "Q" {$iValue += 17}
            "R" {$iValue += 18}
            "S" {$iValue += 19}
            "T" {$iValue += 20}
            "U" {$iValue += 21}
            "V" {$iValue += 22}
            "W" {$iValue += 23}
            "X" {$iValue += 24}
            "Y" {$iValue += 25}
            "Z" {$iValue += 26}
        }
    }
    
    
    return $iValue
    
    #ConvertAtoi "CD"
}

#写修订记录
Function WriteModifyHistory($Excel, $AppendOrNo=$False, $ModifyInfo, $SiteName="", $Modfier="李康", $Comments="CR")
{
    if (!$Excel) {
        write-host "Err:Parameter Excel is null"
        return
    }
    
    #写修订记录
    $Exsheet = OpenSheet $Excel "Change History"
    if (!$Exsheet) {
        return 
    }
    
    $row=1009
    do {
       $value =  $Exsheet.Cells.Item($row,1).Value2
       $row++
    } while(![String]::IsNullOrEmpty($value))
    $row-=1
    
    #是在已有的修订记录上追加， 还是新写一行 
    if ($AppendOrNo) { #追加
        $row -=1;
        $OldModifyInfo = $Exsheet.Cells.Item($row,4).Value2
        $Exsheet.Cells.Item($row,4).Value2 =  "$OldModifyInfo`n$ModifyInfo"
    } else {
        #新写一行
        $Exsheet.Cells.Item($row,1).Value2 =  get-date -uFormat "%Y/%m/%d"
        $Exsheet.Cells.Item($row,3).Value2 =  $Exsheet.Cells.Item($row-1,3).Value2
        $Exsheet.Cells.Item($row,4).Value2 =  "$SiteName`n$ModifyInfo"
        $Exsheet.Cells.Item($row,5).Value2 =  $Modfier
        $Exsheet.Cells.Item($row,6).Value2 =  "$SiteName $Comments"
    }  
}

#修改一个标志位取值并写修订记录
Function ProOneFlag($Excel, $SiteName, $iColumn, $Flag, $Index, $FlagValue, $Comments="CR",$AppendOrNo=$False,$ModifyComments="")
{
    if (!$Excel) {
        write-host "Err:Parameter Excel is null"
        return
    }
    
    #先打开局点说明书，找到局点对应的列
    #Write-Host "Info:Start to ProOneFlag."
    
    $SpecFlag=@{
			"OCGBase_SystemInfo.ServiceCtrlFlag"="OCGBase_SystemInfo.SerCtrlFlag";
			"OCG_NAConfig.ServiceCtrlFlag"="OCG_NAConfig.SerCtrlFlag";
			"OCG_SystemInfo.CallPromptCtrlFlag"="OCG_SystemInfo.CallPrmptCtrlFlg";
			"OCG_VPN_SysInfo.VPNServiceFlag"="OCG_VPN_SysInfo.ServiceFlag"}
            
    #检查是否为不对应的excel sheet
	if (![string]::IsNullOrEmpty($SpecFlag[$Flag])) {
	    $Flag = $SpecFlag[$Flag]
	}
    $Exsheet = OpenSheet $Excel $Flag 
    if (!$Exsheet) {
        return 
    }
    
    #校验$Value是否在取值范围内，暂时不做
    #如果目标取值和当前取值相同则跳过
    $OldValue = ""
    $OldValue = $Exsheet.Cells.Item(([int]$Index+[int]1), $iColumn).Value2
    if ($OldValue -eq $FlagValue) {
        Write-Host "Info:$SiteName $Flag$Index=$FlagValue,New value and old value is  the same."
        return
    }
    
    #修改这个标志位
    $Exsheet.Cells.Item(([int]$Index+[int]1), $iColumn).Value2 = $FlagValue
    $ModifyInfo = "修改$Flag[$Index] $OldValue -> $FlagValue $ModifyComments"
    
    #写修订记录
    WriteModifyHistory  $Excel $AppendOrNo  $ModifyInfo $SiteName "李康" $Comments
    
    Write-Host  "Info:$SiteName 修改$Flag[$Index] $OldValue -> $FlagValue $ModifyComments $Comments"
    #Write-Host "Info:ProOneFlag run completed."
    
    return
      
}



