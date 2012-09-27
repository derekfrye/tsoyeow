# find all matching original files from teh unzip and replace with the formatted counterparts
remove-item atest.xlsx | out-null
get-childitem | where-object {$_.name -match "\.test$"} | foreach { `
	$b =$_; `
	$a = $(get-childitem -recurse test `
		| where-object{ $_.name -eq $b.name.substring(0,$b.name.length-5)} ); `
	[system.io.file]::Copy($b.fullname,$a.fullname,$TRUE) `
}
cd test
.././zip.exe -r atest.xlsx * | out-null
cd ..
move-item -path test/atest.xlsx .