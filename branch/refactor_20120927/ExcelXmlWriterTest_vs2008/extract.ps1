#load dll
[Reflection.Assembly]::LoadWithpartialName(”System.Xml.Linq”) | Out-Null
# remove old directory
remove-item -recurse -path test
# unzip xlsx file
./unzip test.xlsx -d test | out-null
# instantiate for later use
$a = new-object -typeName "object";

$t=[system.io.path]::getdirectoryname($(get-item test).fullname)+[system.io.path]::DirectorySeparatorChar;
# get all the files from the unzip
get-childitem -recurse test | where-object { $_.mode -eq "-a---" } | Foreach { `
	
	$bb = $_;`
	$cc =$t+$bb.name;
	$dd = $bb.fullname;`
#	echo ($dd+ " " + $cc+".test");`
	[system.io.file]::Copy($dd,($cc+".test"),$TRUE);`
}
# get all the production files
get-childitem -recurse ../ExcelXmlWriterNTest/Resources/Book1_extracted | where-object {$_.mode -eq "-a---" } `
	| foreach { `
#	echo ( $_.name + " " + $_.name+".prod"); `
	[system.io.file]::Copy($_.fullname,($t+$_.name+".prod"),$TRUE);`
}
# load all the files into an xdoc and write back out to clean up formatting
get-childitem  | where-object { $_.name -match "\.test$|\.prod$" -and $_.name -notmatch "unused.xml" } `
	| foreach { `
	$a=$_; `
	$d=[system.xml.linq.xelement]::Load($_.fullname); `
	$d.save($_.fullname) `
}

trap [Exception]{

echo ("error with " + $a.name);
}