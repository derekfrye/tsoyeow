get-childitem -recurse | 													`
	where-object { 														`
		$_.mode -eq "d----" -and 											`
			( $_.name -eq "Debug" -or $_.name -eq "DebugFromFile" -or $_.name -eq "DebugFromFileNoThread" 		`
			-or $_.name -eq "DebugNoThread" -or $_.name -eq "Release" ) -and 					`
		$_.fullname -notmatch "ExcelXmlWriterNTest\\bin\\Debug" -and $_.name -notmatch "vshost\."			`
	} |															`
	foreach { if ([system.io.directory]::Exists($_.fullname)) { [system.io.directory]::Delete($_.fullname, $TRUE) } }


get-childitem -recurse | 													`
	where-object { 														`
		$_.mode -eq "-a---" -and 											`
		$_.fullname -match "ExcelXmlWriterTest_vs2008" -and 								`
		( $_.name -match "xml\.(rels\.)?(prod|test)$" -or $_.name -match "^\.rels")					`
	} |															`
	foreach { if ([system.io.file]::Exists($_.fullname)) { [system.io.file]::Delete($_.fullname) } }