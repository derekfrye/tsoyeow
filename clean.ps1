get-childitem -recurse | 	`
	# delete all the Debug* directories
	where-object { 		`
		$_.mode -eq "d----" -and 	`
			( $_.name -eq "Debug" -or $_.name -eq "DebugFromFile" -or $_.name -eq "DebugFromFileNoThread" 		`
			-or $_.name -eq "DebugNoThread" -or $_.name -eq "Release" ) -and 					`
		$_.fullname -notmatch "ExcelXmlWriterNTest\\bin\\Debug" -and $_.name -notmatch "vshost\."			`
		-and $_.fullname -notmatch "ExcelXmlWriterNTest/bin/Debug" `
	} |					`
	foreach { if ([system.io.directory]::Exists($_.fullname)) `
        { `
            echo ("deleting "+$_.fullname+"...");               `
            [system.io.directory]::Delete($_.fullname, $TRUE); `
        } `
    }

get-childitem -recurse |			`
#delete test objects
	where-object { 				`
		$_.mode -eq "-a---" -and 	`
		$_.fullname -match "ExcelXmlWriterTest_vs2008" -and 								`
		( $_.name -match "xml\.(rels\.)?(prod|test)$" -or $_.name -match "^\.rels")					`
	} |					`
	foreach { if ([system.io.file]::Exists($_.fullname)) 
        { `
            echo ("Deleting "+$_.fullname+"..."); `
            [system.io.file]::Delete($_.fullname); `
        } `
    }
    
$pth="ExcelXmlWriterNTest/app.config"; 
$xml = [xml] (get-content $pth)
$node=$xml.configuration.appSettings.SelectNodes("//add[@key='password']")
echo "removing password"
$node.setattribute("value",".")
$node=$xml.configuration.appSettings.SelectNodes("//add[@key='datasource']")
$node.setattribute("value",".")
echo "removing datasource"
$xml.save((get-item $pth).fullname)
