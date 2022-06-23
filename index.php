<?php
	const word_dir = __DIR__."/word/";
	const pdf_dir = __DIR__."/pdf/";
	const copy_dir = __DIR__."/copy/";

	$scan = scandir(word_dir);
	$scan = array_diff($scan, [".", ".."]);

	foreach ($scan as $key => $value)
	{
		if (is_file(word_dir.$value) && $value != '~$rdreset.docx')
		{
			$word = new COM("Word.Application") or die ("Could not initialise Object.");
		
			$word->Visible = 0;
			$word->DisplayAlerts = 0;
			$exp = explode(".", $value);

			copy(word_dir."/$value", copy_dir."/hardreset$key.".$exp[1]);	
			
			$word->Documents->Open(copy_dir."/hardreset$key.".$exp[1]);

			// $word->ActiveDocument->SaveAs(__DIR__.'/output2.doc');
			$word->ActiveDocument->ExportAsFixedFormat(pdf_dir."/$key.pdf", 17, false, 0, 0, 0, 0, 7, true, true, 2, true, true, false);
			
			$word->Quit(false);

			rename(pdf_dir."/$key.pdf", pdf_dir."/".$exp[0].".pdf");
			unlink(copy_dir."/hardreset$key.".$exp[1]);
		}
	}