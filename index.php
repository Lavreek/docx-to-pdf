<?php
	const word_dir = __DIR__."/word/";
	const pdf_dir = __DIR__."/pdf/";
	const copy_dir = __DIR__."/copy/";

	const convert_to = ".pdf";

	$scan = scandir(word_dir);
	$scan = array_diff($scan, [".", "..", ".gitignore"]);

	foreach ($scan as $key => $value)
	{
		if (is_file(word_dir.$value))
		{
			$word = new COM("Word.Application") or die ("Could not initialise Object.");
		
			$word->Visible = 0;
			$word->DisplayAlerts = 0;
			$exp = explode(".", $value);

			copy(word_dir.$value, copy_dir."hardreset$key.".$exp[1]);	
			
			$word->Documents->Open(copy_dir."hardreset$key.".$exp[1]);

			if (convert_to == ".pdf")
			{
				$word->ActiveDocument->ExportAsFixedFormat(pdf_dir.$key.convert_to, 17, false, 0, 0, 0, 0, 7, true, true, 2, true, true, false);
				rename(pdf_dir.$key.convert_to, pdf_dir.$exp[0].convert_to);
			}
			else
				$word->ActiveDocument->SaveAs(pdf_dir.$exp[0].convert_to);
			
			$word->Quit(false);
		}
	}