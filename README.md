dt_format
=========

A collection of tools that I use for formatting translations

ATM, this tool will take translations by WISA and convert the .docx to rtf and apply all of Pervy's Page and Panel annotations to the converted RTF and write it out.
If anyone wanted to use this to format any other translators stuff to Pervy's format I would just need to make some corrections to allow PAGE_DELIMITER and PANEL_DELIMITER arguments to be passed from the command line

- usage: python deathtoll_proof_formatter "c:\some_trans.docx" "c:\formatted_trans.rtf"
- arg 1 = input file
- arg 2 = output file