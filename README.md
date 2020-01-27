# A simple python bib-to-docx parser

As the title would seem to imply:   
takes a bib-tex bibliography file as input, and outputs a formatted docx file.

'bib_to_docx.py' will take 'bib_in.bib' and produce the 'bib_out.docx'.

Haven't packaged into a more reuseable form yet, nor have I added methods for tuning the docx formatting. Haven't even setup reading arguments from the command line (filenames are hardcoded now).

So, this is a very early stages type of thing, and it just parses @article bib entries, but extending shouldn't be hard. Just a quick hack project to fool around with regex and the python-docx (v0.8.1, btw) package.

P.S.: the bib file here was an export from the Web of Knowledge website.