# adb
Adaptive Document Builder

A framework for generating simulated malicious office documents.

## Features

* VBA is distinct for every document (level of distinction depends on the adversary document builder selected)
* Random author based on easily updated/replaced name lists (sets local system registry keys before each document build)
* Random file name based on the most commonly seen file names in malicious document campaigns
* Multiple file formats (doc, docm, XML flat OPC)
* Multiple file extensions (.doc, .docm, .rtf)
* Supports multiple payloads
* Functions for building and randomizing VBA are in a shared library for use across multiple adversary builders
* Modular design and architecture for easy addition of more adversary builders
* debug mode that outputs audit trail of document creation details including VBA contents

## Runs on

Python 3 on Windows</br>
COM is used to interface with an installed and configured Office product

## Pre-requisites

- Office installed
- Office opened once to create first-time-run registry entries
- Office must "Trust access to the VBA project object model" (must be checked)</br>
    https://support.office.com/en-us/article/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6
- Python modules in requirements.txt installed

### Run this on a virtual machine!
 - Disable Windows Defender or add an exclusion for the adb files (before cloning) and your output directory or they might get cleaned
 - Registry entries will be changed when setting the author of documents, so don't run this with any production Office software

## Usage

### List available adversary emulation builders

```
>python adb.py -l
sample_with_network_test
underscore_crew_201806
```

### Build documents

Build 5 documents with vba and payload style resembling underscore_crew_201806 (group that delivered agent tesla during this time period)

* Extension: .doc
* File Format: XML flat OPC

```
>python adb.py -a underscore_crew_201806 -c 5 -o C:\users\h\desktop\out -f flatxml -e doc
[*] Building document Sales_Invoice_6619.doc with author: Valentia A Petersen
[*] Building document Your_Invoices_5801.doc with author: Nydia Shields
[*] Building document Selected_Ticket_9047.doc with author: Felipa Henson
[*] Building document Past_Due_Receipt_4278.doc with author: Minh J Mosley
[*] Building document Final_Bill_7431.doc with author: Kaile Perkins
```


### Help Output
```
usage: program_name [-h] [-a ADVERSARY] [-f FILETYPE] [-e EXTENSION]
                    [-c COUNT] [-l] [-o OUTDIR] [-d]

program description

optional arguments:
  -h, --help            show this help message and exit
  -a ADVERSARY, --adversary ADVERSARY
                        -a --adversary {adversary name} (use -l to list)
  -f FILETYPE, --filetype FILETYPE
                        -f --filetype doc | docm | flatxml
  -e EXTENSION, --extension EXTENSION
                        -e --extension doc | docm | rtf
  -c COUNT, --count COUNT
                        -c --count {# of docs to create}
  -l, --listadversaries
                        -l --listadversaries : list available adversaries and
                        exits
  -o OUTDIR, --outdir OUTDIR
                        -o --outdir {path\to\outdir}
  -d, --debug           -d --debug : print debug statements and playbook for
                        each document
```
