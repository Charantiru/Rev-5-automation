## Overview

These parsers can be used to generate OSCAL formatted content for Rev 5 baselines using the excel source FEDRAMP data. 

## What is needed to run the code
- Python 3
- Pandas libararies
- Regex libraries
- minidom libraries form xlm.dom
- sys module
- get opt module
- xml.etree.Elementree module(Paramater Parser specific)

## Paramater Parser

Have 3 files in a folder: this python code, fedramp excel sheet, nist xml catalog

Open this code in a code editor with a terminal to run

In terminal, enter the three paramters in this format -e <excel path input> -s <sheetname> -n <nist xml relative path> -o <output xml>

You do not have to create a new xml file... the output xml you enter will generate a new file under that name

Use -h or --help to get a better understanding of how to enter paramters in command line


Code should look like

```
python ParamaterParser.py -e "C:\Users\M33084\Internship\Paramater Parser\FedRAMP_Security_Controls_Baseline_Rev 5 (Public Comment) 2021_12_02 (3).xlsx" 
# -s "High Baseline Controls" -n "NIST_SP-800-53_rev5_catalog_main.xml" -o "items2.xml"
```



## Guidance/Requirement Parser


Have 2 files in a folder: this python code, fedramp excel sheet

Open this code in a code editor with a terminal to run

In terminal, enter the three paramters in this format -e <excel path input> -s <sheetname> -o <output xml>

You do not have to create a new xml file... the output xml you enter will generate a new file under that name

Use -h or --help to get a better understanding of how to enter paramters in command line


Code should look like

```
python Guidance_Requirement_Parser.py -e "C:\Users\M33084\Internship\Guidance Parser\FedRAMP_Security_Controls_Baseline_Rev 5 (Public Comment) 2021_12_02 (3).xlsx"
-s "High Baseline Controls" -o "guidance_requirement_high.xml"
```


## Future Enhancement
This Command Line Interface tool could be containerized so that it could be run directly from a terminal



