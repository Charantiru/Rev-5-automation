import pandas as pd
import xml.etree.ElementTree as ET
import re
from xml.dom import minidom
import sys
import getopt

# --------------------------------------------------------------------HOW TO USE----------------------------------------------------------------------------

## HAVE 3 FILES IN A FOLDER: This python code, FEDRAMP excel sheet, NIST xml catalog

## OPEN THIS CODE IN A CODE EDITOR WITH A TERMINAL TO RUN

## IN TERMINAL, ENTER THE THREE PARAMTERS IN THIS FORMAT -e <excel path input> -s <sheetname> -n <nist xml relative path> -o <output xml>

## YOU DO NOT HAVE TO CREATE A NEW XML FILE... THE OUTPUT XML YOU ENTER WILL GENERATE A NEW FILE UNDER THAT NAME

## Use -h or --help TO GET A BETTER UNDERSTANDING OF HOW TO ENTER PARAMTERS IN COMMAND LINE


## CODE SHOULD LOOK LIKE

## python ParamaterParser.py -e "C:\Users\M33084\Internship\Paramater Parser\FedRAMP_Security_Controls_Baseline_Rev 5 (Public Comment) 2021_12_02 (3).xlsx" 
# -s "High Baseline Controls" -n "NIST_SP-800-53_rev5_catalog_main.xml" -o "items2.xml"

# --------------------------------------------------------------------HOW TO USE----------------------------------------------------------------------------


def ParamaterParser(FEDRAMP,sheetname, NIST, xmlfile):


    df1=pd.read_excel(r"" + FEDRAMP, sheet_name=sheetname)
    df1=df1.iloc[1:]
    df1.rename(columns={"Unnamed: 0": "Index" , "Unnamed: 1": "Sort ID", "NIST 800-53 Security Controls Catalog Revision 5": "Family", 
    "Unnamed: 3": "ID", "Unnamed: 4": "Control Name","Unnamed: 5": "Nist Control Description","Unnamed: 6": "Fedramp Assignment/Paramters","Unnamed: 7": "Requirements and Guidance"}, inplace=True)

    controllist=[] #This is list of all the controls in the format eg. "AC-1 (c) (1)"
    updatedcontrollist=[] # This is a list of list of all the ref-id NIST controls given a value from controllist. eg.['ac-1_smt.c.1','ac-01_odp.05']
    paramaterlist=[] #This is a list of parameters for each control from the controlist eg.['at least annually] ', 'significant changes]']
    finaldict={} # This is a dictionary of updated controls to paramaters. eg.{'ac-1_smt.c.1':'at least annually}

    ## This method translates all the control ids into the form needed to search in NIST catalog to find the correct param-id to map
    def create_id(target):
        newstring=target
        newstring= newstring.replace(" ","")
        finalstring=""
        check=False
        for i in range(len(newstring)):
            if i>3 and newstring[i]=="-":
                break
            if newstring[i] == "(":
                if newstring[i+1].isdigit():
                    finalstring+="."
                continue
            elif newstring[i] == ")":
                continue
            elif newstring[i].isspace():
                continue
            elif newstring[i].islower():
                check=True
                finalstring+= "_smt." + newstring[i]
                if newstring[i+1].isdigit():
                    finalstring+="." + newstring[i+1]
            elif newstring[i]=="0" and newstring[i-1]=="(":
                continue
            else:
                finalstring += newstring[i]
        if check==False:
            finalstring += "_smt"
            
        return(finalstring.lower())



    tree = ET.parse(NIST)
    root=tree.getroot()

    ## This method uses the control id that was generated from create_id to find the param-id's
    def findidref(control):
        controlist=[]
        for target in root.findall(".//{http://csrc.nist.gov/ns/oscal/1.0}part[@id='" + control +"']/{http://csrc.nist.gov/ns/oscal/1.0}p/{http://csrc.nist.gov/ns/oscal/1.0}insert"):
            controlist.append(target.attrib["id-ref"])
        updatedcontrollist.append(controlist)

    ## This method creates a list of control ids and list of list of paramters
    for i in df1["Fedramp Assignment/Paramters"].dropna():
        res= re.findall(r'[A-Z]{2}[-][0-9]+[()0-9a-z\s-]*[[].+|[A-Z]{2}[-][0-9]+[()0-9a-z\s-]*', i)
        for splitline in res:
            splitted= splitline.split("[")
            controllist.append(splitted[0])
            splitted=splitted[1:]
            for value in range(len(splitted)):
                splitted[value]= splitted[value].replace("]","")
            paramaterlist.append(splitted)


    ## This runs the control ids from controllist through create_id and findidref
    for i in controllist:
        id=create_id(i)
        findidref(id)


    ##This part of the code updates the dictionary for each of the updated controls mapped to its respective paramters.
    ##If the updated control id cannot find a paramater, its value becomes TBD as it tells the user a paramater is missing from the Excel document
    
    for i in range(len(paramaterlist)):
        for v in range(len(updatedcontrollist[i])):
            if v> len(paramaterlist[i])-1:
                continue
            else:
                finaldict[updatedcontrollist[i][v]]= paramaterlist[i][v]



    ##This last part of the document translates everything necessary into an xml file
    root=minidom.Document()
    xml= root.createElement("root")
    root.appendChild(xml)

    def createxml(dict):
        for key,value in dict.items():
            set= root.createElement("set-parameter")
            set.setAttribute("param-id", key)
            xml.appendChild(set)
            constraint= root.createElement("constraint")
            set.appendChild(constraint)
            description= root.createElement("description")
            constraint.appendChild(description)
            p= root.createElement("p")
            description.appendChild(p)
            text = root.createTextNode(value)
            p.appendChild(text)
            xml_str= root.toprettyxml(indent="\t")
            with open(xmlfile, "w") as f:
                f.write(xml_str)

    createxml(finaldict)


def commandline(argv):
    arg_excel = ""
    arg_nist = ""
    arg_output = ""
    arg_sheetname = ""
    arg_help = "-e <excel path input> -s <sheetname> -n <nist xml relative path> -o <output xml>"
    try:
        opts, args = getopt.getopt(argv[1:], "he:s:n:o:", ["help", "input=", 
        "sheetname=", "nist=", "output="])
    except:
        print(arg_help)
        sys.exit(2)
    
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(arg_help)  # print the help message
            sys.exit(2)
        elif opt in ("-e", "--excel"):
            arg_excel = arg
        elif opt in ("-s", "--sheetname"):
            arg_sheetname = arg
        elif opt in ("-n", "--nist"):
            arg_nist= arg
        elif opt in ("-o", "--output"):
            arg_output = arg
    
    ParamaterParser(arg_excel, arg_sheetname, arg_nist, arg_output)


if __name__ == "__main__":
    commandline(sys.argv)