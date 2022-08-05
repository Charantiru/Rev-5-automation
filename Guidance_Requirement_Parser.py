import pandas as pd
import re
from xml.dom import minidom
import sys
import getopt


# --------------------------------------------------------------------HOW TO USE----------------------------------------------------------------------------

## HAVE 2 FILES IN A FOLDER: This python code, FEDRAMP excel sheet

## OPEN THIS CODE IN A CODE EDITOR WITH A TERMINAL TO RUN

## IN TERMINAL, ENTER THE THREE PARAMTERS IN THIS FORMAT -e <excel path input> -s <sheetname> -o <output xml>

## YOU DO NOT HAVE TO CREATE A NEW XML FILE... THE OUTPUT XML YOU ENTER WILL GENERATE A NEW FILE UNDER THAT NAME

## Use -h or --help TO GET A BETTER UNDERSTANDING OF HOW TO ENTER PARAMTERS IN COMMAND LINE


## CODE SHOULD LOOK LIKE

## python Guidance_Requirement_Parser.py -e "C:\Users\M33084\Internship\Guidance Parser\FedRAMP_Security_Controls_Baseline_Rev 5 (Public Comment) 2021_12_02 (3).xlsx"
## -s "High Baseline Controls" -o "guidance_requirement_high.xml"

# --------------------------------------------------------------------HOW TO USE----------------------------------------------------------------------------





def guidance_requirement_parser(FEDRAMP,sheetname, xmlfile):

    ## The below 4 lines of code allows the variable df1= the guidance and requirement column without any empty rows
    df=pd.read_excel(r"" + FEDRAMP, sheet_name=sheetname)
    df=df.iloc[1:]
    df.rename(columns={"Unnamed: 0": "Index" , "Unnamed: 1": "Sort ID", "NIST 800-53 Security Controls Catalog Revision 5": "Family", 
    "Unnamed: 3": "ID", "Unnamed: 4": "Control Name","Unnamed: 5": "Nist Control Description","Unnamed: 6": "Fedramp Assignment/Paramters","Unnamed: 7": "Requirements and Guidance"}, inplace=True)

    df1=df["Requirements and Guidance"].dropna()


    controllist =[] #This is a list of all the controls
    requirementlist=[] #This is a list that has all the descriptions with "Requirement:...."
    guidancelist=[] #This is a list that has all the descriptions with "Guidance:...."

    ## This method creates the id given the control id from controllist. Turns from AC-3 (4)----> ac-3.4
    def create_id(target):
        newstring=target
        newstring= newstring.replace(" ","")
        finalstring=""
        check=False
        for i in range(len(newstring)):
            if i>3 and newstring[i]=="-":
                break
            if newstring[i] == "(":
                if i+1<len(newstring):
                    if newstring[i+1].isdigit():
                        finalstring+="."
                continue
            elif newstring[i] == ")":
                continue
            else:
                finalstring += newstring[i]
            
        return(finalstring.lower())


    ## This code creates the controllist, requirementlist, and guidancelist
    for line in df1:
        requirements= re.findall(r'[rR][e][q][u][i][r][e][m][e][n][t][s]?[:]\s*(.+)',line)
        requirementlist.append(requirements)
        guidance= re.findall(r'[Gg][u][i][d][a][n][c][e][:]\s+(.+)',line)
        guidancelist.append(guidance)
        controlid= re.findall(r'[A-Z]{2}[-][0-9\s]+[()0-9-]{2,5}|[A-Z]{2}[-][0-9]+',line)
        for i in controlid:
            controllist.append(i)
            break

        

    root=minidom.Document()
    xml= root.createElement("root")
    root.appendChild(xml)

    ## This method transfers all the information from the 3 lists into the necessary format into an xml
    def createxml(controllist,requirementlist, guidancelist):
        for i in range(len(controllist)):
            set= root.createElement("alter")
            set.setAttribute("control-id", create_id(controllist[i]))
            xml.appendChild(set)

            add= root.createElement("add")
            add.setAttribute("position", "ending")
            add.setAttribute("by-id", create_id(controllist[i])+"_smt")
            set.appendChild(add)

            part= root.createElement("part")
            part.setAttribute("id", create_id(controllist[i])+"_fr")
            part.setAttribute("name", "item")
            add.appendChild(part)

            title= root.createElement("title")
            part.appendChild(title)
            text = root.createTextNode(controllist[i] + " Additional FedRAMP Requirements and Guidance")
            title.appendChild(text)

            for index in range(len(requirementlist[i])):
                newpart=root.createElement("part")
                newpart.setAttribute("id", create_id(controllist[i]) + "_fr_smt." + str(index+1))
                newpart.setAttribute("name", "item")
                part.appendChild(newpart)

                prop= root.createElement("prop")
                prop.setAttribute("name", "label")
                prop.setAttribute("value", "Requirement:")
                newpart.appendChild(prop)

                p= root.createElement("p")
                newpart.appendChild(p)
                text=root.createTextNode(requirementlist[i][index])
                p.appendChild(text)

            for index in range(len(guidancelist[i])):
                newpart=root.createElement("part")
                newpart.setAttribute("id", create_id(controllist[i]) + "_fr_gdn." + str(index+1))
                newpart.setAttribute("name", "guidance")
                part.appendChild(newpart)

                prop= root.createElement("prop")
                prop.setAttribute("name", "label")
                prop.setAttribute("value", "Guidance:")
                newpart.appendChild(prop)

                p= root.createElement("p")
                newpart.appendChild(p)
                text=root.createTextNode(guidancelist[i][index])
                p.appendChild(text)



            xml_str= root.toprettyxml(indent="\t")
            with open(xmlfile, "w") as f:
                f.write(xml_str)

    createxml(controllist, requirementlist, guidancelist)

def commandline(argv):
    arg_excel = ""
    arg_output = ""
    arg_sheetname = ""
    arg_help = "-e <excel path input> -s <sheetname> -o <output xml>"
    try:
        opts, args = getopt.getopt(argv[1:], "he:s:o:", ["help", "excel=", 
        "sheetname=", "output="])
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
        elif opt in ("-o", "--output"):
            arg_output = arg
    
    guidance_requirement_parser(arg_excel, arg_sheetname, arg_output)


if __name__ == "__main__":
    commandline(sys.argv)

