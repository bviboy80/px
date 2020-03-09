'''
ParseFixedWidth

Parse a fixed width file and export the results as a CSV for BCC Mail Manager import.
Script uses "Address Line Starting Point" to split Name and Address line fields

(Optional) Third parameter is a text file containing list of company numbers to filter.


Original data Format:

Company Number
Account Number
Company Name
Account Registration Line 1
Account Registration Line 2
Account Registration Line 3
Account Registration Line 4
Account Registration Line 5
Account Registration Line 6
Address Line Starting Point
Status


Format for Mail Manager Import:

CompAcctNo
FullName
Company
Company2
Company3
Company4
Company5
Delivery Address
Alternate 1 Address
Alternate 2 Address
Alternate 3 Address
City-State-Zip
AddrStartPoint
Status

'''

___author__ = "Shaun Thomas"
___date__ = "March 20, 2019"
__version__ = 1.0

import sys
import struct
import csv
import os
import codecs
import re

def main(argv=None):
    
    inFile = os.path.abspath(sys.argv[1])  
    outputDir = os.path.dirname(inFile)
    filename = os.path.basename(inFile)
    baseFilename = filename.split('.')[0]
 
    with open(inFile, 'rb') as d:

        record_dict = {"eligible" : [],
                       "ineligible": [],
                       "filtered": []}

        parse = makeDataParser()
        addrStartList = []

        comp_filter_list = createCompanyFilterList(outputDir)
        
        for recordcount, line in enumerate(d, start=1):
            ascii_line = replaceNonAsciiChars(line, recordcount)
            parsedLine = [x.strip() for x in parse(ascii_line)]
            sortRecords(parsedLine, record_dict, addrStartList, comp_filter_list, recordcount)

            if recordcount % 50000 == 0:
                print "{} records reviewed".format(recordcount)
                            
        writeNCOARecordstoFile(outputDir, record_dict)
        writeFilteredFile(outputDir, record_dict)
        createCounts(outputDir, filename, record_dict, comp_filter_list, addrStartList)     


def createCompanyFilterList(outputDir):
    filterFile = os.path.join(outputDir, "records_to_filter.txt")
    
    comp_filter_list = []

    if os.path.exists(filterFile):
        with open(filterFile, 'rb') as n:
            for line in n:
                comp_filter_list.append(line.strip())
                
            print "\n\n\n"
            print "The following Company Numbers will be check for filtering"
            print "{}".format(", ".join(comp_filter_list))
            print "\n\n\n"
    else:
        print "No filter list. Ensure the filter list is"
        print "in the same location as the data.\n"

    return comp_filter_list


def makeDataParser():
    """ Create a list of slice objects to parse fields 
    in a line. Accumulate the sum of the field widths 
    to get the indices for each field. """
    
    formatStr = "5s 10s 40s 38s 38s 38s 38s 38s 38s 1s 10s"
    fieldstruct = struct.Struct(formatStr)
    parse = fieldstruct.unpack_from
    return parse

    
def replaceNonAsciiChars(line, recordcount):
    """ Convert byte text to unicode chars. Replace non-ASCII,
    the "replacement", "non-breaking space" and "Broken Bar" 
    chars. Convert chars back into bytes. """
    
    char_text = line.decode("utf-8", errors='replace').replace(u'\ufffd', " ")
    replace_nbspace = char_text.replace(u'\u00A6', " ")
    replace_bknbar = replace_nbspace.replace(u'\u00A0', " ")
    latin_line = replace_bknbar.encode("latin-1")
    ascii_char = latin_line.decode("ascii", errors='replace').replace(u'\ufffd', " ")
    ascii_line = ascii_char.encode("ascii")
    if len(ascii_line) != 295:
        print "\nCharacter error found in line {}!".format(recordcount)
        print "Line is now {}!\n".format(len(ascii_line))
        print line
        raise UnicodeError
    return ascii_line


def sortRecords(parsedLine, record_dict, addrStartList, comp_filter_list, recordcount):

    org_hdr = ["Company Number", "Account Number", "Company Name",
                 "Account Registration Line 1", "Account Registration Line 2",
                 "Account Registration Line 3", "Account Registration Line 4",
                 "Account Registration Line 5", "Account Registration Line 6",
                 "Address Line Starting Point", "Status"]

    # Check if record should be filtered.
    compNo = parsedLine[org_hdr.index("Company Number")]
    acctNo = parsedLine[org_hdr.index("Account Number")]
    
    if compNo in comp_filter_list or \
    "{}{}".format(compNo, acctNo) in comp_filter_list:
        name = parsedLine[org_hdr.index("Account Registration Line 1")]
        record_dict["filtered"].append([recordcount, compNo, acctNo, name])
    else:
        taxStatus = parsedLine[org_hdr.index("Status")]
        outputLine = createOutputLine(org_hdr, parsedLine, addrStartList, taxStatus)
        
        if taxStatus == "":
            record_dict["eligible"].append(outputLine)
        else:
            record_dict["ineligible"].append(outputLine)

   
def createOutputLine(org_hdr, parsedLine, addrStartList, taxStatus):

    # Extract Name and Address Lines from record 
    firstaddr = org_hdr.index("Account Registration Line 1")
    lastaddr = org_hdr.index("Account Registration Line 6")+1
    addrfields = [i for i in parsedLine[firstaddr:lastaddr] if i.upper() not in ["", "NULL"]]

    ''' Use Address Start Point to split Name and Address into 
    Registration Lines, Address Lines and City-State-Zip Line '''
    
    addrStartPoint = parsedLine[org_hdr.index("Address Line Starting Point")]
    
    if addrStartPoint not in addrStartList:
        addrStartList.append(addrStartPoint)
    
    addrLineStartIdx = int(addrStartPoint)-1 
    
    compNo = parsedLine[org_hdr.index("Company Number")]
    acctNo = parsedLine[org_hdr.index("Account Number")]

    company_acct_no = "{}{}".format(compNo, acctNo)
    RegistrationLines = addrfields[0:addrLineStartIdx]
    regLineSpaces = (6-len(RegistrationLines)) * [""]
    AddrLines = addrfields[addrLineStartIdx:-1]
    addrSpaces = (4-len(AddrLines)) * [""]
    citystatezip = addrfields[-1]

    outputLine = [company_acct_no] + RegistrationLines + regLineSpaces + \
        AddrLines + addrSpaces + [citystatezip, addrStartPoint, taxStatus]

    return outputLine


def writeFilteredFile(outputDir, record_dict):
    with open(os.path.join(outputDir, "FilteredRecords.txt"),'wb') as f:
        csvOutFiltered = csv.writer(f, quoting=csv.QUOTE_ALL)
        csvOutFiltered.writerow(["Record Sequence","Company Number","Account Number","Name"])

        for line in record_dict["filtered"]:
            csvOutFiltered.writerow(line)


def writeNCOARecordstoFile(outputDir, record_dict):
    header = ["CompAcctNo", "FullName", "Company", 
              "Company2", "Company3", "Company4", "Company5",
              "Delivery Address", "Alternate 1 Address", 
              "Alternate 2 Address", "Alternate 3 Address",
              "City-State-Zip", "AddrStartPoint", "Status"]

    with open(os.path.join(outputDir, "ParsedData.csv"),'wb') as p:
        csvOut = csv.writer(p, quoting=csv.QUOTE_ALL)
        csvOut.writerow(header)
        
        for line in record_dict["eligible"] + record_dict["ineligible"]:
            csvOut.writerow(line)


def createCounts(outputDir, filename, record_dict, comp_filter_list, addrStartList):
    with open(os.path.join(outputDir, "COUNTS.txt"),'wb') as c:

        eligible = len(record_dict["eligible"])
        ineligible = len(record_dict["ineligible"])
        filtered = len(record_dict["filtered"])
        to_process = eligible + ineligible
        total = to_process + filtered

        reportStr = '\r\n'.join([
        "Filename: {}".format(filename),
        "Total Record count: {}".format(total),
        "",
        "Records to process: {}".format(to_process),
        "Eligible: {}".format(eligible),
        "Ineligible: {}".format(ineligible),
        "",
        "Records filtered: {}".format(filtered),
        "Company Numbers filtered: {}".format(comp_filter_list),
        "",
        "\r\nAddress Start Points: {}".format(", ".join(sorted(addrStartList)))
        ])

        print reportStr
        c.write(reportStr)


    
   
if __name__ == "__main__":
    sys.exit(main())
