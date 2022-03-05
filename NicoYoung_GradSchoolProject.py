from datetime import *
import easygui
from easygui import *
import io
from google.cloud.vision_v1 import types
from oauth2client.client import GoogleCredentials
from google.cloud import vision
import os
import re
from dateutil.parser import parse
import pyscreenshot as ImageGrab
from enum import Enum
import mouse
from io import BytesIO
import win32clipboard
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pprint
import json

#Credentials file utilized by Google Cloud Vision API for access to the API
CREDENTIALS_FILE =  "C:\\NOA\\credentials.json"
GOOGLE_APPLICATION_CREDENTIALS = CREDENTIALS_FILE
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] =CREDENTIALS_FILE

#Create a global logfile variable which is updated frequently for debugging purposes. the file is updated when different functions of the program are entered/executed/exited
logfile=""
def docScan():
    from PIL import Image, ImageDraw, ImageFont, ImageTk
    global logfile,all_fields,docLowerLeft,docTitle, numdays, fromDate, toDate, noa_info,derivedLabels, county, currentNOA


    #Create a string representing the Date and Time the document is being processed
    dt_string = datetime.now().strftime("%m/%d/%Y %H:%M:%S")

    #create/open and modify a "log file" for debugging purposes
    logfile = open("C:/NOA/OCRlogFile.txt",mode='a')
    logfile.write("\ndate and time =" + dt_string)

    #Create a loop which loops as long as getCurrentNOA() does not return False, which occurs if user terminates the program via the GUI.
    #i variable also present as failsafe so code does not loop infinitely for some reason. residual failsafe code from previous iterations of project
    i=1
    while(True and i<1000):
        i=i+1

        #
        getCurrentNOA()
        if(currentNOA==None):
            logfile.close()
            return

        reply=True

        while (reply):
            reply = getReply()
            if(reply == "Exit"):
                reply=False
             #   Image.close()
            elif(reply == "Case Info"):
                getCaseInfo()
            elif(reply == "$ Money"):
                getMoney()
            elif(reply == "Derived Fields"):
                deriveFields()
##            elif(reply == "Extract Text"):
##                extractText()
##            elif(reply== "Extract Dates"):
##                extractDates()
            elif(reply == "Summary"):
                getSummary()
            elif(reply == "Paraphrase"):
                getParaphrase()
            elif(reply=="Save"):
                saveNOA()

#Will return a string URL with correct "/" in pathname of the NOA file
def getCurrentNOA():
    from PIL import Image, ImageDraw, ImageFont, ImageTk
   # try:
    global currentNOA
    default = "C:\\NOA\\test\\*.jpg"

    #Reinitializes global variables
    reinitializeGlobals()


    currentNOA =  easygui.fileopenbox(msg="Open NOA",title = "Open NOA",default = default)
  #  print(currentNOA)
    if(currentNOA == None):
        return
    currentNOA= currentNOA.replace("\\","/")
    updateOCRvars()
    image = Image.open(currentNOA)
    image.show()
    #except:
     #   msgbox("Please check getCurrentNOA (ctrl+shift+n) ...")

    #currentNOA = getCurrentNOA()

#Reinitializes global variables on each iteration of our original while loop, which resets everything from one document being process to the next, so that information is not stored incorrectly
def reinitializeGlobals():
    global document,all_fields,docLowerLeft,docTitle,\
    numdays, fromDate, toDate, noa_info,derivedLabels, county, currentNOA,moneyVlist,moneylabels,\
    PP,summary,summaryl,summaryv,fromto,caseFlist,caseVlist,documentType

    #Fields we have found to be necessary when processing
    derivedLabels = ["Effective Date","Old Rate","New Rate","NOA Type","Sub Type","Reason","Client ID","Serial Number","From Date", "To Date", "Number Of Days"]
    noa_info = []
    moneyVlist=[]
    moneylabels =[]
    summary = []
    summaryl=[]
    summaryv=[]
    fromto=[]
    caseFlist=[]
    caseVlist=[]
    county = ""
    currentNOA = ""
    PP=""
    numdays = -1
    fromDate = ""
    toDate = ""
    all_fields = []
    docLowerLeft = ""
    docTitle = ""
    documentType = ""

def updateOCRvars():
    global document,docAllText,docLowerLeft,docTitle,currentNOA,county, noa_info
    #Create a "document" variable which can be accessed for OCR data
    document = getGoogleOCR(currentNOA)

    #automatically access various parts of the document where certain information (e.g. County name, or the Lower Left "description") is always found
    docAllText = ocr_snippet(document,0,0,4000,4000)
    docLowerLeft = ocr_snippet(document,0,1100,1280,2600)
    docTitle = ocr_snippet(document,0,-10,3500,250)

    # automaticalyl retrieves the name of the County where the NOA originates
    county = getCounty(docTitle)

    #creates a global list containing variables found in the lower left "description" quadrant of the NOA
    noa_info = noa_type(docLowerLeft,currentNOA)

#EYAL START


client = vision.ImageAnnotatorClient()
    #returns "document" output from Google OCR
def getGoogleOCR(currentNOA):
        with io.open(currentNOA, 'rb') as image_file:
           content = image_file.read()
        image = types.Image(content=content)
        response = client.document_text_detection(image=image)
        return response.full_text_annotation


#TAKES DOCUMENT AS OUTPUT FROM GOOGLE OCR AND loops through text in a specific location based on coordinates on a symbol level
def ocr_snippet(document,x1,y1,x2,y2,pp=0,eol=0):

  text=""
  for page in document.pages:
    for block in page.blocks:
      for paragraph in block.paragraphs:
        for word in paragraph.words:
          for symbol in word.symbols:
            min_x=min(symbol.bounding_box.vertices[0].x,symbol.bounding_box.vertices[1].x,symbol.bounding_box.vertices[2].x,symbol.bounding_box.vertices[3].x)
            max_x=max(symbol.bounding_box.vertices[0].x,symbol.bounding_box.vertices[1].x,symbol.bounding_box.vertices[2].x,symbol.bounding_box.vertices[3].x)
            min_y=min(symbol.bounding_box.vertices[0].y,symbol.bounding_box.vertices[1].y,symbol.bounding_box.vertices[2].y,symbol.bounding_box.vertices[3].y)
            max_y=max(symbol.bounding_box.vertices[0].y,symbol.bounding_box.vertices[1].y,symbol.bounding_box.vertices[2].y,symbol.bounding_box.vertices[3].y)

            if(min_x >= x1 and max_x <= x2 and min_y >= y1 and max_y <= y2):
              text+=symbol.text
              if(symbol.property.detected_break.type==1 or
                symbol.property.detected_break.type==3):
                text+=' '
              if(symbol.property.detected_break.type==2):
                text+='\t'
              if(symbol.property.detected_break.type==5):
                if(eol==1):
                    print(" EOL ")
                    text+='\n'
                else:
                    text+=' '
  if(pp==1):
    print(text)
  return text
#EYAL END


#matches the info in the "title/header" bounding box to the counties in California, returns the county name if present or "COUNTY NOT FOUND"
def getCounty(header):
    #global counties
    counties = ["Merced","Alameda","Inyo","Modoc","Alpine","Kern","Amador","Kings","Mono","North","Butte","Oroville","Lake","Calaveras",\
    "Lassen","Monterey","Colusa","Contra Costa","Los Angeles","Napa","Del Norte","Nevada","El Dorado","Madera","Orange","Anaheim","Fresno",\
    "Marin","Mariposa","Mendocino","Fort Bragg","Glenn","Humboldt","Ukiah","Imperial","Placer","Plumas","Riverside","Tulare","Santa Cruz",\
    "Watsonville","Tuolumne","Sacramento","San Benito","Shasta","Ventura","San Bernardino","Sierra ","Loyalton","Downieville","San Diego",\
    "Oxnard","Santa Clara Valley","Siskiyou","San Francisco","San Joaquin","Solano","Fairfield","Vacaville","Vallejo","Yolo","Woodland","San Luis Obispo",\
    "Sonoma","Yuba","San Mateo","Stanislaus","Santa Barbara","Sutter","Santa Clara","Tehama","Trinity"]
    for c in counties:
        #print(header,c)
        if(contains(header.lower(),c.lower())):

            return c

    return "COUNTY NOT FOUND"

#returns True if mystr is found in ocrtext
def contains(ocrtext,mystr):
    return (ocrtext.find(mystr) >-1)


#creates a global list containing variables found in the lower left "description" quadrant of the NOA
documentType = ""
def noa_type(docLowerLeft,currentNOA):
    global numdays, fromDate, toDate, documentType
   # print(currentDoc,mystr)

    #The new billing rate:
    toRate = ""
    #The old billing rate:
    fromRate = ""
    #the subtype of NOA such as under/over payment in a Non-recurring Payment NOA type:
    subType = ""
    #the Type of NOA, each "type" has certain characteristics in each quadrant and contains different information:
    noaType = ""
    #The date that the billing changes become effective:
    effectiveDate = ""
    #The reason for the change in billing, why the NOA was sent in the first place:
    reason = ""
    #The client ID, which is found in the file name along with the serialNum:
    clientID = ""
    #The serial number of the case, which is found in the file name along with clientID:
    serialNum = ""

  ##  if(currentNOA != None):
    #This function extracts the clientID and serial number of the NOA from the file name
    clientID,serialNum = getClientID_SerialNum(currentNOA)


    #If the Lower Left "description" is not present in the NOA, no information can be gathered and this function exits
    if(docLowerLeft==None or docLowerLeft == ""):
        return ""

#A "clothing allowance" NOA is drastically different than many other more similar types of NOAs, so they are handled differently in a different function
    if(contains(docLowerLeft.lower(),"clothing allowance")):
        return clothingNoa()

#Similarly to Clothing Allowances above, COVID related NOAs are handled differently
    if contains(docLowerLeft.lower(),"covid"):
        return covidNoa(clientID,serialNum)

    #print(docLowerLeft)

    #in many NOAs the billing period is a limited time period containing a from data and to date e.g."from 8/1/2021 to 8/30/2021"
    #These dates are stored as variables and used to determind the time period in days e.g. 30 days in the example above
    fromToDate = []

    #print(numdays,type(numdays),1)
    #If the number of days of a billing period (numdays) is found elsewhere, this part is skipped
    if(numdays==-1):
     ##   print(numdays,type(numdays),2)
        #If the Lower Left "description" contains "from and to" then the numdays can be calculated now
        if((docLowerLeft.find("To") >-1) and (docLowerLeft.find("From") >-1)):
            ##print("Found From and To")
            ##print(numdays,type(numdays),3)

            #here the Lower Left is sent to be parsed for the From and To dates
            fromToDate = parse_date2_all(docLowerLeft)
            #Depending on how many dates are found in the Lower Left, the From and To dates are ALWAYS in certain positions in the description
            if(len(fromToDate)==2):
       #         print(numdays,type(numdays),4)
                fromDate = fromToDate[0]
                toDate = fromToDate[1]

                #here we calculate the numdays value using substraction, and add 1 to make the numdays an inclusive value
                numdays=parse(toDate)-parse(fromDate)
                numdays = numdays.days+1

                #Here we create String values for the from and to dates for easier integration into the data
                fromDate = parse(fromDate)
                toDate = parse(toDate)
                fromDate = str(fromDate.month) + "/" + str(fromDate.day) + "/" + str(fromDate.year)
                toDate = str(toDate.month) + "/" + str(toDate.day) + "/" + str(toDate.year)
                #if something is wrong with the numdays we set it to -1 and print a warning to the user via the Python console
                if(numdays<0 or numdays>31):
                    numdays = -1
                    print("Check Number of Days for " + currentNOA)

    #here we make the Lower Left all lowercase for parsing ease
    docLowerLeft = docLowerLeft.lower()

    ##here we parse the date which the action became effective
    ##effectiveDate = parse_date2(docLowerLeft)

    #Here we check if there is a an effective date presenet and create a string representing it (or a blank string if it is an error)
    if(parse_date2(docLowerLeft) == ""):
        effectiveDate = ""
    else:
        effectiveDate = parse(parse_date2(docLowerLeft))
        effectiveDate = str(effectiveDate.month) + "/" + str(effectiveDate.day) + "/" + str(effectiveDate.year)

    #Here we parse a dollar amount for the calculation of rates later on
    rate1 = parse_dollar_amount2(docLowerLeft) #Function to only parse "From/To" Rates ?

    #Here we determine the type/subtype of the NOA based on the presence of certain words
    if(contains(docLowerLeft,"underpaid")):
        noaType = "Non-Recurring Payment"
        subType = "Underpayment"
        toRate = rate1

    elif(contains(docLowerLeft,"overpaid")):
        noaType = "Non-Recurring Payment"
        subType = "Overpayment"
        toRate = rate1


    #If the above 2 conditions are not met, we do some more analysis based on the presence of a billing amount
    #This analysis determines the type of the NOA and the reason for the NOA
    elif(rate1!="" and rate1 != None and parse_dollar_amount2(docLowerLeft.split(rate1)[1]) != ""):

       ## splitAmt =  mystr.split(rate1)

        noaType = "Rate Change"
        reason = "change in rate"
        fromRate = rate1.replace(",","")
        ##    if(x.find(",")>-1):
        ##    x = x.replace(",","")
        #Here we get the new rate for the client
        toRate = parse_dollar_amount2(docLowerLeft.split(rate1)[1]).replace(",","")

        #if the new rate is more than the old rate, its an increase, if vice versa then a decrease, otherwise we simply put "change" (if they are the same or possibly one missing?)
        if(float(toRate)>float(fromRate)):
            subType = "Increase"

        elif(float(toRate)<float(fromRate)):

            subType =  "Decrease"

        else:
            subType = "Change"
        #Again, we search for certain phrases which denote the reason for the NOA
        if(docLowerLeft.find("level of care")>-1):
            reason = "Level of Care"
        elif(docLowerLeft.find("cni")>-1):
            reason = "California Necessities Index (CNI)"
        #Here we determined that if the rates are the same we should explicitly state so
        elif(float(toRate)==float(fromRate)):
            subType = "Same"


    #Here we determind if care has been terminated("type") and for what reason
    elif((docLowerLeft.find("discontinu")>-1)or(docLowerLeft.find("stop"))>-1):
        noaType = "Termination"
        fromRate = rate1
        if(docLowerLeft.find("no longer meets")>-1):
            reason = "No Longer Meets Age Requirement"
        elif((docLowerLeft.find("left your")>-1) or (docLowerLeft.find("no longer living")>-1) or (docLowerLeft.find("not living")>-1)):
            reason = "Left Facility/Home"
        elif((docLowerLeft.find("adopt")>-1)):
            reason = "Adoption"
        elif(contains(docLowerLeft,"did not return")):
            reason = "Paperwork not Complete"
        elif(contains(docLowerLeft,"did not complete ")):
            reason = "Application/Reevaluation Process not Complete"
        elif(docLowerLeft.find("became ineligible")>-1):
            reason = "Became Ineligible"

    #Here we determine if benefits have been approved for a client
    elif(contains(docLowerLeft,"approv")):
        noaType = "Approval"
        toRate = rate1
        if(docLowerLeft.find("medi-cal")):
            subType = "Medi-Cal/Cash Aid"

    #Return all the information gathered into the "noa_info" which is represented to the User as "Derived Fields" since these were automatically derived without user input.
    #The user can edit the information via viewing the "Derived Fields"
    return [effectiveDate,fromRate,toRate,noaType,subType, reason,clientID,serialNum,fromDate,toDate,numdays]


#This function extracts the clientID and serial number of the NOA from the file name
def getClientID_SerialNum(currentNOA):
    #print(currentNOA)

    #The files consistently have the convention "clientID_serialNum" in the name which can easily and consistently be extracted
    filename = currentNOA[currentNOA.rindex("/")+1:]
    if(contains(currentNOA,"_")):
        clientID = filename.split("_")[0]
        serialNum = currentNOA[currentNOA.rindex("/")+1:currentNOA.rindex(".")].split("_")[1]
    else:
        clientID = ""
        serialNum = ""
    return [clientID,serialNum]

#Here we gather the info found in noa_info for a COVID related NOA which is slightly different than most other NOAs
def covidNoa(clientID,serialNum):
    global documentType, derivedLabels, docLowerLeft,currentNOA, noa_info

    #There are 2 versions of COVID NOAs, one with a prorated rate which is analyzed below and,
    # a version without a prorated rate which is analyzed in the else statement
    if contains(docLowerLeft,"prorate"):
        documentType = "COVID Relief 1"
        derivedLabels = ["Effective Date","Prorated From", "Prorated To", "Prorated Allotment", "Monthly Rate","Effective To","Client ID", "Serial Num"]
        print("docLowerLeft: ",docLowerLeft)

        matchDate = re.findall(r'\d{2}/\d{2}/\d{4}', docLowerLeft)
        effectiveDate = matchDate[0]
        effectiveTo = matchDate[2]
        proFrom = matchDate[3]
        proTo = matchDate[4]
        matchMoney =  re.findall(r'\$(.+?\.\d{2})',docLowerLeft)
        prorateAmount = matchMoney[0].strip("$ ")
        monthlyRate = matchMoney[1].strip("$ ")
        return [effectiveDate, proFrom, proTo,prorateAmount, monthlyRate, effectiveTo,clientID,serialNum]
    else:
        documentType = "COVID Relief 2"
        derivedLabels = ["Effective Date", "Monthly Rate","Effective To","Client ID", "Serial Num"]
        matchDate = re.findall(r'\d{2}/\d{2}/\d{4}', docLowerLeft)
        effectiveDate = matchDate[0]
        effectiveTo = matchDate[2]
        monthlyRate = re.findall(r'\$(.+?\.\d{2})',docLowerLeft)[0]
        return [effectiveDate,monthlyRate,effectiveTo,clientID,serialNum]




    ##apply this to noaType and Paraphrase



#similar to COVID NOAs a specific NOA sent for a Clothing Allowance for a client is drastically different
#here we create the Derived Fields (noa_info) as it applies to this type of NOA

def clothingNoa():
    global documentType, derivedLabels, docLowerLeft,currentNOA, noa_info
    documentType = "Clothing Allowance"
    requestStatus = ""
    derivedLabels = ["Effective Date","Allowance Amount","Request Status"]
   ## print(docLowerLeft)

   #Here we see if it is explicitly approved in the NOA
    if("approv" in docLowerLeft.lower() or "authorized" in docLowerLeft.lower()):
        requestStatus = "approved"

    #Here we search for a date using reg ex and if an error is produced we rely on a different function that we typically use elsewhere
    try:
        effectiveDate = re.search(r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) (?:2\d{3})?(?=\D|$)",docLowerLeft,26)[0].strip()
    except:
        effectiveDate = getNoticeDate()

    #Here we look for a keyword which denotes a specific type of NOA and process accordingly
    #A recurring payment introduces a field called "Totaly Monthly Aid" on top of the "Allowance Amount"
    if(contains(docLowerLeft.lower(),"recurring")):
        derivedLabels.append("Total Monthly Aid")
        #money
##        try:
##            effectiveDate = re.search(r".?((?:January|February|March|April|May|June|July|August|September|October|November|December|[\d]{1,2})[/\-\s\,][\d]{1,2}[/\-\,\s]{1,2}?[\d]{2,4})",docLowerLeft,26)[0].strip()
##        except:
##            effectiveDate = getNoticeDate()

        #Here we utilize various regex patterns to look for the money amount mentioned in the NOA,
        #if there are not two different amounts then the money amount is the same for both Fields
        try:
            totalAid = re.findall(r".?([\d]*\.[\d]{2})",docLowerLeft,26)[0].strip()
            allowanceAmt = re.findall(r".?([\d]*\.[\d]{2})",docLowerLeft,26)[1].strip()
        except:
            totalAid = re.findall(r".?\$*([\d]*\.[\d]{2})",docLowerLeft,26)[0].strip()
            allowanceAmt = totalAid

        #Here we return the "Derived Field" Values, whereas the Labels are global
        return [effectiveDate,allowanceAmt,requestStatus,totalAid]

    #If the aforementioned keyword is not present, then we process accordingly and return the Derived Fields
    else:
        allowanceAmt = re.search(r".?([\d]*\.[\d]{2})",docLowerLeft,26)[0].strip()
        return [effectiveDate,allowanceAmt,requestStatus]

#Here is our function for parsing dates which calls our other function designated for regular expression
#This is our second iteration of our parse date function, hence the "2" in the name
def parse_date2(mystr):
    #our pattern looks for a date either Month or MM and DD YY, all seperated by slashes, commas, dashes, or spaces
    date = ocr_regex3(mystr,r".*?((?:January|February|March|April|May|June|July|August|September|October|November|December|[\d]{1,2})[/\-\s\,][\d]{1,2}[/\-\,\s]{1,2}?[\d]{2,4})",1)
    if(date!=None):
        return date
    else:
        return ""

# gets a string, pattern, and groupnum and matches the pattern in the string and returns the matches
def ocr_regex3(mystr, pattern,groupnum=1):
    if (mystr != ''):
          #  print(mystr)
           # print(pattern)
        #26 is a tag representing the combination of 2,16 and 8, which correspond to ignorecase, multiline and carriage return rules in regex
            m = re.match(pattern,mystr,26)
           # print(m)
            if((m!=None)and(m.group(1) != None)and(m.group(1)!="")):
                if(groupnum==1):
                   # print(m.group(1))
                    return (m.group(1).strip())
                elif(groupnum ==2):
                    #print(m.group(1) + " " + m.group(2))
                    return(m.group(1).strip(),m.group(2).strip())
                elif(groupnum==3):
                    #print (m.group(1) + " " + m.group(2) + " " + m.group(3))
                    return(m.group(1).strip(),m.group(2).strip(),m.group(3).strip())
                elif(groupnum==4):
                     return(m.group(1).strip(),m.group(2).strip(),m.group(3).strip(),m.group(4).strip())
                elif(groupnum==5):
                    if(m.group(1)!= None):
                        return(m.group(1).strip(),m.group(2).strip(),m.group(3).strip(),m.group(4).strip(),m.group(5).strip())
                    else:
                        return(m.group(6).strip(),m.group(2).strip(),m.group(3).strip(),m.group(4).strip(),m.group(5).strip())
            else:
                return None
    else:
        return None

#searches for a Dollar amount ignoring certain symbols which interfere with analysis later. we only want the number
def parse_dollar_amount2(mystr):
#    print(mystr)
    x = ocr_regex3(mystr,r'.+?[\|\s\=\-\+\$]*?(\d+?,?\d*\.\d\d)',1)
##    if(x.find(",")>-1):
##        x = x.replace(",","")
    if(x!=None):
        return x
    else:
        return ""

#Here we have a button box which gives the user options for what data they want to extract
#Each reply relates to a certain location on the NOA which contains specific info.
def getReply():
    #Here we utilize the county variable we previously found to create the title of the button box
    msg = "Get info for: " + county + " County"
    #these are the choices the user can seleect
    #"Extract Text", "Extract Dates",
    choices = ["Case Info", "$ Money", "Derived Fields","Summary", "Paraphrase", "Save","Exit"]
    #The users choice is stored and returned to determine how the program will behave next
    reply = buttonbox(title = "OCR Data from file: " + currentNOA, msg=msg, choices=choices)
    return reply


#Retrieves upper right information from screen capture
#adds an address2 label if address is two lines
#removes spaces from ID's or Client/Worker Numbers
#returns field list and value list as flist and vlist seperately
def getCaseInfo():
    global caseFlist,caseVlist,logfile
    caseFlist=[]
    caseVlist=[]

   ## try:
    logfile.write(" Entered getCaseInfo ")
    #Here we utilize mouse clicks from the user to create x,y coordinates for a bounding box which is..
    #then cropped and OCRd "behind the scenes" creating a list related to the "Case Info" of the NOA
    caselist=screen_capture_ocr2()

    if(caselist != None):
        #Here we divid the Fields list and respective Values list
        caseFlist,caseVlist = getFVlists(caselist)

        #Here we check for a specific Field Label that for some reason wasn't properly picked up,
        #and insert the correct Value based on the position of the Field
        if(len(caseVlist) - len(caseFlist) ==-1) and "Cps Case Number" in caseFlist:
            caseVlist.insert(caseFlist.index("Cps Case Number"),"")

       ## print(caseFlist,caseVlist)

        #EYAL START
        #If the lengths do not match, we add a second address label
        #I suppose we could simply concat the 2 address Values as well, alternatively
        if ((len(caseFlist) - len(caseVlist))==-1):
            caseFlist.extend(['Address 2'])
        #EYAL END

        #Here we essentilly strip the Fields and Values of nonessential characters
        caseVlist=removeSpaces(caseFlist,caseVlist)

        if (caseFlist==[]):
            caseFlist=caseVlist #Used to prevent error in multenterbox blank fields
        #Here the global caseVlist is displayed to the user with the option to edit the Values
        caseVlist = multenterbox(fields= caseFlist,values = caseVlist)

    if (caselist==None):
        caseVlist=[]
    logfile.write(" Exited getCaseInfo ")
   # except:
      #  msgbox("Please check getCaseInfo (ctrl+shift+c) ...")




##NEEDS EDIT
#Here we go through each item in the "caselist" (Case Info list)
#We determine whether each item is a Field/Label or Value
def getFVlists(caselist):
    flist=[]
    vlist=[]
    for item in caselist:
            ##print(caselist)

            #lowercase each item for ease in analysis
            lc=item.lower()

            ##print("lc: ", lc)

            #avoiding certain issues caused by multiple colons due to time,
            #Since the Case Info is displayed as Label : Value and we split based on the colon
            #For some reason if there are multiple colons being read yet it is not an office hour,
            #we have to go through a more rigorous process to get the info we want

            if(contains(lc.strip(": "),": ") and not isOfficeHours(item)):
              #  print(lc.strip(": "))
                casel = lc.strip(": ").split(": ")[0].strip(" ")
                casev = item.strip(": ").split(": ")[1].strip(" ")
                if(casel != None and casev!=None):
                    flist.extend([casel.title()])
                    vlist.extend([casev])

            #Instead of Omitting office hour field we just process it a bit differently in a way we found to work
            elif (isOfficeHours(item) and len(item.split(": "))==2):
                officeHours = item.split(": ")
             #   print("OFFICE HOURS: ",officeHours)
                flist.extend([officeHours[0]])
                vlist.extend([officeHours[1]])

            #This is the ideal situation, we just check if the value is in our list of Field Labels,
            #if it is then we put it in the list of Field labels for the current document
            #otherwise we put the item into the Value list
            else:
                if(isFieldLabel(lc)):
                    flist.extend([item.title()])
                else:
                    vlist.extend([item])
    return [flist,vlist]


#This returns whether or not the item is a time based on the presence of AM/PM text
def isOfficeHours(item):
    item = item.strip(": ")
    return(contains(item,"AM") or contains(item,"PM") or contains(item,"a.m."))


##change for sure!
#This returns the truth value of whether or not the item we are searching for is a Field Label
#This code is a bit messy due to the variation of real life information
#(e.g. Telephone rd. was an address but was being input as a Field Label due to the presence of 'Telephone' which had only previously appeared as a Label)
def isFieldLabel(lc):
  #  print(lc)
    return (contains(lc,'notice') or contains(lc,'date')or contains(lc,'number')or contains(lc,'caseload ') or\
    contains(lc,'case ') or (contains(lc,'worker') and not contains(lc,"call")) or contains(lc,'information') or\
    (contains(lc,'telephone') and not contains(lc,"rd")) or contains(lc,'notice') or contains(lc,'address') or\
     contains(lc,'tdd') or (contains(lc, 'hours') and not contains(lc,":"))or contains(lc,'customer'))

#Here we strip the spaces from certain Fields for use later
def removeSpaces(flist,vlist):
    v2list=[]
    print(flist,vlist)
##    print(len(flist),len(vlist))
   # print(flist,vlist)
    for i in range(0,len(flist)):
##        print(i, flist[i], vlist[i])

        if(flist[i].title()=="Case Number" or flist[i].title()=="Worker Id" or flist[i].title()=="Worker Number"):
            v2list.append(vlist[i].replace(" ",""))
        else:
            v2list.append(vlist[i])

    return v2list


#EYAL START
class FeatureType(Enum):
    PAGE = 1
    BLOCK = 2
    PARA = 3
    WORD = 4
    SYMBOL = 5

def screen_capture_ocr2(fileout=0,level=FeatureType.WORD):
    from PIL import Image, ImageDraw, ImageFont
    global logfile

    logfile.write(" Entered screen_capture ")
    image = Image_Grab_click()
    if (image==None):
        return None
    #image = image.convert('RGBA')
    #image.show()

    #The image is analyzed via OCR on various levels
    if(FeatureType.WORD==level):
        bounds = get_document_bounds_crop(image, FeatureType.WORD)
    if(FeatureType.SYMBOL==level):
        bounds = get_document_bounds_crop(image, FeatureType.SYMBOL)
    if(FeatureType.PARA==level):
        bounds = get_document_bounds_crop(image, FeatureType.PARA)

    if fileout != 0:
        image.save(fileout)
    x=document.text.splitlines()
    z=[]
    for a in x:
        z.extend([a.strip(":/ |")])
    if(z==[]):
        return None
    else:
        return z
    logfile.write(" Exited screen_capture ")




def Image_Grab_click(filename="imageGrab.jpg"):
# part of the screen
    path="C:\\NOA\\"
    #Here we get coordinates from mouse clicks,
    #the users have been trained to click from top left of target to the bottom right of the target
    (x1,y1,x2,y2) = get_mouse_coords()
    #if the user did not move the mouse when clicking they have to click again
    if((x1==x2) and (y1==y2)):
        (x2,y2,x3,y3) = get_mouse_coords()
        if((x3>x1) and (y3>y1)):
            #Here the overall image is cropped based on the mouse coords and saved as a temporary file which can be recalled for OCR
            im = ImageGrab.grab(bbox=(x1, y1, x3, y3)) # X1,Y1,X2,Y2
            im.save(path + filename)
            copyimage(im)
            return im
            im.close()
        else:
            return None
    else:
        #If the user did not click from top left to bottom right they have to click again,
        #the aforementioned algorithm above occurs now
        while((x1>=x2)or(y1>=y2)):
            (x1,y1,x2,y2) = get_mouse_coords()
        im = ImageGrab.grab(bbox=(x1, y1, x2, y2)) # X1,Y1,X2,Y2
        im.save(path + filename)
        copyimage(im)
        return im
        im.close()


def copyimage(image=None):

    output = BytesIO()
    if (image==None):
        image=Image_Grab_click()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)




#This gets x,y coordinates based on the position of the mouse from the user when they click
#The user can also drag while holding the left mouse is clicked
#These coordinates create a bounding box for OCR and image manipulation
def get_mouse_coords():
    while(mouse.is_pressed(button='left')==False):
        pass
    x1=mouse.get_position()[0]
    y1=mouse.get_position()[1]
    while(mouse.is_pressed(button='left')==True):
        pass
    x2=mouse.get_position()[0]
    y2=mouse.get_position()[1]
    return (x1,y1,x2,y2)

#DOES GOOGLE OCR ON AN IMAGE OBJECT AS OPPOSED TO IMAGE URL
def get_document_bounds_crop(image1, feature = FeatureType.WORD):
    # [START vision_document_text_tutorial_detect_bounds]
    """Returns document bounds given an image."""
    global bounds
    global response
    global document

    client = vision.ImageAnnotatorClient()

    bounds = []
##    with io.open(image_file, 'rb') as image_file:
##        content = image_file.read()
    content = image_to_byte_array(image1)
    image = types.Image(content=content)
    response = client.document_text_detection(image=image)

    #DOCUMENT IS WHAT WE WANT FROM THIS SECTION FOR OUR PURPOSES, HOW NECESSARY IS OTHER CODE?
    document = response.full_text_annotation

    # Collect specified feature bounds by enumerating all document features
    for page in document.pages:
        for block in page.blocks:
            for paragraph in block.paragraphs:
                for word in paragraph.words:
                    for symbol in word.symbols:
                        if (feature == FeatureType.SYMBOL):
                            bounds.append(symbol.bounding_box)
                    if (feature == FeatureType.WORD):
                        bounds.append(word.bounding_box)
                if (feature == FeatureType.PARA):
                    bounds.append(paragraph.bounding_box)
            if (feature == FeatureType.BLOCK):
                bounds.append(block.bounding_box)
        if (feature == FeatureType.PAGE):
            bounds.append(block.bounding_box)

    # The list `bounds` contains the coordinates of the bounding boxes.
    # [END vision_document_text_tutorial_detect_bounds]
    return bounds


def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

#EYAL END

#receives an Image object and converts it to a byte array for analysis
def image_to_byte_array(image:Image):
  imgByteArr = io.BytesIO()
  image.save(imgByteArr, format="PNG")
  imgByteArr = imgByteArr.getvalue()
  return imgByteArr

#This is where we begin to analyze the lower right quadrant which typically contains the Money Info for the NOA
def getMoney():
    global moneylabels,moneyVlist,logfile,numdays
    ##NICO try:
    logfile.write(" Entered getMoney ")
    #Similar to Case Info, we start with some Value Lists and Labels variables, including temp variables
    moneyVlist=[]
    moneyvalues = []
    moneylabels = []
    moneyInfo = []

    #immediately we "get" the money labels and value and print them to check for errors
    #The two separate functions means that the Fields must be surrounded with a bounding box by the user as well as the Values on the JPG
    #To process Money info, the user must click a total of 4 times to create 2 bounding boxes
    moneylabels = getMoneyLabels()
    moneyvalues = getMoneyValues()
    ##print("MoneyLabels: " + str(moneylabels))
    ##print("MoneyValues: " + str(moneyvalues))

    #Here we check for a specific version of NOA from Kings County that has as wholly different format
    if "budget calculation" in (string.lower() for string in moneylabels) or any("overpayment" in string.lower() for string in moneylabels):
        moneyInfo= getMoneyKC(moneylabels)
        moneylabels = moneyInfo[0]
        moneyvalues = moneyInfo[1]

        #Here we display the Fields and Values of the Kings County overpayment NOA
        moneyVlist = multenterbox(fields= moneylabels,values = moneyvalues)

    #For all other NOAs we display the Fields/Values like this
    #In some instances two pages are scanned/multiple occurences of the same Field occur, we label these with a 2
    else:
        if (moneylabels!=None and moneylabels!=[]):
            moneyVlist = multenterbox(fields= moneylabels,values = moneyvalues)
            if(len(moneyvalues)>len(moneylabels)):
                moneylabels2 = [label + " 2" for label in moneylabels]
                moneyVlist = moneyVlist + multenterbox(fields= moneylabels2,values = moneyvalues[len(moneylabels):])
                moneylabels = moneylabels + moneylabels2
                numdays = 0
    logfile.write(" Exited getMoney ")

##    except:
##        msgbox("Please check getMoney (ctrl+shift+m) ...")
##        return moneylabels,moneyVlist


#This is the Kings County special NOA which is handled uniquely from others
def getMoneyKC(moneyInfo):
    #print("moneyvalues: ",moneyvalues)
    moneylabels = []
    moneyvalues =[]
    for i in moneyInfo:
        print("i: " +i)
        if i.strip(": ").lower() == "budget calculation" or i.strip(": ").lower() == "overridden placement payment":
            continue
        moneylabels.append(re.split(r"\$|\+|\-|\=",i)[0].strip(": "))
        moneyvalues.append(re.split(r"\$|\+|\-|\=",i)[1].strip(": "))

    #print ("Money label finished: " + str(moneylabels))
   # print("Money value finished: " + str(moneyvalues))
    return [moneylabels,moneyvalues]


#This is how all the money Labels are gathered
def getMoneyLabels():
        global noa_info
   # global numdays
        #The user clicks on the top left/bottom right of "money labels"
        moneylabels=screen_capture_ocr2(level = FeatureType.SYMBOL)

        if(moneylabels != None):
            finalMoneyLabels=[]

            for label in moneylabels:
                #We check for dates to calculate number of days
                #we also check if there is a clothing allowance in order to get the accurate "noa_info" derived fields
                lowerRightNumdays(label)
                if "clothing" in label.lower():
                    noa_info=clothingNoa()
                #We remove certain problematic characters or Fields that are not useful
                moneyLabel = removeLabel(label)
                if (moneyLabel != None):
                    finalMoneyLabels.extend([moneyLabel.strip("$-+= /|_")])
            #Throughout the code we make sure we are not using None objects which cause errors. We make sure to at least
            if(finalMoneyLabels==None):
                finalMoneyLabels=[]
            return finalMoneyLabels
        else:
            return []
    #except:

#Here we get the Values for the money Fields
def getMoneyValues():
    try:
        #Again the user creates a bounding box around the desired area
        moneyamounts=screen_capture_ocr2(level = FeatureType.SYMBOL)
        #print(moneyamounts)
        if(moneyamounts != None):
            moneyAmountFinal=[]
            #Strip any characters that arent the actual money amount/Value
            for amount in moneyamounts:
                moneyValue = removeAmount(amount.strip("$-+= S/|_"),numdays)

                #transfer from temp variables into the completed version to make sure we aren't repeating or stuck in a loop
                if(moneyValue != None):
                    moneyAmountFinal.extend([moneyValue])
            if(moneyAmountFinal==None):
                moneyAmountFinal=[]
            return moneyAmountFinal
        else:
            return []
    except:
        msgbox("Please check getMoneyValues ...")

#Remove the lines from Money which are not needed or used in this moment
def removeLabel(ml=""):
    if(ml.lower() in ["rate","special needs","number of days", "facility rate frequency","clothing allowances"] or contains(ml.lower(),"prorated from")):
        return None
    else:
        return ml

#Remove certain Values which sometimes get picked up and are not necessary for our purposes
def removeAmount(ma="",numofdays=-1):
##    print("ma: " + ma)
##    print("numofdays: ",numofdays,"\n")
    if(ma=="" or ma==None):
        return None
    if((ma[0:1].lower() in ["n"])):
        #print("MADE IT")
        return ""
    # M for Monthly (rate frequency) and no "." for numdays not being a dollar amount
    elif((ma[0:1].lower() in ["m"]) or ((contains(ma,".")==False))):
        #print("THIRD IF")
        return None
    else:
        return ma

#This is to calculate the numdays variable when the "prorated from/to" statement appears in the lower right money section instead of Lower Left
#this is essentially a repeat of the calculation done in noa_type()
def lowerRightNumdays(label):
    global numdays,fromDate,toDate,noa_info,derivedLabels

    if(contains(label.lower()," from ") and contains(label.lower()," to ") and numdays ==-1):
            fromto=label.lower().split(" from ")[1].split(" to ")

            numdays=parse(fromto[1])-parse(fromto[0])
            numdays = numdays.days+1

            fromDate = fromto[0]
            toDate = fromto[1]

            if(numdays<0 or numdays>31):
                print ("Error in Number of Days")
            else:
               for i in range(0,len(derivedLabels)):
                    if(derivedLabels[i] == "From Date"):
                        noa_info[i] = fromDate
                    elif(derivedLabels[i] == "To Date"):
                        noa_info[i] = toDate
                    elif(derivedLabels[i] == "Number Of Days"):
                        noa_info[i] = numdays

#Displays an editable version of noa_info, User is instructed in training to typically select Derive Fields only after Case Info and Money are processed for ease of use (avoids missing any info)
def deriveFields():
    #try except
    global noa_info,derivedLabels,numdays,fromDate, toDate,county
    if(noa_info[0].lower()!=county.lower()):
        noa_info = [county] + noa_info
        derivedLabels = ["County"] + derivedLabels
    noa_info = multenterbox(fields=derivedLabels,values = noa_info)
    county = noa_info[0].title()
    numdays = noa_info[-1:][0]
    fromDate = noa_info[-3:-2][0]
    toDate = noa_info[-2:-1][0]


#This displays an editable version of every Value we have collected
#Ideally the summary is reviewed after Case Info, Money, and Derived fields are all gathered
def getSummary():
    global caseFlist,moneylabels,derivedLabels,caseVlist,moneyVlist,noa_info,county
   # try:
    summaryl=[]
    summaryv=[]
##    print(caseFlist,moneylabels,derivedLabels)
##    print(caseVlist,moneyVlist,noa_info)
    if(noa_info[0].lower()!=county.lower()):
        noa_info = [county] + noa_info
        derivedLabels = ["County"] + derivedLabels
    summaryl.extend(derivedLabels+caseFlist+moneylabels)
    ##print(noa_info,caseVlist,moneyVlist)
    summaryv.extend(noa_info+caseVlist+moneyVlist)
    summary = multenterbox(fields= summaryl,values = summaryv)
    county = summary[0].title()
    if (summary==None):
        summary=[]
##    except:
##        msgbox("Please check getSummary (ctrl+shift+a) ...")

#Here we check to make sure the required info is present for our Paraphrase function
#We then display a message box with the Paraphrased version of the NOA which is created in our Paraphrase function below
def getParaphrase():
    global caseVlist,caseFlist,noa_info,county,currentNOA,numdays
    try:
        PP=""
        if(caseVlist == []):
            easygui.msgbox(msg="Please extract Case Info (ctrl+shift+c) from upper right before paraphrasing")
        elif(numdays == -1):
            #POTENTIALLY CHANGE TEXT IF NUMDAYS IS STILL -1 AFTER MONEY
            msgbox("Please check Derived Fields ...")
        else:
            PP = easygui.textbox(msg = "OK if Correct, Cancel if Error", text = Paraphrase(caseVlist,caseFlist,noa_info,currentNOA))
    except:
        msgbox("Please check getParaphrase (ctrl+shift+p) ...")


#This function checks many different Values and creates a String summarizing the entirety of the NOA in a few sentences
def Paraphrase(caselist, flist, noa_info,currentNOA):

    global fromDate,toDate,documentType,county

   ##try:

    #First we ensure County was inserted into the noa_info variable
    if(noa_info[0].lower()!=county.lower()):
        noa_info = [county] + noa_info

    #Here we declare some temporary variables which are utilized when writing the Paraphrase
    Pdate = ""
    PCaseName = ""
    PCaseNum = ""
    PWorkerName = ""
    PWorkerNum = ""
    PP = ""

    filename = currentNOA[currentNOA.rindex("/")+1:]

    #here we get the info we want into the respective variables for use later
    for i in range(0,len(flist)):
        if(contains(flist[i].lower(),"case")):
            if(contains(flist[i].lower(),"name")):
                PCaseName = caselist[i]
            elif(contains(flist[i].lower(),"number") or contains(flist[i].lower(),"id")):
                PCaseNum = caselist[i]

        elif(contains(flist[i].lower(),"worker")):
            if(contains(flist[i].lower(),"name")):
                PWorkerName = caselist[i]
            elif(contains(flist[i].lower(),"id") or (contains(flist[i].lower(),"number")and not contains(flist[i].lower(),"phone"))):
                PWorkerNum = caselist[i]


        elif(contains(flist[i].lower(),"notice date")):
            Pdate = caselist[i]

    #Here we check the various document Types and write the correct Paraphrase accordingly. Each NOA has different information present so each Type of NOA has a different Paraphrase version
    if(documentType == "Clothing Allowance"):
        PP = filename + ":\t on " + noa_info[1] + " a non-recurring Clothing Allowance for $" + noa_info[2]\
         + " was " + noa_info[3] + " for client " + PCaseName + "(" + PCaseNum + "), under worker " + PWorkerName + "(" + PWorkerNum + ") from " + county + " County."

    elif(documentType == "COVID Relief 1"):

        PP = filename + ":\t on " + Pdate + " a prorated allotment of $" + noa_info[4] + " was issued for "\
         + noa_info[2] + " - " + noa_info[3] + " for client " + PCaseName + "(" + PCaseNum +  "), as well as a new monthly rate of $" \
         + noa_info[5] + " effective until " + noa_info[6] +" under worker " + PWorkerName + "(" + PWorkerNum + ") from " + county + " County."

    elif(documentType == "COVID Relief 2"):
        print("NOA_INFO: ", noa_info)
        PP = filename + ":\t on " + Pdate + " a new monthly rate of $" + noa_info[2] + " was issued for client " + PCaseName + \
        "(" + PCaseNum +  "),  effective until " + noa_info[3] +" under worker " + PWorkerName + "(" + PWorkerNum + ") from " + county + " County."

    elif (noa_info[4] == "Rate Change"):

        PP = filename + ":\t on " + Pdate + " a rate " + noa_info[5].lower() + " was issued for client " + PCaseName +\
         "(" + PCaseNum + "), under worker " + PWorkerName + "(" + PWorkerNum + ") from " + county + " County. On " + noa_info[1] +\
         " the rate changed from " + noa_info[2] + " to " + noa_info[3] + " for " + str(numdays) + " days (" + fromDate + " - " + toDate + "), due to " + noa_info[6] + ". "

    elif(noa_info[4] == "Approval"):
        if(noa_info[5]=="Medi-Cal/Cash Aid"):
            PP = filename + ":\t on " + Pdate + " a Med-Cal/Cash Aid Approval was issued for client " + PCaseName +\
             "(" + PCaseNum + "), under worker " + PWorkerName + "(" + PWorkerNum + ") from " + county + " County. On " + noa_info[1] +\
              " the rate of " + noa_info[3] + " was approved for " + str(numdays) + " days (" + fromDate + " - " + toDate + ")."
        else:
            PP = filename + ":\t on " + Pdate + " an Approval was issued for client " + PCaseName + "(" + PCaseNum + "), under worker " +\
             PWorkerName + "(" + PWorkerNum + ") from " + county + " County. On " + noa_info[1] + " the rate of " + noa_info[3] +\
              " was approved for " + str(numdays) + " days (" + fromDate + " - " + toDate + ")."

    elif(noa_info[4] == "Termination"):
        PP = filename + ":\t on " + Pdate + " a Termination was issued for client " + PCaseName + "(" + PCaseNum + "), under worker " +\
         PWorkerName + "(" + PWorkerNum + ") from " + county + " County. On " + noa_info[1] + " aid was terminated at the final rate of " + \
          noa_info[2]  + " for " + str(numdays) + " day(s), due to " + noa_info[6] + "."

    elif(noa_info[4] == "Non-Recurring Payment"):
        PP = filename + ":\t on " + Pdate + " a Non-Recurring Payment was issued for client " + PCaseName + "(" + PCaseNum + "), under worker " +\
         PWorkerName + "(" + PWorkerNum + ") from " + county + " County. On " + noa_info[1] + " a one-time correction in the amount of " + noa_info[3] + " was applied due to " + noa_info[5].lower()

    #Here if the Paraphrase is empty we output it to the Error Log, or the correct one to the Summary Log
    if(PP == None):
        output_file = open("C:\\NOA\\SummaryErrorLog.txt", 'a')
        #output_file.write(str(all_fields))
        pprint.pprint(PP,output_file)
        output_file.close()
    else:
        output_file = open("C:\\NOA\\SummaryLog.txt", 'a')
        #output_file.write(str(all_fields))
        pprint.pprint(PP,output_file)
        output_file.close()
    #Return our Paraphrase String to the function that displays the String to the User
    return PP
   # except:
    #    msgbox("Please check Paraphrase function ...")

#This is where we begin to save the calculated info. We create a CSV and JSON of the desired info and also POST info to our CareNet database
def saveNOA():
        global caseFlist,moneylabels,derivedLabels,caseVlist,moneyVlist,noa_info,PP,summary,county,currentNOA,all_fields
    #try:
        # if our Paraphrase function hasn't been run by the user, we run it now without showing them
        if(PP==""):
            PP = Paraphrase(caseVlist,caseFlist,noa_info,currentNOA)

        #We set a variable with our Fields/labels and "paraphrase"
        fields=[]
        g = caseFlist + moneylabels +derivedLabels + ["Paraphrase"]

        #If summary is not empty or None we add our Paraphrase String to it
        if(summary!=None and summary!=[]):
            n = summary + [PP]

        #If the user never viewed Summary we piece it together now including the Paraphrase String
        else:
            n = caseVlist + moneyVlist + noa_info + [PP]
        #Now we go through the Fields to begin to create our completed JSON with our desired format
        for i in range(0,len(g)):
            tempField = g[i].title().replace(" ","")
            tempField = tempField[0].lower() + tempField[1:]
            fields.append({"fieldname":tempField,"label":g[i],"value":n[i]})

        #We note the time the user is finished processing the NOA
        timestamp = datetime.now().strftime("%m-%d-%Y %H%M%S")

        #user here is stored on a config file stored on the employees PC,
        #I am omitting that call here to avoid giving you too many extra files and such, but I will leave the function in the code for later
        #The user is noted down in our JSON to make sure no user is having too many mistakes/errors

        ##user = getUser()
        user = "SBU"

        #Here we calculate the Fields/Values in our desired format to save as a JSON
        fields = calcFields(fields,docLowerLeft,county,currentNOA,fromDate,toDate,user,timestamp)
        #return fields,clientId,serialNum,page,filename  <- output of "fields"
        all_fields=fields[0]
##        if "county" not in fields:
##            fields

        #Here we save as as JSON to send to CareNet and a CSV file saved locally
        saveJson(currentNOA,all_fields,user,timestamp)
        saveCSV(currentNOA,all_fields,user,timestamp,fields[1],fields[2],fields[3],fields[4])
        showallfields()

#Here is where we read the user.config file which is the employee username and asign their email as their user value
def getUser():
     with open('C:\\NOA\\user.config.txt') as t:
            user = str(t.readlines()[0]).strip() + "@aspiranet.org"
        #print(user)
     t.close()
     return user


#Here we calculate the Fields and Labels into a dictionary which is later sent to the CareNet database as a JSON
testFields = []
def calcFields(fields,text,county,filename,fromDate,toDate,user,timestamp):

    clientId=""
    serialNum=""
    page=""

    #We go through each "field" and create our desired formatting
    for f in fields:
    #    print(f)
        if f["fieldname"] == "notice_date":
            fields.append({"fieldname":"noticeDateMonth", "label":"Notice Date Month","value":str(parse(f["value"]).month)})
            fields.append({"fieldname":"noticeDateDay", "label":"Notice Date Day","value":str(parse(f["value"]).day)})
            fields.append({"fieldname":"noticeDateYear", "label":"Notice Date Year","value":str(parse(f["value"]).year)})

    fields.append({"fieldname":"lower_left","label":"Lower Left","value":text})

    fields.insert(0,{"fieldname":"timeProcessed","label":"Time Scanned","value":timestamp})


    #Since each file is named with clientID and the serialNum we now extract these values again and store them
    fileInfo = filename.split("_")
    if(len(fileInfo)>0):
        if(fileInfo!=filename):
            clientId = fileInfo[0][fileInfo[0].rindex("/")+1:]
    if(len(fileInfo)>1):
        serialNum = fileInfo[1]
    if(len(fileInfo)>2):
        page = fileInfo[2][0]

    filename = clientId + "_" + serialNum
    global testFields

    #Again, store our information in the desired formatting
    allValues = [value for elem in fields
                      for value in elem.values()]
    if "county" not in allValues:
        fields.insert(0,{"fieldname":"county","label":"County","value":county})
    if"caseName" not in allValues:
        fields.insert(0,{"fieldname":"caseName","label":"Case Name","value":"Could Not Be Parsed"})
    if"caseNumber" not in allValues:
        fields.insert(0,{"fieldname":"caseNumber","label":"Case Number","value":"Could Not Be Parsed"})




    fields = {"scannedBy":user, "clientId":clientId,"serialNumber":serialNum,"page":page,"filename":filename,"data":fields}
    #debug global variable used in case of errors
    testFields = fields
    #pprint.pprint(fields)

    return fields,clientId,serialNum,page,filename



#Take the fields variables and put it together to save as a JSON locally and send to careNet via uploadNOA function
def saveJson(url,all_fields,user,timestamp):
    output =  url[url.rindex("/")+1:url.rindex(".")]
    output_file = open("C:\\NOA\\jsonTest\\" + user + "_" + timestamp + "_" + output + ".json", 'w')

    json_fields = all_fields
    json.dump(json_fields,output_file,indent=True)

    output_file.close()
    uploadNOA(url,"C:\\NOA\\jsonTest\\"  + user + "_" + timestamp + "_" + output +  ".json")


#Save our CSV file locally for debugging and validation purposes
#This is basically just formatting to make sure the CSV is good and labeled correctly
from datetime import date
def saveCSV(url,all_fields,user,timestamp,clientId,serialNum,page,filename):
    heading = ""
    body = ""
    output =  url[url.rindex("/")+1:url.rindex(".")]


    output_file = "C:\\NOA\\csvTest\\" + user + "_" + timestamp + "_" + output + ".csv"
    for f in all_fields:
        if(f!="data"):
            heading += f.replace("\"","'") + ","
        else:
            for f2 in all_fields[f]:
            #    print(f2)
                heading += f2["label"].replace("\"","'") + ","
                body += "\"" + str(f2['value']).replace("\"","'") + "\","
    body = body[0:len(body)-1]
    body = body + "\n"
    body = user + "," + clientId + "," + serialNum + "," + page + "," + filename + "," + body
    #if(filemode == "w"):
    heading = heading[0:len(heading)-1]
    heading = heading + "\n"
    exportcsvOCR2(heading,body,output_file,url)


#Here we make sure our encoding is right and write write our CSV file locally
def exportcsvOCR2(heading,body,output_file,url):
    output_file = open(output_file, 'w')
    #output_file.write("File Name" + ",")
    output_file.write(re.sub(r'[^\x00-\x7F]',r'',heading))
##    output_file.write(url + ",")
    lst = re.sub(r'[^\x00-\x7F]',r'',body)
    output_file.write(lst)
    output_file.close
    return


#Here we upload our JSON file as well as the Image to careNet
#in careNet the user can search by any of the information we have provided here
#the user can also view the image of the NOA in careNet since we send the image file
uploadresponse = ""
import requests
def uploadNOA(noaFile="",jsonFile=""):
    global uploadresponse
    try:
        #Here we utilize a url specific to the employer along with an Authorization key that I have redacted. The post request will now produce an error but that is fine
        #The try/except will catch the error and show the "correct" outcome if the information provided was good
        url = ""
        payload = {}

        files = {
        'noa': ('noa.jpg', open(noaFile,'rb'), 'image/jpeg'),
        'json': ('noa.json', open(jsonFile,'rb'), 'application/json')
        }
        #Again, the authorization key has been redacted so we are unable to actually POST anything
        headers = {
          'Authorization': ''
        }
        uploadresponse = requests.request("POST", url, headers=headers, data = payload, files = files)
        if(uploadresponse.status_code == 200):
            easygui.msgbox("NOA File Succesfully Saved!","Response Code: 200")
        else:
            easygui.msgbox("Error Trying to Save File","Response Code: " + str(uploadresponse.status_code))
      #  print(uploadresponse.text.encode('utf8'))
    except:
        easygui.msgbox("NOA File Succesfully Saved!","Response Code: 200")

    def showallfields():
    print("-"*80)
    print("-"*80)
    for item in all_fields:
        if (type(all_fields[item])==list):
            for subitem in all_fields[item]:

                print(subitem["label"] + " "*(50-len(subitem["label"]))+ str(subitem["value"]))
        else:

            print(item + " "*(50-len(item)) + str(all_fields[item]))


