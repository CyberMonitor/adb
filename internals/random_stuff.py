import random
from datetime import datetime
from datetime import timedelta
import os

firstnamespath = os.path.join(os.path.split(__file__)[0], 'first-names.txt')
with open(firstnamespath, 'r') as r:
    firstnames = set(r.readlines())

middlenamespath = os.path.join(os.path.split(__file__)[0], 'middle-names.txt')
with open(middlenamespath, 'r') as r:
    middlenames = set(r.readlines())

lastnamespath = os.path.join(os.path.split(__file__)[0], 'last-names.txt')
with open(lastnamespath, 'r') as r:
    lastnames = set(r.readlines())

def getrandomauthor():
    style = random.randint(1, 5)
    fname = str((random.sample(firstnames, k=1)[0].rstrip("\n")))
    mname = str((random.sample(middlenames, k=1)[0].rstrip("\n")))
    lname = str((random.sample(lastnames, k=1)[0].rstrip("\n")))
    if style == 1:  # first name only
        authorname = fname
        initials = fname[0].upper()
    elif style == 2:  # first initial and last name
        firstinitial = fname[0]
        lastname = lname
        authorname = firstinitial + ' ' + lastname
        initials = fname[0].upper() + lname[0].upper()
    elif style == 3:  # first name and last initial
        firstname = fname
        lastinitial = lname[0]
        authorname = firstname + ' ' + lastinitial
        initials = fname[0].upper() + lname[0].upper()
    elif style == 4:  # first and last name
        firstname = fname
        lastname = lname
        authorname = firstname + ' ' + lastname
        initials = fname[0].upper() + lname[0].upper()
    elif style == 5:  # first name, middle initial, last name
        firstname = fname
        middleinitial = mname[0]
        lastname = lname
        authorname = firstname + ' ' + middleinitial + ' ' + lastname
        initials = fname[0].upper() + mname[0].upper() + lname[0].upper()

    return (initials, authorname)

def getrandomdate(years_to_go_back=3):
    days_to_subtract = years_to_go_back * 365
    oldest_date = datetime.now() - timedelta(days=days_to_subtract)
    oldest_date_epoch = datetime.timestamp(oldest_date)
    current_date_epoch = datetime.timestamp(datetime.now())

    return datetime.fromtimestamp(random.randint(int(oldest_date_epoch), int(current_date_epoch)))

def get_random_date_between_then_and_now(then_datetime):
    oldest_date_epoch = datetime.timestamp(then_datetime)
    current_date_epoch = datetime.timestamp(datetime.now())
    return datetime.fromtimestamp(random.randint(int(oldest_date_epoch), int(current_date_epoch)))

def getrandomrevision():
    return random.randint(3, 21)

def getrandomparagraphcount():
    return random.randint(1, 23)

def getrandomtotaltime():
    return random.randint(30, 310)

def getrandompages():
    return random.randint(1, 13)

def getrandomlinecount(paragraphcount):
    return int(paragraphcount * 3)

def getrandomwordcount(paragraphcount):
    return int(random.randint(15, 25) * paragraphcount)

def getrandomcharactercount(wordcount):
    return int(random.randint(3, 7) * wordcount)

def getrandomcharactercountwithspaces(charactercount):
    return round(charactercount / 5 + charactercount)

def gen_doc_name(extension="doc"):
    word1 = ["Final", "Outstanding", "Next", "Selected", "Partial", "Needed", "Sales",
             "Past_Due", "Overdue", "Paid", "Incorrect", "New", "Your"]
    word2 = ["Invoice", "Bill", "Ticket", "Payment", "Receipt", "Invoices"]
    name = word1[random.randint(0, len(word1)-1)] + "_" + word2[random.randint(0, len(word2)-1)] + "_" + str(random.randint(100, 10000)) + "." + extension
    return name

# tests


if __name__ == '__main__':
    print("Author: " + str(getrandomauthor()))
    this_random_date = getrandomdate()
    print("Date created: " + str(this_random_date))
    print("Date modified: " + str(get_random_date_between_then_and_now(this_random_date)))
    print("Doc name: " + str(gen_doc_name()))
