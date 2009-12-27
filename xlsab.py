#!/usr/bin/env python
import re
import sys
import xlrd

from AddressBook import *
from Foundation import *

def multi_value_from_items(items, default_label, *labels):
    if len(items) == 0:
        return None
        
    filled_labels = list(labels)
    while len(filled_labels) < len(items):
        filled_labels.append(default_label)
    ifl = zip(items, filled_labels)
            
    mv = ABMutableMultiValue.new()

    for item, labels in ifl:
        mv.addValue_withLabel_(item, labels)
    
    mv.setPrimaryIdentifier_(mv.identifierAtIndex_(0))
    return mv


ME_NICK = None

if len(sys.argv) < 2:
  print("usage: import.py filename [me_nick]")
  sys.exit(1)

xls_path = sys.argv[1]

if len(sys.argv) > 2:
    ME_NICK = sys.argv[2]

book = xlrd.open_workbook(xls_path)
sh = book.sheet_by_index(0)

if sh.nrows < 1:
    print("must have at least 1 row")
    sys.exit(1)

tm = {}
index = 0

# build tag map
for t in sh.row(0):
    ut = t.value
    if re.search("title", ut):
        tm["title"] = index
    else:
        tm[ut] = index
        
    index = index + 1

print("tags: %s" % tm)

ab = ABAddressBook.sharedAddressBook()

IMPORT_TOOL_MARK = "tag:abtool"

for p in ab.people():
    tag_str = p.valueForProperty_(kABNoteProperty)
    if tag_str and re.search(IMPORT_TOOL_MARK, tag_str):
        ab.removeRecord_(p)

me_record = None

print ab

for rx in range(1, sh.nrows):
    row = sh.row(rx)

    v = lambda t, y: map(lambda x: x.strip(), str.split(row[tm[t]].value.encode('utf-8'), y))
    
    tag_str = row[tm["tags"]].value.encode('utf-8')
    
    if len(tag_str):
        tags = v("tags", ",")
    else:
        tags = []
        
    nick = row[tm["nickname"]].value.encode('utf-8')  
    name = row[tm["name"]].value.encode('utf-8')
    name_components = v("name", ", ")
    kanji = row[tm["kanji"]].value.encode('utf-8') 
    phones = v("phones", "; ")
    emails = v("emails", "; ")
    addrs = v("addresses", "; ")
    
    title_orgs = v("title", "; ")
    title_org_components = map(lambda x: x.strip(), str.split(title_orgs[0], " // "))
    
    notes = row[tm["notes"]].value.encode('utf-8')
	

    person = ABPerson.new()
    # print("nick %s, phones: %s" % (nick, ",".join(phones)))
	
    if re.search("org", tag_str):
        person.setValue_forProperty_(kABShowAsCompany, kABPersonFlags)

        if len(name) == 0 and len(nick) > 0:
            person.setValue_forProperty_(nick, kABOrganizationProperty)
        else:
            person.setValue_forProperty_(name, kABOrganizationProperty)
    else:
        if len(name_components) == 1:
            if len(name_components[0]) == 0 and len(nick) > 0:
                person.setValue_forProperty_(nick, kABNicknameProperty)
            else:
                person.setValue_forProperty_(name_components[0], kABFirstNameProperty)			
        elif len(name_components) == 2:
            person.setValue_forProperty_(name_components[0], kABLastNameProperty)
            person.setValue_forProperty_(name_components[1], kABFirstNameProperty)			
        elif len(name_components) >= 3:
            # TODO: Join the remaining components
            person.setValue_forProperty_(name_components[0], kABLastNameProperty)
            person.setValue_forProperty_(name_components[1], kABFirstNameProperty)
            person.setValue_forProperty_(name_components[2], kABSuffixProperty)
    
        if len(title_org_components) > 0:
            if len(title_org_components) == 1:
                person.setValue_forProperty_(title_org_components[0], kABOrganizationProperty)
            else:
                # TODO: Join the remaining components
                person.setValue_forProperty_(title_org_components[0], kABJobTitleProperty)
                person.setValue_forProperty_(title_org_components[1], kABOrganizationProperty)

    
    # person.setValue_forProperty_(nick, "Nickname")
    
    tag_notes = map(lambda x: "tag:" + x, tags)
    tag_notes += [IMPORT_TOOL_MARK]
    
    if len(kanji) > 0:
        tag_notes += ["kanji:%s" % kanji]
    
    note_appendix = "\n".join(tag_notes)
    
    if len(notes) > 0:    
        notes += "\n" + note_appendix
    else:
        notes = note_appendix

    person.setValue_forProperty_(multi_value_from_items(emails, kABEmailWorkLabel, kABEmailWorkLabel, kABEmailHomeLabel), kABEmailProperty)
    person.setValue_forProperty_(multi_value_from_items(phones, kABPhoneHomeLabel, kABPhoneMobileLabel, kABPhoneWorkLabel), kABPhoneProperty)
    
    mapped_addrs = map(lambda x: { kABAddressStreetKey : x }, addrs)
    person.setValue_forProperty_(multi_value_from_items(mapped_addrs, kABAddressWorkLabel, kABAddressWorkLabel, kABAddressHomeLabel), kABAddressProperty)
    
    person.setValue_forProperty_(notes, kABNoteProperty)
    ab.addRecord_(person)

    if ME_NICK and re.search(ME_NICK, nick):
        me_record = person

if me_record:
    ab.setMe_(me_record)

ab.save()
 