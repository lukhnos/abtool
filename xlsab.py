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


ME_NICK = "lukhnos"

if len(sys.argv) < 2:
  print("usage: import.py filename")
  sys.exit(1)

xls_path = sys.argv[1]

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

    if re.search(ME_NICK, nick):
        me_record = person

if me_record:
    ab.setMe_(me_record)

ab.save()

    

    # rows.each do | row |
    # 
    #   # names = entry[:name].split(/,/)
    #   names = entry[:name].componentsSeparatedByString(", ")
    # 
    #   # puts "Splitting name: #{names.join(";;;")} for name: '#{entry[:name]}'"
    # 
    #   if names.size == 1 || entry[:tags] =~ /org/ || entry[:tags] =~ /service/
    #     if entry[:tags] =~ /org/ || entry[:tags] =~ /service/
    #       person["Organization"] = entry[:name]
    #       person["ABPersonFlags"] = 1
    #     else
    #       person["Last"] = names[0]
    #     end
    #   elsif names.size > 1
    #     person["Last"] = names.shift
    #     # TO DO: Check this rule...
    #     person["First"] = names.join(", ")    
    #   else
    #   end
    # 
    #   # trim begin/end spaces: a.gsub!(/^\s*(.+?)\s*$/, '\1')
    # 
    #   # email
    #   emails = entry[:emails].split(/;/).map { |a| a.gsub(/\s/, "") }
    #   person["Email"] = add_multi_value(emails)
    # 
    #   phones = entry[:phones].split(/;/).map { |a| a.gsub(/^\s*(.+?)\s*$/, '\1') }
    #   person["Phone"] = add_multi_value(phones)    
    # 
    #   if entry[:addresses].length > 0  
    #     addrs = entry[:addresses].componentsSeparatedByString(";").map { |a| a.gsub(/^\s*(.+?)\s*$/, '\1').gsub(/\s?\/\/\s?/, "\n") } 
    # 
    #     # puts addrs.join(">>>")
    # 
    #     if first_addr = addrs.shift
    # 
    #       mv = ABMutableMultiValue.new  
    #       mv.addValue({ "Street" => first_addr }, :withLabel => "work")
    # 
    #       while next_addr = addrs.shift
    #         mv.addValue({ "Street" => next_addr }, :withLabel => "work")
    #       end
    #       mv.setPrimaryIdentifier(mv.identifierAtIndex(0))    
    #       person["Address"] = mv
    #     end
    #   end
    # 
    #   if entry[:notes].length > 0
    #     n = entry[:notes] + "\nabtool" 
    #     person["Note"] = n    
    #   else
    #     person["Note"] = "abtool"
    #   end
    # 
    #   sab.addRecord person
    # end
    # 
    # sab.save

    