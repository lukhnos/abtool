#!/usr/bin/env python
import re
import sys
import xlrd

from AddressBook import *
from Foundation import *

def multi_value_from_items(items):
    if len(items) == 0:
        return None
        
    mv = ABMutableMultiValue.new()

    for item in items:
        mv.addValue_withLabel_(item, "work")
    
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

me_record = None

print ab

for rx in range(1, sh.nrows):
    row = sh.row(rx)

    v = lambda t, y: map(lambda x: x.strip(), str.split(row[tm[t]].value.encode('utf-8'), y))
    
    tags = row[tm["tags"]].value.encode('utf-8')  
    nick = row[tm["nickname"]].value.encode('utf-8')  
    name_components = v("name", ", ")
    kanji_components = v("kanji", " ")
    phones = v("phones", "; ")
    emails = v("emails", "; ")
    addrs = v("addresses", "; ")
    title_org_components = v("title", " // ")
    notes = row[tm["notes"]].value.encode('utf-8')
    
    print("nick %s, phones: %s" % (nick, ",".join(phones)))
    
    person = ABPerson.new()
    person.setValue_forProperty_(nick, "Nickname")
    
    if len(notes) > 0:    
        notes += "\nnewabt"
    else:
        notes = "newabt"

    person.setValue_forProperty_(multi_value_from_items(emails), "Email")
    
    person.setValue_forProperty_(notes, "Note")
    ab.addRecord_(person)

    if re.search(ME_NICK, nick):
        me_record = person

if me_record:
    ab.setMe_(me_record)

ab.save()

    

    # rows.each do | row |
    #   person["Nickname"] = entry[:nickname]
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

    