#!/usr/local/bin/macruby
require "rubygems"
require "yaml"
framework "Cocoa"
framework "AddressBook"

class ABPerson
  def []=(k, v)
    return if !v
    self.setValue v, :forProperty => k
  end
end

def add_multi_value(items)
  return nil if !(first_item = items.shift)
  mv = ABMutableMultiValue.new  
  mv.addValue first_item, :withLabel => "work"

  while next_item = items.shift
    mv.addValue next_item, :withLabel => "work"
  end
  mv.setPrimaryIdentifier(mv.identifierAtIndex(0))    
  mv
end

sab = ABAddressBook.sharedAddressBook

rows = YAML.load(STDIN.readlines.join("\n"))
colmap = rows.shift.inject([{}, 0]) { |m, k| [(m[0][k.downcase.gsub(/\W/, "").intern] = m[1] ; m[0]), m[1] + 1]  }[0]
STDERR.puts colmap.to_yaml

rows.each do | row |
  entry = colmap.keys.inject({}) { |m, k| m[k] = row[colmap[k]] || "" ; m }

  person = ABPerson.new

  person["Nickname"] = entry[:nickname]

  names = entry[:name].split(/, /)
  
  if names.size == 1 || entry[:tags] =~ /org/
    if entry[:tags] =~ /org/
    else
      person["Last"] = names[0]
    end
  elsif names.size > 1
    person["Last"] = names.shift
    # TO DO: Check this rule...
    person["First"] = names.join(", ")    
  else
  end
  
  # trim begin/end spaces: a.gsub!(/^\s*(.+?)\s*$/, '\1')
  
  # email
  emails = entry[:emails].split(/;/).map { |a| a.gsub(/\s/, "") }
  person["Email"] = add_multi_value(emails)
  
  phones = entry[:phones].split(/;/).map { |a| a.gsub(/^\s*(.+?)\s*$/, '\1') }
  person["Phone"] = add_multi_value(phones)    
  
  person["Note"] = entry[:notes] + " abtool"
  sab.addRecord person
end

sab.save
