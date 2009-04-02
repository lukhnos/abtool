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
# STDERR.puts rows.to_yaml

rows.each do | row |
  entry = colmap.keys.inject({}) { |m, k| m[k] = row[colmap[k]] || "" ; m }

  person = ABPerson.new

  person["Nickname"] = entry[:nickname]

  # names = entry[:name].split(/,/)
  names = entry[:name].componentsSeparatedByString(", ")
  
  # puts "Splitting name: #{names.join(";;;")} for name: '#{entry[:name]}'"
  
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
  
  if entry[:addresses].length > 0  
    addrs = entry[:addresses].componentsSeparatedByString(";").map { |a| a.gsub(/^\s*(.+?)\s*$/, '\1').gsub(/\s?\/\/\s?/, "\n") } 
    
    # puts addrs.join(">>>")
    
    if first_addr = addrs.shift
        
      mv = ABMutableMultiValue.new  
      mv.addValue({ "Street" => first_addr }, :withLabel => "work")

      while next_addr = addrs.shift
        mv.addValue({ "Street" => next_addr }, :withLabel => "work")
      end
      mv.setPrimaryIdentifier(mv.identifierAtIndex(0))    
      person["Address"] = mv
    end
  end

  if entry[:notes].length > 0
    n = entry[:notes] + "\nabtool" 
    person["Note"] = n    
  else
    person["Note"] = "abtool"
  end
  
  sab.addRecord person
end

sab.save
