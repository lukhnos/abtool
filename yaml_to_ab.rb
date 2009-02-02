#!/usr/local/bin/macruby
require "rubygems"
require "yaml"
framework "Cocoa"
framework "AddressBook"

class ABPerson
  def []=(k, v)
    self.setValue v, :forProperty => k
  end
end

sab = ABAddressBook.sharedAddressBook

rows = YAML.load(STDIN.readlines.join("\n"))
colmap = rows.shift.inject([{}, 0]) { |m, k| [(m[0][k.downcase.gsub(/\W/, "").intern] = m[1] ; m[0]), m[1] + 1]  }[0]
STDERR.puts colmap.to_yaml

rows.each do | row |
  entry = colmap.keys.inject({}) { |m, k| m[k] = row[colmap[k]] || "" ; m }

  person = ABPerson.new

  names = entry[:name].split(/, /)
  
  if names.size == 1
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
  
  
  person["Note"] = entry[:notes] + " abtool"
  person["Nickname"] = entry[:nickname]
  sab.addRecord person
end

sab.save
