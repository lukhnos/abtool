#!/usr/bin/ruby

# Public domain. Written by Lukhnos D. Liu (lukhnos@lukhnos.org)

$KCODE = "u"
require "rubygems"
require "yaml"
require "parseexcel"

if ARGV.size == 0
  STDERR.puts "usage: xls_to_json <spreadsheets>..."
  exit 1
end

ARGV.each do |filename|
  wb = Spreadsheet::ParseExcel.parse(filename)
  rows = wb.worksheet(0).map { |r| r }.compact
  grid = rows.map { |r| r.map() { |c| c ? c.to_s("utf-8") : "" } }
    

  puts grid.to_yaml
end
