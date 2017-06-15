

require 'rubygems'
require 'roo'
require 'lcsort'
require 'sequel'
require 'sqlite3'

puts "foo"

# Open spreadsheet
xlsxLookup = Roo::Spreadsheet.open('./LookupSheet.xlsx')

# Create in-memory database and lookup table in that database
DBLookupAndResults = Sequel.sqlite
DBLookupAndResults.create_table :lookup do
	string :RangeName
	string :BeginCallNumber
	string :EndCallNumber
	string :CollectionCode
end


lookup = DBLookupAndResults[:lookup]

# Loop over Lookup sheet, inserting into Lookup database table
xlsxLookup.each_row_streaming do |row|
	lookup.insert(:RangeName => row[0].to_s, :BeginCallNumber => row[1].to_s, :EndCallNumber => row[2].to_s, :CollectionCode => row[3].to_s)
end

# Create dataset of All Records, then sort it
dsAllRecords = DBLookupAndResults[:lookup]
dsAllRecords = dsAllRecords.order(:BeginCallNumber)

# Loop over lookup dataset
dsAllRecords.each do |row|
	beginCallNumberSortKey = row[:BeginCallNumber].to_s
	beginCallNumberSortKey = Lcsort.normalize(beginCallNumberSortKey)
	endCallNumberSortKey = row[:EndCallNumber].to_s
	endCallNumberSortKey = Lcsort.normalize(endCallNumberSortKey)

	if (beginCallNumberSortKey.nil? || endCallNumberSortKey.nil?)
		puts row
	end	
end



