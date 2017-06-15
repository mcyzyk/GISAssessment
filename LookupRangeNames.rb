# LookupRangeNames.rb
# mcyzyk, May 2017
#
# This script looks up Range Name per Call Number
#
# It takes as its input a spreadsheet consisting of a single column 
# of call numbers, no header. This spreadsheet must be named "input.xlsx"
# and must be placed in the same directory with this script.
#
# It results in a CSV file, output.csv, consisting of two columns: 
# One holding CallNumbers; the second holding associated RangeNames.
# No column headers.

require 'rubygems'
require 'roo'
require 'lcsort'
require 'sequel'
require 'sqlite3'
require 'csv'

puts "foo"

# Open spreadsheets
xlsx = Roo::Spreadsheet.open('./input.xlsx')
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

# Create results table in database
DBLookupAndResults.create_table :results do
        string :CallNumber
        string :RangeName
end

results = DBLookupAndResults[:results]

# CollectionCodes:
# 	Eisenhower General Collection
# 	Eisenhower D Level Blue Labels
# 	Eisenhower A Level International Government Doc
# 	Eisenhower A Level Atlases

# Create dataset of All Records, then sort it
dsAllRecords = DBLookupAndResults[:lookup]
dsAllRecords = dsAllRecords.order(:BeginCallNumber)

# Create dataset of Eisenhower General Collection that are NOT Quarto or Folio 
theQueryString = Sequel.lit('BeginCallNumber NOT LIKE "%QUARTO" AND BeginCallNumber NOT LIKE "%FOLIO" AND CollectionCode = "Eisenhower General Collection"')
dsEisenhowerGeneralCollection = dsAllRecords.where(theQueryString)

# Create dataset of Quartos from Eisenhower General Collection
theQueryString = Sequel.lit('BeginCallNumber LIKE "%QUARTO" AND CollectionCode = "Eisenhower General Collection"')
dsQuarto = dsAllRecords.where(theQueryString)

# Create dataset of Folios from Eisenhower General Collection
theQueryString = Sequel.lit('BeginCallNumber LIKE "%FOLIO" AND CollectionCode = "Eisenhower General Collection"')
dsFolio = dsAllRecords.where(theQueryString)

# Create dataset of Eisenhower D Level Blue Labels
theQueryString = Sequel.lit('CollectionCode = "Eisenhower D Level Blue Labels"')
dsEisenhowerDLevelBlueLabels = dsAllRecords.where(theQueryString)

# Create dataset of Eisenhower A Level International Government Doc
theQueryString = Sequel.lit('CollectionCode = "Eisenhower A Level International Government Doc"')
dsEisenhowerALevelInternationalGovernmentDoc = dsAllRecords.where(theQueryString)

# Create dataset of Eisenhower A Level Atlases
theQueryString = Sequel.lit('CollectionCode = "Eisenhower A Level Atlases"')
dsEisenhowerALevelAtlases = dsAllRecords.where(theQueryString)


# Loop over CallNumbers
xlsx.each_row_streaming do |callnumberCurrent|

	# At this point, callnumberCurrent is an Excel row.  Must get at individual value and create its Lcsort sortkey
	callnumberCurrentSortKey = callnumberCurrent[0].to_s
	callnumberCurrentSortKey = Lcsort.normalize(callnumberCurrentSortKey)

	puts callnumberCurrent[0].to_s

	# Next if sortkey is nil
	if (callnumberCurrentSortKey == nil)
		# Insert into results table
		results.insert(:CallNumber => callnumberCurrent[0].to_s, :RangeName => "RANGENAME NOT FOUND")
		puts
		next
	end

	# Set current dataset pointer 
        if callnumberCurrent[0].to_s.include? "QUARTO"
                dsCurrent = dsQuarto
        elsif callnumberCurrent[0].to_s.include? "FOLIO"
                dsCurrent = dsFolio
        elsif callnumberCurrent[0].to_s.include? "Eisenhower D Level Blue Labels"
                dsCurrent = dsEisenhowerDLevelBlueLabels
        elsif callnumberCurrent[0].to_s.include? "Eisenhower A Level International Government Doc"
                dsCurrent = dsEisenhowerALevelInternationalGovernmentDoc
        elsif callnumberCurrent[0].to_s.include? "Eisenhower A Level Atlases"
                dsCurrent = dsEisenhowerALevelAtlases
        else
                dsCurrent = dsEisenhowerGeneralCollection
        end

	# Winnow the current dataset to JUST the rows beginning with the first letter of the current call number
	firstLetter = callnumberCurrent[0].to_s.chr
	dsCurrent = dsCurrent.where(Sequel.like(:BeginCallNumber, "#{firstLetter}%")).or(Sequel.like(:EndCallNumber, "#{firstLetter}%"))

	# Loop over lookup dataset
	dsCurrent.each do |row|
		# if callnumberCurrent GTE row[:BeginCallNumber] AND callnumberCurrent LTE row[:EndCallNumber] then we've found our RangeName!
		beginCallNumberSortKey = row[:BeginCallNumber].to_s
		beginCallNumberSortKey = Lcsort.normalize(beginCallNumberSortKey)
		endCallNumberSortKey = row[:EndCallNumber].to_s
		endCallNumberSortKey = Lcsort.normalize(endCallNumberSortKey)

		# If not nil
		if (!beginCallNumberSortKey.nil? && !callnumberCurrentSortKey.nil? && !endCallNumberSortKey.nil?)

			callnumberArray = [beginCallNumberSortKey, callnumberCurrentSortKey, endCallNumberSortKey]
			callnumberArraySorted = callnumberArray.sort

			if (callnumberCurrentSortKey == beginCallNumberSortKey || callnumberCurrentSortKey == endCallNumberSortKey || callnumberCurrentSortKey == callnumberArraySorted[1].to_s) 
				# We've found our RangeName!
				# Insert into results table
				results.insert(:CallNumber => callnumberCurrent[0].to_s, :RangeName => row[:RangeName].to_s)
				puts row[:RangeName].to_s
				puts
				break
			end
		end	
	end
end

# Output results dataset to CSV file
CSV.open("output.csv", "wb") do |csv|
	results.each do |row|
		csv << row.values.to_a
	end
end

