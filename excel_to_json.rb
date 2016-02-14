require 'roo'
require 'colorize'
require 'optparse'

def is_number? string
  true if Float(string) rescue false
end

def open_xlsx(document)
  extension = document[(document.length-4)..(document.length)]
  if extension.upcase() =="XLSX"
    puts "[OK] file extension verified, Importing : ".colorize(:green)<< document
    excel = Roo::Spreadsheet.open(document)
  else
    puts "[OH] Bad file extention. quitting :(".colorize(:red)
    abort()
  end
  return excel
end

def convert(sheet)

  puts "[OK] Detected #{sheet.last_row} rows and #{sheet.last_column} columns".colorize(:green)
  # Iterate through each sheet
  # intilising the data var
  headers = sheet.row(sheet.first_row)
  data = "["
  #iterating oer the sheet rows
  for irow in 2..sheet.last_row
    data += "{"
    for icol in sheet.first_column..sheet.last_column
      if is_number?  sheet.cell(irow,icol)
          data += '"' + "#{headers[icol-1]}" + '":'+ "#{sheet.cell(irow,icol)}"+ ', '
        else
          data += '"' + "#{headers[icol-1]}" + '":"'+ "#{sheet.cell(irow,icol)}"+ '", '
        end
    end
    data += "},"
  end
  # Gsub replaces all instances of ", }" by "}".
  # Remove  the last ,
  # Add ] to the end
  json_data = data.gsub(", }", "}")[0..-2] + "]"
  return json_data
end

def write_file(data , path, filename)
  #writing to file
  output_file = "#{path}/#{filename}.json"
  puts "[OK] writing to file :".colorize(:green) << output_file
  file = File.open(output_file, 'w+')
  file.write(data)
  file.close
end

def main
  options = {}
  optparse = OptionParser.new do |opts|
    opts.banner = "Usage: ruby excel_to_json.rb [options]"
     opts.on('-i', '--input filename', 'Input a valid .xlsx file') { |file| options[:xlsxfile] = file }
     opts.on('-o', '--output PATH', 'Output directory') { |dir| options[:outputdir] = dir }
     opts.on('-v', '--verbose', 'Run verbosely') { |v| options[:verbose] = v }
   end
   begin
     optparse.parse!
     if options[:outputdir].nil? then options[:outputdir] = "./" end
     if options[:xlsxfile].nil?
       raise OptionParser::MissingArgument
     end
   rescue OptionParser::ParseError => e
     puts "[OH] No file supplied, Noting to do!!".colorize(:red)
     puts optparse
     exit
   end
   # Read the xlsx file
  excel = open_xlsx(options[:xlsxfile])
  # loop through sheets
  excel.each_with_pagename do |name, sheet|
    puts "[OK] Converting sheet: ".colorize(:green) << name
    json_data = convert(sheet)
    write_file(json_data, options[:outputdir], name)
  end
end

main()
