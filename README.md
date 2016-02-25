# Excel_to_json
Excel_to_json is a Simple cli tool convert Excel files to JSON format.

## Installation
Excel_to_json is written in Ruby using the roo library. Installation is as easy as:
```
cd excel_to_json
gem install bundler
bundle install
```
This tool was devloped using ruby version 2.2.4 and it tested sucessfully using ruby 2.1.5.
```
rvm install 2.1.5
rvm use 2.1.5
```
##Usage 
This tool takes an XLSX file as input. It parse and convert the file sheets to JSON format and save them to separated files under the output folder. 
ruby excel_to_json.rb -ri templates/sample.xlsx -o templates/ 
```
ruby excel_to_json.rb --help
Usage: ruby excel_to_json.rb [options]
    -i, --input filename             Input a valid .xlsx file
    -o, --output PATH                Output directory
    -r, --recursive                  Loop through all Excel sheets
```
If non recursive mode is chosen only the first sheet willl be converted.

PS: if you don't need the colorized output.the no-color version of this tool is better for you : excel_to_json_nc.rb
## About Excel_to_json
Excel files are awesome for the end user but are hated by developers because they are such a pain to be correctly parsed.
Excel_to_json try to avoid you that pain by translating the Excel data to JSON format which can be easily programmatically processed.

## Known issues
Excel_to_json currently doesn't handle the " (Double quote), feel free to contribute.
I was too lazy to do it because it just suites my needs for now.
