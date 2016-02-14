# Excel_to_json
Excel_to_json is a Simple cli tool convert Excel files to JSON format.

## Installation
Excel_to_json is written in Ruby using the roo library. Installation is as easy as:
```
cd excel_to_json
gem install bundler
bundle install
```
ruby version 2.1.5 is supported.
```
rvm install 2.1.5
rvm use 2.1.5
```
## About Excel_to_json
Excel files are awesome for the end user but are hated by developers because they are such a pain to be correctly parsed.
Excel_to_json try to avoid you that pain by translating the Excel data to JSON format which can be easily programmatically processed.

## Known issues
Excel_to_json currently doesn't handle the " (Double quote), feel free to contribute.
I was too lazy to do it because it just suites my needs for now.
