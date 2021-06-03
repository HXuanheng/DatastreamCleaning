# DatastreamCleaning

#Step: 1
pip3 install -r requirements.txt

#Step: 2
write_excel.py -> prepare excel (.xlsx) files for datastream add-in to get data
input: specify /resources/parameters/fields.xlsx and /resources/parameters/identifiers.xlsx
python write_excel.py
output: /results/to_be_upgraded/

#Step: 3
update the excel files in /results/to_be_upgraded/ with the datastream add-in

#Step: 4
clean_data.py-> clean and merge datastream add-in output
input: move the upgraded excel files in /resources/to_be_cleaned/
python clean_data.py
output: /results/cleaned/ and /results/merged/