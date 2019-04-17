A python tool to lookup if a given pincode exists or not by web scraping and finally mapping state & city data to the output sheet.

# Python packages used
  1. pandas -  To scrap urls
  2. BeautifulSoup - To scrap urls in case Pandas fail to output required data.
  3. requests - To make GET calls to urls.
  4. xlrd - To read microsoft excel.
  5. xlwt - To write to an excel file.

## Set input file path
   `loc = ("pincode.xlsx")`

## Set URL where pincode lookup has to be done
   `url = "https://www.mapsofindia.com/pincode/"`
Note: Currently lookup can be done in two websites only. Go through the code for the URLs.

## Set output file path
   `output_file_path = "D:/python/pincode_lookup/out/"`

Run the code, A balloon tooltip will appear once completed.(Only for windows)
This tool will pick up the pincodes from input sheet and do a lookup in the URLs provided.
Once an existing pincode is found, Its state & city will be written to the output file.
