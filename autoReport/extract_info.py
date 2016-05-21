import click
from get_spreadsheet_info import ExtractInfo


# Option to ouput information into a spreadsheet from a list/file of queries

@click.command('extract_info')
@click.option('--xlsx', help="XLSX workbook to scrape from")
@click.argument('query', required=False)
@click.option('--sheet', default="0", help="specify sheet in workbook to scrape from")
@click.option('--row_header', default=1, help="row number which contains column names")
@click.option('--item', multiple=True, default=None, help="specify column name to scrape from")
@click.option('--keys/--none', default='n', help="print all column names in spreadsheet")

def main(xlsx, query, sheet="0", row_header=1, item=None, keys=None):
    ''' Extract information from a given spreadsheet.

        An argument is used to search the first column
        for a match and return the column information
        specified in the --item option(s). 
        
    '''
    # warn user that an argument was not given 
    if not query:
        print("WARNING: No argument given")
    
    # initialise the object
    extract = ExtractInfo(xlsx,sheet,row_header)          
    
    # open the spreadsheet and workbook
    spreadsheet = extract.open_spreadsheets()
    
    # use the headers of the above work book to create a header dict
    create_header_dict = extract.create_header_dicts(spreadsheet)
    
    # get row number in which the query is within
    row_of_interest = extract.get_row(spreadsheet, query)

    # get the query information for the query given and append to dict
    query_info = extract.get_query_info(row_of_interest,spreadsheet)
    
    # simplify dict name
    header_dict = extract.header_contents
    
    # determine output upon options selected
    output = output_options(item,keys,header_dict)
    return output


def output_options(items,keys,header_dict):
    
    # iterate through each element and get its item from the dict.
    if items:
       for item in items:
            out = (item,"-",str(header_dict.get(item)))
            print(" ".join(out))
    
    # print all keys to the screen
    elif keys:
        for key,value in sorted(header_dict.items()):
            print(key)
    
    else:
        # print all keys and items
        for key, value in sorted(header_dict.items()):
            print(key,value)

if __name__ == '__main__':
    main()
