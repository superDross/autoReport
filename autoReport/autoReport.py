import openpyxl
import click
import tqdm
from get_spreadsheet_info import ExtractInfo 
from create_template import create_template
from fill_report import FillReport


@click.command('autoReport')
@click.argument('variant_alias_file')
@click.option('--xlsx')
@click.option('--sheet', default="0")
@click.option('--output')
@click.option('--images')
@click.option('--pdf/--xlsx_only',default='n')
def main(variant_alias_file, output, images, xlsx, sheet, pdf):
    ''' For each variant alias, extract the approriate variant and mutation
        information and export as a variant confirmation report.
        
        variant_alias_file: should be in a text file where each new line
                            contains a different variant alias or a single
                            variant alias.
    '''
    
    # intialise 2 objects using All_Variants & Mutation ID sheets
    extract = ExtractInfo(xlsx, sheet, 2)    

    # open the All_Variants and Mutation ID sheet
    variant_sheet = extract.open_spreadsheets()

    # use the headers of the above workbooks to create a header dicts
    create_variant_dict = extract.create_header_dicts(variant_sheet)

    # create a template report
    create_template(output)
    
    # for each variant alias...

    variant_list = [var_alias for var_alias in open(variant_alias_file, "r")]

    for var_alias in tqdm.tqdm(variant_list):
        var_alias = var_alias.rstrip("\n")
        
        # Open template every iteration, otherwise information leftover from last 
        # iter in cells
        template = openpyxl.load_workbook("test_template.xlsx")
        template_sheet = template.worksheets[0]
    
        # extract variant information from spreadsheet
        variant_row = extract.get_row(variant_sheet, var_alias)
        variant_info = extract.get_query_info(variant_row,variant_sheet)
        
        # get the completed header dictionarys
        variant_dict = extract.header_contents
        
        # with the collected information, fill out the template report. 
        Fill = FillReport(template, template_sheet, var_alias, variant_dict, images)
        Fill.fill_report()
       

        # insert sanger trace and IGV image into the report
        Fill.insert_image(var_alias+"_F", "B22", 600, 242)
        Fill.insert_image(var_alias+"_R", "B22", 600, 242)
        Fill.insert_image(var_alias+"_IGV", "B35", 600, 242)

        
        # get mutation information for the given variant alias if an ID is there
        Fill.pick_comment()

        #print("\t".join((var_alias,variant_dict.get("Category"),
                        # variant_dict.get("Mutation_ID"))))

        output_name = (output+str(variant_dict.get("Sample_Name"))+
                            "_"+str(variant_dict.get("Variant_Alias"))+"_"+
                            "VariantConfirmationReport"+".xlsx")

        # save the filled sheet and convert to PDF
        template.save(output_name)
        if pdf:
            Fill.convert2pdf(output_name)
        
        # clear the dicts items
        variant_dict = {key: "-" for key in variant_dict}
    



if __name__ == '__main__':
    main()


