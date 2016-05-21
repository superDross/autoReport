import openpyxl
from get_spreadsheet_info import ExtractInfo 
from create_template import create_template
from fill_report import FillReport
from tqdm import * 

def produce_variant_report(variant_alias_file, output_path, image_dir, xlsx, sheet):
    ''' For each variant alias, extract the approriate variant and mutation
        information and append them to the variant confirmation template
        
            variant_alias_file: should be in a text file where each new line
                                contains a different variant alias or a single
                                variant alias.
    '''
    # intialise 2 objects using All_Variants & Mutation ID sheets
    extract = ExtractInfo(xlsx, sheet, 2)    
    extract_mutation = ExtractInfo(xlsx, "Mutations ID", 2)

    # open the All_Variants and Mutation ID sheet
    variant_sheet = extract.open_spreadsheets()

    # use the headers of the above workbooks to create a header dicts
    create_variant_dict = extract.create_header_dicts(variant_sheet)

    # create a template report
    create_template(output_path)
    
    # for each variant alias...

    variant_list = [var_alias for var_alias in open(variant_alias_file, "r")]

    for var_alias in tqdm(variant_list):
        var_alias = var_alias.rstrip("\n")
        
        # Open template every iteration, otherwise information leftover from last iter in
        # cells 
        template = openpyxl.load_workbook("test_template.xlsx")
        template_sheet = template.worksheets[0]
    
        # extract variant information from spreadsheet
        variant_row = extract.get_row(variant_sheet, var_alias)
        variant_info = extract.get_query_info(variant_row,variant_sheet)
        
        # get the completed header dictionarys
        variant_dict = extract.header_contents
        
        # with the collected information, fill out the template report. 
        Fill = FillReport(template, template_sheet, var_alias, variant_dict, image_dir)
        Fill.fill_report()

        # insert sanger trace and IGV image into the report
        Fill.insert_image(var_alias+"_F", "B22")
        Fill.insert_image(var_alias+"_IGV", "B35")

        
        # get mutation information for the given variant alias if an ID is there
        Fill.pick_comment()

        #print("\t".join((var_alias,variant_dict.get("Category"),
                        # variant_dict.get("Mutation_ID"))))

        output_file_name = (output_path+str(variant_dict.get("Sample_Name"))+
                            "_"+str(variant_dict.get("Variant_Alias"))+"_"+
                            "VariantConfirmationReport"+".xlsx")

        # save the filled sheet and convert to PDF
        template.save(output_file_name)
        Fill.convert2pdf(output_file_name)
        
        # clear the dicts items
        variant_dict = {key: "-" for key in variant_dict}









produce_variant_report("report_in.txt", "/home/david/", 
                       "/home/david/configuration/ideas/autoReport/test/images/", 
                       "/home/david/configuration/ideas/autoReport/autoReport/All_Yale_&_UK_Variants.xlsx",
                       "All_Variants")
