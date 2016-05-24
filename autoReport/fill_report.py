import openpyxl
import re
import sys
import os
import subprocess

class FillReport(object):
    
    def __init__(self, template, template_sheet, var_alias, variant_dict, image_dir):

        self.template = template
        self.template_sheet = template_sheet
        self.var_alias = var_alias
        self.variant_dict = variant_dict
        self.image_dir = image_dir
        
        # create a variable and intialiase so all methods can use it
        category = self.variant_dict.get("Category")
        self.category = category

    def fill_report(self):
        ''' Append the information placed into the variant_dict dict into the 
            template report
        '''
        # construct mutation frequency and allele depth deatils 
        mutation_frequency = (re.sub("%", "", str(self.variant_dict.get
                                  ("Allele_Frequency_ESP(%)"))+"       "+
                              str(self.variant_dict.get
                                  ("Allele_Frequency_ExAC(%)"))+"       "+
                              str(self.variant_dict.get
                                  ("Allele_Frequency_dbSNP(%)"))))
        
        allele_depth = (str(self.variant_dict.get("Allele_Depth(REF)"))+","+
                        str(self.variant_dict.get("Allele_Depth(ALT)")))

        # uses the information stored in the get_variant_information dicts to write 
        # to the template report.
        self.template_sheet["F4"] = self.variant_dict.get("Sample_Name")
        self.template_sheet["F7"] = self.variant_dict.get("Gene")
        self.template_sheet["F8"] = self.variant_dict.get("Exon_No.")
        self.template_sheet["F9"] = self.variant_dict.get("HGVSc")
        self.template_sheet["F10"] = self.variant_dict.get("HGVSp")
        self.template_sheet["F11"] = self.variant_dict.get("Variant_Position").replace(" ","")
        self.template_sheet["F12"] = round(self.variant_dict.get("Allele_Balance"),2)
        self.template_sheet["F13"] = allele_depth
        self.template_sheet["F14"] = mutation_frequency        
        self.template_sheet["F15"] = "Y"


    def insert_image(self, query, cell, width, height):
        ''' Find an image based upon the query give and insert the image into the 
            spreadsheet
        '''
        image = [self.image_dir+image for image in os.listdir(self.image_dir) 
                 if query in image]
        if image:
            resized_image = openpyxl.drawing.image.Image(image[0],size=(width,height))
            self.template_sheet.add_image(resized_image, cell)
        else:
            pass
    
    def pick_comment(self, mutation_dict=""):
        ''' fill out the comment section in a manner dependent upon the variant category
        '''
        if self.category in ("ClinVarPathogenic","HGMD"):
            self.hgmd_clinvar_comment(mutation_dict)
        elif self.category == "Gly-X-Y":
            self.glyxy_comment()
        elif self.category == "LOF":
            self.lof_comment()
        elif self.category == "Rules":
            self.template_sheet["F16"] = "Rules category, do something"
        elif self.category == "Other":
            self.template_sheet["F16"] == "Other Category"
        elif self.category == "DamagingMissense":
            self.template_sheet["F16"] == "Damaging Missense"
        elif self.category == "BenignMissense":
            self.template_sheet["F16"] == "Benign Missense"
        else:
            self.template_sheet["F16"] == "Unknown"
       
       
    def hgmd_clinvar_comment(self, mutation_dict):
        ''' Add a comment associated with the HGMD or ClinVar accession number
            found in the database to the template report
        '''    
        
        if self.variant_dict.get("Variant_Class") == "DM?":
            comment =("This mutation has been asserted as a likely disease-causing\nmuatation in the HGMD database"+"\n\nHGMD Accession: "+self.variant_dict.get("Mutation_ID")+ "\nHGMD Classification: "+self.variant_dict.get("Variant_Class")+"\n"+self.variant_dict.get("First_Publication")+"\n\nDate of Variant Class Change From DM to DM?: "+str(self.variant_dict.get("Date_Class_Change"))+"\n"+self.variant_dict.get("Variant_Class_Change"))
            self.template_sheet["F16"] = comment
        else:
            comment = "This mutation has been asserted as a disease-causing\nmuatation in the HGMD database"+"\n\nHGMD Accession: "+self.variant_dict.get("Mutation_ID")+ "\nHGMD Classification: "+self.variant_dict.get("Variant_Class")+"\n"+self.variant_dict.get("First_Publication")
            self.template_sheet["F16"] = comment


    def glyxy_comment(self):
        ''' Add a comment associated with GLY-X-Y variant_sheet to the template report
        '''
        comment =("This mutation is predicted to disrupt the collagen triple"+ 
                  " helical structure and is therefore likely to be pathogenic")
        self.template_sheet["F16"] = comment


    def lof_comment(self):
        '''Add a comment associated with LOF variant_sheet to the template report. The 
           particular comment added is dependant upon the exon number in which the 
           variant lies within
        ''' 
        transcript_id = self.variant_dict.get("HGVSc")
        hgvs = transcript_id.split(":")[1]
        exon = self.variant_dict.get("Exon_No.")
        intron = self.variant_dict.get("Intron_No.")

        if "-" in hgvs or "+" in hgvs:
            self.template_sheet["F16"] = ("Splicing variant. This will require further"+ 
                                          " investigation")
            self.template_sheet["B8"] = "Intron"
            self.template_sheet["F8"] = intron

        elif exon != "-":
            exon_num = int(exon.split("/")[0])
            exon_total = exon.split("/")[1]
            
            if exon_num in (int(exon_total)-2, int(exon_total)-1, int(exon_total)): 
                self.template_sheet["F16"] = ("This mutations is expected to produce"+
                                              "a truncated product")
            elif exon_num < int(exon_total)-2:
                self.template_sheet["F16"] = ("This mutation introduces a premature"+
                                              "stop codon and is likely to be \npathogenic")
            else:
                self.template_sheet["F16"] = ("LOF mutation present, but the"+ 
                                              "outcome cannot be determined \nwithout"+ 
                                              "exon numbering information")
        
        elif intron != "-":
            self.template_sheet["B8"] = "Intron"
            self.template_sheet["F16"] = "This mutation affects the intron"
            self.template_sheet["F8"] = intron
    
    def convert2pdf(self, input_file):
        ''' convert the xlsx file to a pdf file. 
        '''
        # supress font.config warnings displaying
        FNULL = open(os.devnull, 'w')
        renamed_file = input_file.replace("xlsx","pdf")

        if sys.platform in ("cygwin","linux2","linux"):
            subprocess.call(["ssconvert", input_file, renamed_file],
                            stdout=FNULL, stderr=subprocess.STDOUT)
        
        elif sys.platform =="win32":
            # untested code below
            from win32com import client
            xlApp = client.Dispatch("Excel.Application")
            books = xlApp.Workbooks.open(input_file)
            ws = books.Worksheets[0]
            ws.Visible = 1
            ws.ExportAsFixedFormat(0, renamed_file)
        
        else:
            print("Unrecognised system platform.\nYour system platform is " + sys.platform)

        


