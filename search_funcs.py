import polars as pl
from datetime import datetime

def search_method_str(token, doc_ref, heading_ref = "/Users/rebeccakrall/Data/Proposal-Report-Data/Final_References/headings.csv"):
    ref = pl.read_csv(heading_ref)
    ref = (ref.with_columns([pl.col('heading').str.to_lowercase()]))

    filtered_documents = ref.filter(pl.col('heading').str.contains(token.lower()))\
                        .select('document_id')

    
    doc_list = filtered_documents['document_id'].sort().to_list()
    
    if len(doc_list):
        return doc_ref.filter(pl.col('document_id').str.contains_any(doc_list))
    else:
        return None


def find_matching_methods(method_list, doc_ref, method_ref = "/Users/rebeccakrall/Data/Proposal-Report-Data/Final_References/methods.csv" ):
    ref = pl.read_csv(method_ref)

    filtered_documents = ref.group_by('document_id').agg(pl.col('method')) \
                        .filter(pl.col('method').map_elements(lambda methods: all(method in methods for method in method_list))) \
                        .select('document_id')

    doc_list = filtered_documents['document_id'].sort().to_list()
    # return doc_ref.filter(pl.col('document_id').str.contains_any(doc_list))

    if len(doc_list):
        return doc_ref.filter(pl.col('document_id').str.contains_any(doc_list))
    else:
        return None


def find_matching_compounds(compound_list, doc_ref, compound_ref = "/Users/rebeccakrall/Data/Proposal-Report-Data/Final_References/common_compound_matches.csv"):
    ref = pl.read_csv(compound_ref)
    filtered_documents = ref.group_by('document_id').agg(pl.col('compound')) \
                        .filter(pl.col('compound').map_elements(lambda methods: all(method in methods for method in compound_list))) \
                        .select('document_id')

    doc_list = filtered_documents['document_id'].sort().to_list()
    # return doc_ref.filter(pl.col('document_id').str.contains_any(doc_list))
    if len(doc_list):
        return doc_ref.filter(pl.col('document_id').str.contains_any(doc_list))
    else:
        return None



def find_date_range(year, doc_ref):

    doc_ref = doc_ref.with_columns(doc_ref['study_date'].str.to_date('%m-%d-%y'))
    if type(year) == tuple:
        return doc_ref.filter(pl.col('study_date').is_between(datetime(year[0], 1, 1), datetime(year[1], 12, 31))).sort(by = 'study_date')
    else:
        return doc_ref.filter(pl.col('study_date').is_between(datetime(year, 1, 1), datetime(year, 12, 31))).sort(by = 'study_date')
    

def find_client_match(client, doc_ref):
    return doc_ref.filter(pl.col('client') == client)



def get_possible_methods():
    methods = pl.read_excel("/Users/rebeccakrall/Data/Proposal-Report-Data/Reference_Data/Assay_List.xlsx")
    return methods['Assay'].unique().sort().to_list()

def get_possible_clients():
    clients = pl.read_excel("/Users/rebeccakrall/Data/Proposal-Report-Data/Final_References/Final_Client_Codes.xlsx")
    return clients['client'].sort().to_list()

def get_possible_compounds():
    cmpds = pl.read_excel("/Users/rebeccakrall/Data/Proposal-Report-Data/Final_References/Melior common drugs.xlsx")
    return cmpds['Compound'].sort().to_list()

def get_possible_years(doc_ref):
    # doc_ref = pl.read_excel("/Users/rebeccakrall/Data/Proposal-Report-Data/Final_References/all_filename_detail_050442.xlsx")
    doc_ref = doc_ref.with_columns(doc_ref['study_date'].str.to_date('%m-%d-%y'))

    # years = doc_ref['study_date'].dt.year().unique().sort().to_list()
    return doc_ref['study_date'].dt.year().min(), doc_ref['study_date'].dt.year().max()