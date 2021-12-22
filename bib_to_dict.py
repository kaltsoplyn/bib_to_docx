import re
import json

# parses a bib-tex file into a list of BibRecord bibliographic records
class BibToDictList:
    def __init__(self, filename = ''):
        self.records = import_bib(filename)

# class to hold bib records in json/dict format
# (as shaped by the import_bib function)
class BibRecord:
    def __init__(self, rec = {}):
        self.record = rec
    
    def set_record(self, rec):
        self.record = rec
    
    def append(self, rec):
        _rec = self.record
        self.record = {**_rec, **rec}
    
    def get_tag(self, tag):
        try: out = self.record[tag]
        except: return ''
        return out   

    def get_author(self, index):
        auth_list = self.record['author'].split(' and ')
        return auth_list[index-1] if index-1 < len(auth_list) else ''

    def fix_authors(self):
        auth_list = self.record['author'].split(', and ')
        if len(auth_list) == 1: auth_list = self.record['author'].split(' and ')
        for i in range(len(auth_list)):
            auth = auth_list[i].split(', ')
            auth_name = re.sub(r'^(\w)[\w]*?(\s|\-)(\w)[\w]*', r'\1\2\3', auth[1]).replace(' ', '').replace('.', '')
            auth_list[i] = ', '.join([auth[0].title(), re.sub('[a-z]*', r'', auth_name)])
        self.record['author'] = ', and '.join(auth_list)

# function to import a bib-tex file and convert it into a list of BibRecords
def import_bib(filename):
    try:    
        # collect all bib entries to records list
        with open(filename) as bib_file:
            bib = bib_file.read()
            bib = bib[1:]
            reg = re.compile(r'^@[\w\W\s]*?,\n}', re.MULTILINE)
            records = reg.findall(bib)

        # convert records to json, after a ridiculous amount of reg exp substitutions
        for i in range(len(records)):
            records[i] = re.sub(r'Abstract =[\w\W\d\s]*?},', r'', records[i])
            records[i] = re.sub(r'Funding-Ack[\w\W\d\s]*?},', r'', records[i])
            records[i] = re.sub(r'Funding-Text[\w\W\d\s]*?}},', r'', records[i])
            records[i] = re.sub("{\[}", "[", re.sub("{\]}", "]", records[i]) )   # convert double stupid {[} and {]}, to [, ]
            records[i] = re.sub("}}", "}", re.sub("{{", "{", records[i]) )   # convert double {{, }} to single {, }
            records[i] = re.sub(r'ISI:[\d\w]*?,', r'"ref_id" : "' + str(i + 1) + '", ', records[i])   # remove the ISI entry at the beginning / replace it with an id
            records[i] = re.sub(r'WOS:[\d\w]*?,', r'"ref_id" : "' + str(i + 1) + '", ', records[i])   # remove the WOS entry at the beginning / replace it with an id
            records[i] = re.sub(r'\n([\w\-\s]+?) = ', lambda pat: '"' +  pat.group(1).lower() + '" :', records[i])   # parse keys from e.g. ' Author = ' to ' "author" : '
            records[i] = re.sub(r',\n}', r'>>>', re.sub(r'@article{', r'<<<', records[i]))   # enclose whole object in distinct <<<, >>> instead of {, }
            records[i] = re.sub(r',\n}', r'>>>', re.sub(r'@incollection{', r'<<<', records[i]))   # enclose whole object in distinct <<<, >>> instead of {, }
            records[i] = re.sub(r'\n', r' ', records[i])   # remove newlines
            records[i] = re.sub(r'{(.*?)}', r'"\1"', records[i])   # substitute {field} with "field"
            records[i] = re.sub(r'\\', r'\\\\', records[i])   # escape possible backslashes \
            records[i] = re.sub(r'<<<', r'{', re.sub(r'>>>', r'}', records[i]))   # go back from <<<, >>> to {, }
            try:
                records[i] = json.loads(records[i])   # now each record should be in json string format > parse into dicts
            except json.decoder.JSONDecodeError as e:
                print("JSON Error at index {}: {}\n".format(i, e))
                print("Erroneous parsed record at that index (inspect record for errors):\n{}\n".format(records[i]))
                records[i] = {"Title": "JSON Error at index {}: {}".format(i, e), "Author": "json.decoder.jsonDecoderError"}

        # homogenize titles into .title() format
        for rec in records:
            rec['title'] = (re.sub(r'[ ]+', r' ', rec['title'])).title()   # remove excessive whitespace resulting from previously removing newlines
            rec['author'] = (re.sub(r'[ ]+', r' ', rec['author']))         # remove excessive whitespace resulting from previously removing newlines
            try:
                if 'gold' in rec['oa'].lower() or 'hybrid' in rec['oa'].lower() :
                    rec['oa'] = 'gold'
                elif 'green' in rec['oa'].lower() :
                    rec['oa'] = 'green'
            except:
                rec['oa'] = ''

        # finally, convert all dicts in the list into BibRecord instances
        for i in range(len(records)):
            records[i] = BibRecord(records[i])
            records[i].fix_authors()

        return records

    except:
        print('Parsing failed: parsing error, empty/non-bib file or IO error in filename, "{}"'.format(filename))
        return []