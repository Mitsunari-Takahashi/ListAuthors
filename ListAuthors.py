import xlwings as xw
from operator import itemgetter, attrgetter
import string
import copy


LIST_ALPHABETS = list(string.ascii_uppercase)


def get_alphabet_code(number, offset=0):
    m = number+offset
    period = int(m/len(LIST_ALPHABETS))
    if m<len(LIST_ALPHABETS):
        code = LIST_ALPHABETS[m]
    elif m<len(LIST_ALPHABETS)*len(LIST_ALPHABETS):
        code = LIST_ALPHABETS[period] + LIST_ALPHABETS[m%len(LIST_ALPHABETS)]
        
    return code


class Entity:
    def __init__(self, data_body, data_header):
        if len(data_body)==len(data_header):
            for col, val in zip(data_header, data_body):
                setattr(self, col, val)
        else:
            print('Data and header mismatch!')

    
class Member:
    def __init__(self, personal_data, data_header):
        if len(personal_data)==len(data_header):
            for col, val in zip(data_header, personal_data):
                if col in ('Affiliations', 'Groups'):
                    if isinstance(val, int):
                        setattr(self, col, [str(val)])
                    elif isinstance(val, str):
                        if not ',' in val:
                            setattr(self, col, [val])
                        else:
                            setattr(self, col, val.split(','))
                else:
                    setattr(self, col, val)
            self.importance = 0
        else:
            print('Data and header mismatch!')

  
class Affliation:
    def __init__(self, ID, Name_E, Name_J, ShortName_E=None, ShortName_J=None):
        self.ID = ID
        self.Name_E = Name_E
        self.Name_J = Name_J
        self.ShortName_E = ShortName_E
        self.ShortName_J = ShortName_J


class Group:
    def __init__(self, ID, Name):
        self.ID = ID
        self.Name = Name


def read_entities(sheet, top_cell='A1'):
    sheet.activate()
    list_entities = []
    table = xw.Range(top_cell).expand('table').value
    table_header = table[0]
    table_body = table[1:]

    for entity_data in table_body:
        list_entities.append(Entity(entity_data, table_header))

    return list_entities


def read_members(sheet, all_groups, author_group_ids, top_cell='A1'):
    sheet.activate()
    list_members = []
    member_table = xw.Range(top_cell).expand('table').value
    member_table_header = member_table[0]
    member_table_body = member_table[1:]

    for personal_data in member_table_body:
        print(personal_data)
        member = Member(personal_data, member_table_header)
        for group_id in member.Groups:
            if int(group_id) in author_group_ids:
                list_members.append(member)
                break

    print(list_members)
    return list_members


def multisort(xs, specs):
    xs_copy = copy.deepcopy(xs)
    for key, reverse in reversed(specs):
        xs_copy.sort(key=attrgetter(key), reverse=reverse)
    return xs_copy

    
def sort_members(list_members, keys_sort=[('importance',True), ('SurName_Kana', False), ('GivenName_Kana', False)]):
    return multisort(list_members, keys_sort)


def print_authors(list_authors, list_all_affiliations, lang='Japanese', format='JPS'):
    list_str_author = []
    dict_affiliation_code = {}
    list_affiliation = []
    list_str_affiliation = []
    for author in list_authors:
        # List affiliation
        list_author_affiliation = []
        for affi in author.Affiliations:
            if not affi in dict_affiliation_code.keys():
                dict_affiliation_code[affi] = get_alphabet_code(len(dict_affiliation_code.keys()))
                for acan in list_all_affiliations:
                    if int(acan.ID)==int(affi):
                        list_affiliation.append(acan)
                        if format in ['JPS']:
                            if lang in ['Japanese']:
                                list_str_affiliation.append('{name}^{code}^'.format(name=acan.ShortName_J, code=dict_affiliation_code[affi]))
                            elif lang in ['English']:
                                list_str_affiliation.append('{name}^{code}^'.format(name=acan.ShortName_E, code=dict_affiliation_code[affi]))
                        break
            list_author_affiliation.append(dict_affiliation_code[affi])

        name_components = []
        name_lang = lang
        for alang in [lang, 'English']: # Englishe is alternative to your language
            iname_key = 1
            name_key = 'Name_Print_{0}_{1}'.format(alang, iname_key)
            while name_key in set(author.__dict__.keys()):
                name_compo = getattr(author, name_key)
                if not name_compo is None:
                    name_components.append(name_compo)
                iname_key+=1
                name_key = 'Name_Print_{0}_{1}'.format(alang, iname_key)
            if len(name_components)>0:
                name_lang = alang
                break
    
        if format in ['JPS']:
            str_author_affiliation = ','.join(sorted(list_author_affiliation))
            if name_lang in ['Japanese']:
                list_str_author.append('{name}^{affi}^'.format(name=''.join(name_components), affi=str_author_affiliation))
            elif name_lang in ['English']:
                list_str_author.append('{name}^{affi}^'.format(name=' '.join(name_components), affi=str_author_affiliation))
            
        # if format in ['JPS']:
        #     str_author_affiliation = ','.join(sorted(list_author_affiliation))

        #     # Choose a personal name style to print
        #     if lang in ['Japanese']:
        #         if author.SurName_Kanji is None:
        #             surname = author.SurName_Kana
        #         else:
        #             surname = author.SurName_Kanji if len(author.SurName_Kanji)>0 else author.SurName_Kana
        #         if author.GivenName_Kanji is None:
        #             givenname = author.GivenName_Kana
        #         else:
        #             givenname = author.GivenName_Kanji if len(author.GivenName_Kanji)>0 else author.GivenName_Kana
                    
        #         list_str_author.append('{sur}{given}^{affi}^'.format(sur=surname, given=givenname, affi=str_author_affiliation))

        #     elif lang in ['English']:
        #         list_str_author.append('{aut.GivenName_Alphabet};{aut.SurName_Alphabet}^{affi}^'.format(aut=author, affi=str_author_affiliation))

                
    return [', '.join(list_str_author), ', '.join(list_str_affiliation)]
        
        

def main():
    wb = xw.Book.caller()
    wb.sheets['Main'].activate()
    # Load your setup
    groups_author = xw.Range("Main!B2").expand('right').value
    ai = xw.Range("Main!B3").expand('right').value
    authors_important = ai if isinstance(ai, list) else [ai]
    
    # Load the whole group list
    list_groups = read_entities(wb.sheets['Group'])
    # Load the whole affiliation list
    list_affiliations = read_entities(wb.sheets['Affiliation'])
    
    # Load the whole member list
    list_members = read_members(wb.sheets['Member'], list_groups, groups_author)
    
    # For each member included in the "important author" list, increase Member.importance 
    for mem in list_members:
        if mem.ID in authors_important:
            mem.importance += 1
    sorted_members = {}
    sorted_members['Japanese'] = sort_members(list_members, keys_sort=[('importance', True), ('Name_Sort_Japanese_1', False), ('Name_Sort_Japanese_2', False)])
    sorted_members['English'] = sort_members(list_members, keys_sort=[('importance', True), ('Name_Sort_English_1', False), ('Name_Sort_English_2', False)])

    # Display the result
    wb.sheets['Main'].activate()
    # Print in JPS style
    results = {}
    for ilang, lang in enumerate(('Japanese', 'English')):
        results[lang] = print_authors(sorted_members[lang], list_affiliations, lang=lang) #list(sorted_members_jp[0].Affiliations)
        wb.sheets['Main']["B{}".format(ilang*3+6)].value = results[lang][0]
        wb.sheets['Main']["B{}".format(ilang*3+7)].value = results[lang][1]
        
if __name__ == "__main__":
    xw.Book("myproject.xlsm").set_mock_caller()
    main()
