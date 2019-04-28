from bs4 import BeautifulSoup
import subprocess
import re
import os

image_ext = ['.png', '.jpg']

all_files = [os.path.splitext(f) for f in os.listdir('.') if os.path.isfile(f)]
doc_files = [file_tuple for file_tuple in all_files if '.docx' in file_tuple[1]]
fig_files = [file_tuple for file_tuple in all_files if file_tuple[1] in image_ext]

#Ensure sorted by first author's last name
doc_files = sorted(doc_files, key=lambda x: x[0])

output_tex = 'output.tex'
#Clear output file
latex = open(output_tex, 'w')
latex.close()

#Sanity Zone; pllease update as you go..
greek = {'α': '$/alpha$', 'β': '$/beta$', 'γ': '$/gamma$', 'δ': '$/delta$',
         'ε': '$/epsilon$', 'ζ': '$/zeta$', 'η': '$/eta$', 'θ': '$/theta$',
         'ι': '$/iota$', 'κ': '$/kappa$', 'λ': '$/lambda$', 'μ': '$/mu$',
         'ν': '$/nu$', 'ξ': '$/xi$', 'π': '$/pi$', 'ρ': '$/rho$', 'σ': '$/sigma$',
         'τ': '$/tau$', 'υ': '$/upsilon$', 'φ': '$/phi$', 'χ': '$/chi$', 'ψ': '$/psi$', 'ω': '$/omega$'}
latin_with_accents = {'é': r'\'{e}', 'è': r'\`{e}', 'ü': r'\"{u}', 'ä': r'\"{a}', 'ö': r'\"{o}'}
fix_pandoc = {r'\textasciitilde{}': '$/sim$'}
special_characters = {} #Pandoc handles most special characters already, except "~", which is addressed in fix_pandoc

minor_fixes = {**greek, **latin_with_accents, **fix_pandoc, **special_characters}

#Name segments that should not be capitalized
do_not_capitalize = ['van', 'der', 'van\'t']

##################### Functions #####################
def textfix(text):
    #return text
    if '\n' in text:
        paragraphs = text.split('\n')
    else:
        paragraphs = [text]
    fixed_paragraphs = []
    for p in paragraphs:
        s = subprocess.run(['pandoc', '-f', 'html', '-t', 'latex'], input=p, stdout=subprocess.PIPE, universal_newlines=True)
        fixed_text = s.stdout[:-1]
        for fix in minor_fixes.keys():
            fixed_text = fixed_text.replace(fix, minor_fixes[fix])
        fixed_text = ' '*(len(p)-len(p.lstrip())) + fixed_text + ' '*(len(p)-len(p.rstrip())) #Side effect of pandoc is that it strips off the leading/trailing white spaces
        fixed_text = fixed_text.replace('\n', ' ')
        fixed_paragraphs.append(fixed_text)
    total_fixed_text = r'\\'.join(fixed_paragraphs)
    return total_fixed_text

def choose_template(references=False, figure=False):
    if not references or figure:
        return {'begin': r'\begin{posterabs}', 'end': r'\end{posterabs}'}
    if not references:
        return {'begin': r'\begin{posterabswfig}', 'end': r'\end{posterabswfig}'}
    if not figure:
        return {'begin': r'\begin{posterabswref}', 'end': r'\end{posterabswref}'}
    return {'begin': r'\begin{posterabswrefwfig}', 'end': r'\end{posterabswrefwfig}'}

def proc_title(title=''):
    return '{'+title+'}'

def proc_author_info(author):
    presenting = author[0]
    name_components = []
    for name in author[1:]:
        if name in do_not_capitalize:
            name_components.append(textfix(name))
        else:
            name_components.append(textfix(name.capitalize()))
    if presenting:
        return {'name': r'\underline{'+' '.join(name_components)+'}', 'presenting?': presenting}
    return {'name': ' '.join(name_components), 'presenting?': presenting}

def proc_authors(authors, affiliations):
    author_tex = []
    label = ''
    for author, assoc_affil in zip(authors, affiliations):
        author_string = proc_author_info(author)['name']+',$^{'+','.join(assoc_affil)+'}$'
        author_tex.append(author_string)
        if proc_author_info(author)['presenting?']:
            label = author[-1].capitalize() + author[1][0].capitalize()
    author_tex_string = '{'+' '.join(author_tex)+'}'
    return {'tex': author_tex_string, 'label': label}

def proc_affil(affil_info):
    affiliation_texts = ['{']
    for n in sorted(list(affil_info.keys())):
        if n == max(list(affil_info.keys())):
            affiliation_texts.append('$^'+str(n)+'$'+textfix(affil_info[n]))
        else:
            affiliation_texts.append('$^'+str(n)+'$'+textfix(affil_info[n])+r'\\')
    affiliation_texts.append('}')
    return affiliation_texts

def proc_poster_num(poster_number='#'):
    return r'{P\#}'

def proc_fig(fig_file, width, height):
    fig_texts = ['{'+fig_file+'}']
    fig_texts.append('{'+str(width/100)+'}')
    fig_texts.append('{'+str(height/100)+'}')
    return fig_texts

def proc_ref(ref_info):
    reference_texts = ['{']
    for n in sorted(list(ref_info.keys())):
        if n == max(list(ref_info.keys())):
            reference_texts.append('{['+str(n)+']} '+ref_info[n])
        else:
            reference_texts.append('{['+str(n)+']} '+ref_info[n]+r'\\')
    reference_texts.append('}')
    return reference_texts

def write_latex(output_name, title, authors, associated_affiliations, affiliations_list, abstract_text, figure='', width=0, height=0, references=[]):
    if output_name[-4:] != '.tex' and '.' not in output_name:
        output_name += '.tex'
    ref_avail = len(references) > 0
    fig_avail = len(figure) > 0

    latex = open(output_name, 'a')
    def write_line(line=''):
        latex.write(line+'\n')
        return
    write_line(choose_template(ref_avail, fig_avail)['begin'])
    write_line(proc_title(title))
    write_line(proc_authors(authors, associated_affiliations)['tex'])
    for affil in proc_affil(affiliations_list):
        write_line(affil)
    write_line(proc_poster_num())
    if ref_avail:
        for ref in proc_ref(references):
            write_line(ref)
    if fig_avail:
        for l in proc_fig(figure, width, height):
            write_line(l)
    write_line(abstract_text)
    
    presenter_label = proc_authors(authors, associated_affiliations)['label']

    #Some people forget to indicate presenting author wtf...
    if len(presenter_label) == 0:
        #Assume filename is the presenter
        file_id[0].split('_')
        presenter_label = file_id[0].split('_')[0].capitalize() + file_id[0].split('_')[1].capitalize()[0]

    write_line(r'\label'+'{'+presenter_label+'}')
    write_line(choose_template(ref_avail, fig_avail)['end'])

    presenting_author = [author[1:] for author in authors if author[0]]
    #Some people forget to indicate presenting author wtf...
    if len(presenting_author) == 0:
        presenting_author = [tuple(file_id[0].split('_'))]

    presenting_author = [name.capitalize() for name in presenting_author[0] if name not in do_not_capitalize]
    presenting_name_components = []
    for name in presenting_author:
        if name in do_not_capitalize:
            presenting_name_components.append(name)
        else:
            presenting_name_components.append(name.capitalize())
    presenting_author = ' '.join(presenting_name_components)

    phantom = r'\phantomsection\addcontentsline{toc}{subsection}{\hyperref['+presenter_label+']'+r'{\textbf{'+presenting_author+'}'+r'\\'+proc_title(title)[1:-1]+'}'+'}'
    write_line(phantom)
    write_line()
    latex.close()
    return

#docx to html by Pandoc. Below commands should be run under Linux/Unix Shell.
for file_id in doc_files:
    ##################### Part I. Parsing Word File into HTML #####################
    input_doc = ''.join(file_id)
    print('Currently working on '+input_doc)
    html_name = file_id[0]+'.html'
    subprocess.run(['pandoc', input_doc, '-f', 'docx', '-t', 'html', '-o', html_name])

    #Look whether submitter submitted a figure file
    matching_figures = [''.join(fig_id) for fig_id in fig_files if fig_id[0].lower() == file_id[0].lower()]
    if len(matching_figures) > 0:
        fig_name = matching_figures[0]
    else:
        fig_name = ''
    if len(fig_name) > 0:
        print('Submitter has a figure: '+fig_name+'\n')

    #Parse html file into BeautifulSoup :) for ease of extracting contents
    #See https://www.crummy.com/software/BeautifulSoup/bs4/doc/ for Beautiful Soup docs :)
    f = open(html_name, 'r')
    soup = BeautifulSoup(f, 'html.parser')

    #Extract all table/cell contents into a list of table objects
    all_tables = soup.find_all('table')

    #If someone messed up
    if len(all_tables) == 0:
        print(input_doc+' is problematic. No tables were found. Please check the docx file manually as the submission format is likely wrong.\n')
        continue

    #The first table has the abstract title (supposedly..)
    abstract_title_components = [textfix(str(c)) for c in all_tables[0].find_all(['th', 'td'])[0].contents]
    abstract_title = ''.join(abstract_title_components) ###Processed abstract title in latex format!

    #The second table has the the authors & affiliations (supposedly..)
    #Extract all fields ('tr') in the authors & affiliations table
    author_and_affil_fields = all_tables[1].find_all('tr')
    author_list = []
    associated_affil = []
    for entry in author_and_affil_fields[1:]:
        num_author = int(re.sub('[^0-9]', '', entry.find_all('td')[0].text))
        author_name = entry.find_all('td')[1].text
        #If no author in this line, skip over
        if len(author_name) == 0:
            continue
        author_name = author_name.split(' ')
        presenting = len(entry.find_all('td')[1].find_all('em')) > 0
        author_affil = [a for a in re.sub('[^0-9]', ' ', entry.find_all('td')[2].text).split(' ') if a != '']
        author_info = tuple([presenting] + author_name)
        author_list.append(author_info)
        associated_affil.append(author_affil)
    #The format of the author_list and associated_affil: author informations are tuples: tuple(presenting?, from first name to last name); affiliations are lists of corresponding affiliation numbers as strings (not integers).

    #The third table has the affiliation names
    #Extract all fields ('tr') in the affiliations
    affiliation_fields = all_tables[2].find_all('tr')
    affil_dict = {}
    for entry in affiliation_fields:
        affil_num = int(re.sub('[^0-9]', '', entry.find_all(['td', 'th'])[0].text))
        affil = entry.find_all(['td', 'th'])[1]
        if affil.p != None:
            affil = affil.p.text
        else:
            affil = affil.text
        if len(affil) > 0:
            affil_dict[affil_num] = affil

    #The fourth table has the abstract body text
    abstract_content = [textfix(str(c)) for c in all_tables[3].find_all(['th', 'td'])[0].contents]
    abstract_main = ''.join(abstract_content) ###Processed abstract text in latex format!

    #The fifth table has the references
    #Extract all fields ('tr') in the references

    if len(all_tables) > 4: #People who delete their tables: I hate you. You are annoying!
        reference_fields = all_tables[4].find_all('tr')
        ref_list = {}
        for entry in reference_fields:
            ref_num = int(re.sub('[^0-9]', '', entry.find_all(['td', 'th'])[0].text))
            ref = entry.find_all(['td', 'th'])[1]
            fancy_texts = [str(x) for x in ref.find_all(['em', 'strong', 'sup'])]
            raw_ref_text = entry.find_all(['td', 'th'])[1]
            raw_ref_text = ''.join([textfix(str(c)) for c in raw_ref_text.contents])
            if len(raw_ref_text) > 0:
                ref_list[ref_num] = raw_ref_text

    f.close()

    ##################### Part II. Writing TEX File #####################
    write_latex(output_tex, abstract_title, author_list, associated_affil, affil_dict, abstract_main, figure=fig_name, width=0, height=0, references=ref_list)
    print('\n')

html_files = [f for f in os.listdir('.') if os.path.isfile(f) and os.path.splitext(f)[1] == '.html']
for htm in html_files:
    os.remove(htm)