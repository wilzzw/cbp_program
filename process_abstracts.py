##################### Created by Wilson Zeng on April 27th, 2019 ##############################

from bs4 import BeautifulSoup
import subprocess
import re
import os

##################### Mutable Inputs: Can & Should be Updated as it Goes ######################

# Available image extensions (mutable)
# This code will find image files associated with the following extensions in the current
# directory by matching the image file name against Lastname_Firstname.docx
# !Important: Potential manual work before executing the script is fixing typos in people's
# filenames. Apparently people misspell their own names nowadays.
image_ext = ['.png', '.jpg']

# Extract file names in the current working directory into tuple: (file_name, extension)
all_files = [os.path.splitext(f) for f in os.listdir('.') if os.path.isfile(f)]
doc_files = [file_tuple for file_tuple in all_files if '.docx' in file_tuple[1]]
fig_files = [file_tuple for file_tuple in all_files if file_tuple[1] in image_ext]

# Sorted by submitters' last names alphabetically
doc_files = sorted(doc_files, key=lambda x: x[0])

# Code output will be written to output_tex
output_tex = 'output.tex'

# Clear current output file if available
latex = open(output_tex, 'w')
latex.close()

# Sanity Zone; please update as you go..
# As far as I am aware, PANDOC processes '&' and '%' correctly.
# Do not double-process them by adding to the dictionary special_characters!!
greek = {'α': r'$\alpha$', 'β': r'$\beta$', 'γ': r'$\gamma$', 'δ': r'$\delta$',
         'ε': r'$\epsilon$', 'ζ': r'$\zeta$', 'η': r'$\eta$', 'θ': r'$\theta$',
         'ι': r'$\iota$', 'κ': r'$\kappa$', 'λ': r'$\lambda$', 'μ': r'$\mu$',
         'ν': r'$\nu$', 'ξ': r'$\xi$', 'π': r'$\pi$', 'ρ': r'$\rho$', 'σ': r'$\sigma$',
         'τ': r'$\tau$', 'υ': r'$\upsilon$', 'φ': r'$\phi$', 'χ': r'$\chi$', 'ψ': r'$\psi$', 'ω': r'$\omega$'}
latin_with_accents = {'é': r'\'{e}', 'è': r'\`{e}', 'ü': r'\"{u}', 'ä': r'\"{a}', 'ö': r'\"{o}'}

# The symbol I know PANDOC does not handle properly
fix_pandoc = {r'\textasciitilde{}': '$/sim$' # PANDOC gives '\textasciitilde{}' for tilde, bit it ends up being the tilde that is too high up
             }
# PANDOC handles most special characters already, except "~", which is addressed in fix_pandoc (2019-04-28)
# If there are any other special characters PANDOC didn't parse, might want to add it to special_characters.
# For the ones which PANDOC process into something not desired, might want to add it to fix_pandoc instead.
special_characters = {}

# Pool all fixes together. These fixes will be applied after we have allowed PANDOC to do its job (i.e. To clean up what PANDOC did not do a great job in).
minor_fixes = {**greek, **latin_with_accents, **fix_pandoc, **special_characters} 

# Name segments that should not be capitalized
# Anything else? Please update :)
do_not_capitalize = ['van', 'der', 'van\'t']

################################## Functions! Functions! ######################################
# General conversion of text extracted from html into latex format
# First let PANDOC handles the dirty work
# Then apply our specific fixes
# text: str, the text to be converted
# Returns: str, the input text in latex format
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

# Function to be called to write the beginning & end of the main abstract latex environment
def choose_template(references=False, figure=False):
    if not references or figure:
        return {'begin': r'\begin{posterabs}', 'end': r'\end{posterabs}'}
    if not references:
        return {'begin': r'\begin{posterabswfig}', 'end': r'\end{posterabswfig}'}
    if not figure:
        return {'begin': r'\begin{posterabswref}', 'end': r'\end{posterabswref}'}
    return {'begin': r'\begin{posterabswrefwfig}', 'end': r'\end{posterabswrefwfig}'}

# Function to be called to write abstract tile in latex
def proc_title(title=''):
    return '{'+title+'}'

# Function to be called to process an author's information (variable 'author')
# An author information is a tuple: (presenting, First_name, ..., Last_name)
# If the author is the presenting author, presenting == True, and vice versa
# Returns the name string of the author separated by spaces, as well as whether the author is presenting
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

# Function to be called to process a list of authors' information and the corresponding affiliation numbers
# authors are a list of author information tuples; affiliations are a list of lists of corresponding affiliation numbers (as strings of integers)
# Returns: 1. str, the list of authors with annotated affiliations in latex format; 2. presenting_author's label string in the format: LastF
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

# Function to be called to produce affiliations in latex format
def proc_affil(affil_info):
    affiliation_texts = ['{']
    for n in sorted(list(affil_info.keys())):
        if n == max(list(affil_info.keys())):
            affiliation_texts.append('$^'+str(n)+'$'+textfix(affil_info[n]))
        else:
            affiliation_texts.append('$^'+str(n)+'$'+textfix(affil_info[n])+r'\\')
    affiliation_texts.append('}')
    return affiliation_texts

# Function to be called to write poster number in latex format
def proc_poster_num(poster_number='#'):
    return r'{P\#}'

# Function to be called to write lines indicating inserting figure in latex format
# width and height are in number of pixels
def proc_fig(fig_file, width, height):
    fig_texts = ['{'+fig_file+'}']
    fig_texts.append('{'+str(width/100)+'}')
    fig_texts.append('{'+str(height/100)+'}')
    return fig_texts

# Function to be called to produce references in latex format
def proc_ref(ref_info):
    reference_texts = ['{']
    for n in sorted(list(ref_info.keys())):
        if n == max(list(ref_info.keys())):
            reference_texts.append('{['+str(n)+']} '+ref_info[n])
        else:
            reference_texts.append('{['+str(n)+']} '+ref_info[n]+r'\\')
    reference_texts.append('}')
    return reference_texts

# Function to be called to write the latex file output
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

    #Some people forget to indicate presenting author...
    if len(presenter_label) == 0:
        #Assume filename is the presenter
        file_id[0].split('_')
        presenter_label = file_id[0].split('_')[0].capitalize() + file_id[0].split('_')[1].capitalize()[0]

    write_line(r'\label'+'{'+presenter_label+'}')
    write_line(choose_template(ref_avail, fig_avail)['end'])

    presenting_author = [author[1:] for author in authors if author[0]]
    #Some people forget to indicate presenting author...
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

# docx to html by PANDOC. The way the subprocess.run() is written in this code requires this python code to be run under Unix Shell, such as Bash.
for file_id in doc_files:
    ##################### Part I. Parsing Word File into HTML #####################
    input_doc = ''.join(file_id)
    print('Currently working on '+input_doc)
    html_name = file_id[0]+'.html'
    subprocess.run(['pandoc', input_doc, '-f', 'docx', '-t', 'html', '-o', html_name])

    # Look whether submitter submitted a figure file
    matching_figures = [''.join(fig_id) for fig_id in fig_files if fig_id[0].lower() == file_id[0].lower()]
    if len(matching_figures) > 0:
        fig_name = matching_figures[0]
    else:
        fig_name = ''
    if len(fig_name) > 0:
        print('Submitter has a figure: '+fig_name+'\n')

    # Parse html file into BeautifulSoup :) for ease of extracting contents
    # See https://www.crummy.com/software/BeautifulSoup/bs4/doc/ for Beautiful Soup docs :)
    f = open(html_name, 'r')
    soup = BeautifulSoup(f, 'html.parser')

    # Extract all table/cell contents into a list of table objects
    all_tables = soup.find_all('table')

    # If someone messed up the document so badly/did not use the submission template
    if len(all_tables) == 0:
        print(input_doc+' is problematic. No tables were found. Please check the docx file manually as the submitter likely did not use the provided abstract template.\n')
        continue

    # The first table has the abstract title (supposedly..)
    abstract_title_components = [textfix(str(c)) for c in all_tables[0].find_all(['th', 'td'])[0].contents]
    abstract_title = ''.join(abstract_title_components) ###Processed abstract title in latex format!

    # The second table has the the authors & affiliations (supposedly..)
    # Extract all fields ('tr') in the authors & affiliations table
    author_and_affil_fields = all_tables[1].find_all('tr')
    author_list = []
    associated_affil = []
    for entry in author_and_affil_fields[1:]:
        num_author = int(re.sub('[^0-9]', '', entry.find_all('td')[0].text))
        author_name = entry.find_all('td')[1].text
        # If no author in this line, skip over
        if len(author_name) == 0:
            continue
        author_name = author_name.split(' ')
        presenting = len(entry.find_all('td')[1].find_all('em')) > 0
        author_affil = [a for a in re.sub('[^0-9]', ' ', entry.find_all('td')[2].text).split(' ') if a != '']
        author_info = tuple([presenting] + author_name)
        author_list.append(author_info)
        associated_affil.append(author_affil)
    # The format of the author_list and associated_affil: author informations are tuples: tuple(presenting?, from first name to last name); affiliations are lists of corresponding affiliation numbers as strings (not integers).

    # The third table has the affiliation names
    # Extract all fields ('tr') in the affiliations
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

    # The fourth table has the abstract body text
    abstract_content = [textfix(str(c)) for c in all_tables[3].find_all(['th', 'td'])[0].contents]
    abstract_main = ''.join(abstract_content) ###Processed abstract text in latex format!

    # The fifth table has the references
    # Extract all fields ('tr') in the references

    if len(all_tables) > 4: #People who delete tables in the abstract template: I hate you. You are annoying!
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

# Remove temporary html files
html_files = [f for f in os.listdir('.') if os.path.isfile(f) and os.path.splitext(f)[1] == '.html']
for htm in html_files:
    os.remove(htm)