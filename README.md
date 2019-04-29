# cbp_program
This repo contains the script used to process attendee abstract submissions.
The script has 2 main actions in series:
(1) Convert docx into html by PANDOC because Microsoft Word is a pain in the ***
(2) Extract information neatly from html by Beautiful Soup, convert to proper latex syntax and write it to latex output.

# To use this code... #
Please ensure you have python 3.5 and above installed on your Unix shell. For Windows 10 users, the equivalent command-line interpreter is Bash. For older versions of Windows, I don't know enough but maybe it is necessary to install GitBash.
Execute the python script in the directory with the abstract files.

# Manual work before executing the script #
1. If there are any, save all doc files to docx in the current working directory using Microsoft Word. DOC files incredibly outdated and people should stop using it :(
2. Fix any typos in the submission files. i.e. The image files should have matching docx files. Apparently some competent people can't spell their names nowadays.
3. The program should skip over submission files that do not follow the template provided (instead of failing and stalling). The easiest fix to that is to manually place the contents into the abstract_template docx file.

# Manual work after executing the script #
Put in the figure dimensions in pixels/100. This feature has not been added.

# Future update wish-list #
1. Read figure dimensions from figure files and directly parse them.

This is the first version. Please report any problems and bugs to wilsonzzw@gmail.com, or send a pull request.
(Last updated 2019-04-29)
