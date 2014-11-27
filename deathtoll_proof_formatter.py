__author__ = 'sabo'
import sys
from docx import opendocx, getdocumenttext

global panel_counter
global current_page

PAGE_FORMAT_OPEN = '\\par\\par{\\rtlch\\fcs1 \\af0\\afs24 \\ltrch\\fcs0 \\f39\\fs24\\highlight3\\insrsid8093251\\charrsid92860 \\hich\\af39' \
                    '\\dbch\\af31505\\loch\\f39 Page '
PAGE_FORMAT_CLOSE = '}{\\rtlch\\fcs1 \\af0\\afs24 \\ltrch\\fcs0 \\f39\\fs24\\highlight3\\insrsid8093251\\charrsid9795096 \\hich' \
                    '\\af39\\dbch\\af31505\\loch\\f39 :}{' \
                    '\\rtlch\\fcs1 \\af0\\afs24 \\ltrch\\fcs0 \\f39\\fs24\\insrsid8093251\\charrsid92860 \\hich\\af39\\dbch\\af31505' \
                    '\\loch\\f39  }{\\rtlch\\fcs1 \\af0\\afs24 \\ltrch\\fcs0 \\f39\\fs24\\cf6\\insrsid8093251\\charrsid92860 \\hich' \
                    '\\af39\\dbch\\af31505\\loch\\f39 ======================================\\line }{\\rtlch\\fcs1 \\af0\\afs24' \
                    '\\ltrch\\fcs0 \\f39\\fs24\\lang2057\\langfe1041\\langnp2057\\insrsid8093251\\charrsid10886310\\par}'
PANEL_FORMAT_OPEN = '\\par{\\rtlch\\fcs1 \\af0\\afs24 \\ltrch\\fcs0 \\f39\\fs24\\highlight4\\insrsid8093251\\charrsid16219165 \\hich\\af39' \
                    '\\dbch\\af31505\\loch\\f39 Panel '
PANEL_FORMAT_CLOSE = ':}'
PANEL_DELIMITER = '----'
PAGE_DELIMITER = 'page '


def get_trans_txt(path):
    document = opendocx(path)
    document_txt = getdocumenttext(document)

    para_text_list = []
    for para_text in document_txt:
        para_text_list.append(para_text.encode("utf-8"))

    as_txt = '\n\n'.join(para_text_list)

    return as_txt


def replace_page_annotations(text):
    return PAGE_FORMAT_OPEN + text.split(' ')[1] + PAGE_FORMAT_CLOSE


def replace_panel_annotations(text):
    return PANEL_FORMAT_OPEN + str(panel_counter) + PANEL_FORMAT_CLOSE


if __name__ == "__main__":
    in_path = opendocx(sys.argv[1])
    #in_path = "C:\\Users\\steven\\Desktop\\Deathtoll\\diamond_cut_diamond\\19\\Dcd19_trans.docx"
    rtf_out_file = open(sys.argv[2], 'w')
    #rtf_out_file = open('rtf_out.rtf', 'w')

    rtf_header_file = open('rtf_header.txt', 'r')
    out_rtf = rtf_header_file.read()

    txt_formatted_txt = get_trans_txt(in_path)

    formatted = []
    panel_counter = 0
    current_page = 0
    for line in txt_formatted_txt.split('\n'):
        if line.lower().startswith('page '):
            print 'PAGE FOUND %s' % line
            current_page = line.split(' ')[1].strip()
            panel_counter = 0
            formatted.append(replace_page_annotations(line))
            panel_counter += 1
            formatted.append(replace_panel_annotations(line))
            continue
        elif line.strip() == PANEL_DELIMITER:
            print 'PANEL FOUND %s' % line
            panel_counter += 1
            formatted.append(replace_panel_annotations(line))
            continue
        print 'TEXT FOUND %s' % line
        formatted.append(line + "\\par  ")

    rtf_out_file.write(out_rtf)

    for line in formatted:
        rtf_out_file.write(line + '\r\n')