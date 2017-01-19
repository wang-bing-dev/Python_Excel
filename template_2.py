import os
import tempfile
import zipfile
import shutil
from lxml import etree
from get_info_excel import *


updated_xml_content = ''
data_list = []


def make_client_insured_info():
    """
    Get updated info from excel file.
    :return:
    """
    global data_list
    data_lists = read_from_excel()
    data_list = data_lists[0]


def get_word_xml(docx_filename):
    """
    Get xml content from source docx file.
    :param docx_filename:
    :return:
    """
    global zip
    zip = zipfile.ZipFile(docx_filename)
    xml_content = zip.read('word/document.xml')
    return xml_content


def get_footer2_xml(docx_filename):
    """
    Get xml content from word/footer2.xml.
    :param docx_filename:
    :return:
    """
    zip = zipfile.ZipFile(docx_filename)
    xml_content = zip.read('word/footer2.xml')
    return xml_content


def get_footer1_xml(docx_filename):
    """
    Get xml content from word/footer1.xml
    :param docx_filename:
    :return:
    """
    zip = zipfile.ZipFile(docx_filename)
    xml_content = zip.read('word/footer1.xml')
    return xml_content


def get_xml_tree(xml_string):
    """
    Get xml string from xml file.
    :param xml_string:
    :return:
    """
    return etree.fromstring(xml_string)


def get_xml_string(xml_tree):
    """
    Get xml tree string from xml content.
    :param xml_tree:
    :return:
    """
    return etree.tostring(xml_tree, pretty_print=True)


def iter_text(xml_tree):
    """Iterator to go through xml tree's text nodes"""
    for node in xml_tree.iter(tag=etree.Element):
        if _check_element_is(node, 't'):
            # print node.text
            yield (node)


def _check_element_is(element, type_char):
    """
    Check tagname <w:t> in xmltree.
    :param element:
    :param type_char:
    :return:
    """
    word_schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return element.tag == '{%s}%s' % (word_schema, type_char)


def update_xml_content(xml_tree):
    """
    Modify the document content.
    :param xml_tree:
    :return:
    """
    global updated_xml_content
    for node in iter_text(xml_tree):
        if node.text == '#CLIENT':
            node.text = data_list['CLIENT INFORMATION/CLIENT'].upper()
        elif node.text == '#CLIENT_ADDRESS':
            node.text = data_list['CLIENT INFORMATION/ADDRESS 1'].upper()
        elif node.text == '#CLIENT_SUITE':
            node.text = data_list['CLIENT INFORMATION/STE'].upper()
        elif node.text == '#CLIENT_CITY_ZIPCODE':
            node.text = data_list['CLIENT INFORMATION/ZIPCODE'].upper()
        elif node.text == '#CONTACT_NAME':
            node.text = data_list['CLIENT INFORMATION/CONTACTFIRST NAME'].upper() + ' ' + data_list['CLIENT INFORMATION/CONTACTLAST NAME'].upper()
        elif node.text == '#CONTACT_PHONE':
            node.text = data_list['CLIENT INFORMATION/PHONE #'].upper()
        elif node.text == '#Claim No':
            node.text = data_list['CLAIM/PO #'].upper()
        elif node.text == '#Date of Loss':
            node.text = data_list['INSPDATE'].upper()
        elif node.text == '#INSURED':
            node.text = data_list['INSURED INFORMATION/INSURED'].upper()
        elif node.text == '#LOSS LOCATION':
            node.text = data_list['INSURED INFORMATION/LOSS LOCATION'].upper()
        elif node.text == '#CITY_ZIPCODE':
            node.text = data_list['INSURED INFORMATION/CITY'].upper() + ', ' + data_list['INSURED INFORMATION/ZIPCODE'].upper()
        elif node.text == '#SEX':
            node.text = data_list['CLIENT INFORMATION/sex'].upper()
        elif node.text == '#NAME':
            node.text = data_list['CLIENT INFORMATION/CONTACTFIRST NAME'].upper() + ' ' + data_list['CLIENT INFORMATION/CONTACTLAST NAME'].upper()
        else:
            pass

    return xml_tree


def update_footer_xml(xml_tree):
    """
    Modify the footer infomation.
    :param xml_tree:
    :return:
    """
    for node in iter_text(xml_tree):
        if node.text == '#date':
            node.text = data_list['DATEREC\'D']
        elif node.text == '#file_no':
            node.text = 'FILENO.' + data_list['FILENO.']
    return xml_tree


def _write_and_close_docx(xml_tree, xml_footer2_tree, xml_footer1_tree):
    """
        Create a temp directory, expand the original docx zip.
        Write the modified xml to word/document.xml
        Zip it up as the new docx
    """

    tmp_dir = tempfile.mkdtemp()
    zip = zipfile.ZipFile('test.docx')
    zip.extractall(tmp_dir)

    # Write updated info to document.xml
    with open(os.path.join(tmp_dir, 'word/document.xml'), 'w') as f:
        xmlstr = etree.tostring(update_xml_content(xml_tree))
        f.write(xmlstr)
    # Write updated footer info to footer2.xml
    with open(os.path.join(tmp_dir, 'word/footer2.xml'), 'w') as f:
        xml_footer_str = etree.tostring(update_footer_xml(xml_footer2_tree))
        f.write(xml_footer_str)

    # Write updated footer info to footer1.xml
    with open(os.path.join(tmp_dir, 'word/footer1.xml'), 'w') as f:
        xml_footer_str = etree.tostring(update_footer_xml(xml_footer1_tree))
        f.write(xml_footer_str)

    # Get a list of all the files in the original docx zipfile
    filenames = zip.namelist()

    # Now, create the new zip file and add all the filex into the archive
    with zipfile.ZipFile('demo.docx', "w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmp_dir, filename), filename)

    # Clean up the temp dir
    shutil.rmtree(tmp_dir)

if __name__ == '__main__':
    # Get client info and insured info
    make_client_insured_info()

    # Get footer2 content from test.docx
    xml_footer2_content = get_footer2_xml('test.docx')

    # Get footer1 content from test.docx
    xml_footer1_content = get_footer1_xml('test.docx')

    # Get doc content from test.docx
    xml_content = get_word_xml('test.docx')

    # Get xml tree for footer1.xml, footer2.xml, document.xml.
    xml_tree = get_xml_tree(xml_content)
    xml_footer2_tree = get_xml_tree(xml_footer2_content)
    xml_footer1_tree = get_xml_tree(xml_footer1_content)

    # Modify and save doc
    _write_and_close_docx(xml_tree, xml_footer2_tree, xml_footer1_tree)
