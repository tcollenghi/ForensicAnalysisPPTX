#DO FORENSIC ANALYSIS IN A POWERPOINT FILE 
#Author: @_hugsy_
#Date: 2018-09-12
#Version: 1.0
#Usage: python FORENSIC.py <input.pptx>
#open file and read it 
import sys
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from zipfile import BadZipfile
from zipfile import LargeZipFile

#open powerpoint file located at C:\Users\tulio and read it
def read_pptx(pptx_file):
    try:
        with ZipFile(pptx_file, 'r') as zip:
            for name in zip.namelist():
                if name.endswith('C:/Users/tulio/presentation.xml'):
                    return zip.read(name)
# check any possible changes in the file on location C:/Users/tulio/presentation.xml 
# and print the changes in a file named changes.txt
# print the result in a file named C:/Users/tulio/result.txt if changed or not 
    except BadZipfile:
        print('Bad zip file')
        sys.exit(1)
    except LargeZipFile:
        print('Large zip file')
        sys.exit(1)
    except:
        print('Unexpected error:', sys.exc_info()[0])
        sys.exit(1)

#print the file with a new name even if is ok 
#check if the file is ok and print the result in a file named C:/Users/tulio/result.txt 
# if changed or not
# if the file is not ok print the result in a file named changes.txt
# if changed or not
# if the file is not ok print the result in a file named changes.txt
def print_file(pptx_file):
    try:
        with ZipFile(pptx_file, 'r') as zip:
            for name in zip.namelist():
                if name.endswith('C:/Users/tulio/result.txt'):
                    return zip.read(name)
    except BadZipfile:
        print('Bad zip file')
        sys.exit(1)
    except LargeZipFile:
        print('Large zip file')
        sys.exit(1)
    except:
        print('Unexpected error:', sys.exc_info()[0])
        sys.exit(1)
    