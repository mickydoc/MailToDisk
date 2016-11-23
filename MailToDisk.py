# -*- coding: utf-8 -*-
#########################################################################################################
# MailToDisk
#
# Replacement for XAMPP (or any other development stack) mailtodisk.exe
# - Setup in config if mail should get opened after creation
# - Setup in config file if 300MB Folder limit should get set (need some I/O time)
# - Rereated Classes/Functions - should be now much faster than original
#
# Credits: Kay Vogelgesang, XAMPP - apachefriends.org (original script)
#########################################################################################################

#########################################################################################################
# Import Libs
#########################################################################################################

import os
import sys

scriptpath = os.path.realpath(os.path.dirname(sys.argv[0])).replace('\\','/')
print(scriptpath);
sys.path.insert(0, scriptpath)
sys.path.insert(0, os.path.abspath('..'))
sys.path.insert(0, os.path.join(scriptpath, 'lib'))

if sys.maxsize == 2147483647:
    sys.path.insert(0, os.path.join(scriptpath, 'libX86'))
    sys.path.insert(0, scriptpath + '/libX86/libraries.zip')
    print ("Running x86")
else:
    sys.path.insert(0, os.path.join(scriptpath, 'libX64'))
    sys.path.insert(0, scriptpath + '/libX64/libraries.zip')
    print ("Running x64")


from time import gmtime,strftime
from datetime import datetime
import dict4ini

#####################################################################################################
# Load Settings
#####################################################################################################

cybeSystemsMainSettings = {}
iniMainSettings = None
allowed_file_extensions = ['htm', 'html', 'txt', 'eml']

mainConfigFile = scriptpath + '/config.ini'

cybeSystemsMainSettings = dict4ini.DictIni(mainConfigFile)
iniMainSettings = dict4ini.DictIni(mainConfigFile)

def defaultMainSettingsIni():
    cybeSystemsMainSettings['Main'] = {}
    #This option open the eml file with your default assigned program after generation (e.g. Thunderbird, Outlook...)
    cybeSystemsMainSettings['Main']['OpenGeneratedEmail'] = True
    cybeSystemsMainSettings['Main']['300MBFolderLimit'] = True
    cybeSystemsMainSettings['Main']['OutputFolder'] = "mailoutput"
    cybeSystemsMainSettings['Main']['OutputFileExtension'] = "eml" # Can be "htm", "html", "txt, "eml". See "allowed_file_extensions"
    cybeSystemsMainSettings['Main']['UseConservativeFileNaming'] = False

defaultMainSettingsIni()

def replaceSetting():
    for section in list(iniMainSettings.keys()):
        for opt in list(iniMainSettings[section].keys()):
            value = iniMainSettings[section][opt]
            if isinstance(value, list):
                cybeSystemsMainSettings[section][opt] = value
            else:
                cybeSystemsMainSettings[section][opt] = str(value)
            if value == 'true' or value == 'True':
                cybeSystemsMainSettings[section][opt] = True
            if value == 'false' or value == 'False':
                cybeSystemsMainSettings[section][opt] = False
            if str(value).isdigit() == True:
                cybeSystemsMainSettings[section][opt] = int(value)

#Get Values from ini file -> If not found use default values
replaceSetting()

#####################################################################################################
# Helper functions
#####################################################################################################

def getFolderSize(folder):
    total_size = os.path.getsize(folder)
    for item in os.listdir(folder):
        itempath = os.path.join(folder, item)
        if os.path.isfile(itempath):
            total_size += os.path.getsize(itempath)
        elif os.path.isdir(itempath):
            total_size += getFolderSize(itempath)
    return total_size

class MailToDisk():

    mailOutputFolder = os.path.join(os.path.dirname(scriptpath), cybeSystemsMainSettings['Main']['OutputFolder'])
    
    chosen_file_extension = str(cybeSystemsMainSettings['Main']['OutputFileExtension'])
    
    if cybeSystemsMainSettings['Main']['UseConservativeFileNaming'] == False:
        filename = mailOutputFolder + "\\mail_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    else:
        filename = mailOutputFolder + "\\mail-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    
    if chosen_file_extension in allowed_file_extensions:
        filename = filename + "." + chosen_file_extension
    else:
        filename = filename + ".txt"
    

    # Security restriction: mailoutput folder may not have more then 300 MB overall size for write in
    if cybeSystemsMainSettings['Main']['300MBFolderLimit'] == True:
        if os.path.exists(mailOutputFolder):
            mailOutputFolderSize = getFolderSize(mailOutputFolder)
            if mailOutputFolderSize > 314572800: # 300 MB
                f = open(mailOutputFolder + "\\MAILTODISK_WRITE_RESTRICTION_FOLDER_MORE_THEN_300_MB.txt", 'w')
                f.write("MailToDisk will NOT write in a folder with a overall size of 300 MB (security limit)! Please clean up this folder!!!")
                f.close()
                sys.exit(1)

    def getEmailText(self):
        eMailText = ""
        while 1:
            nextLine = sys.stdin.readline()
            eMailText = eMailText + nextLine
            if not nextLine:
                break
        return eMailText

    def writeEmlFile(self):
        if not os.path.exists(self.mailOutputFolder):
                os.makedirs(self.mailOutputFolder)
        eMailText = self.getEmailText()
        f = open(self.filename, 'w')
        f.write(eMailText)
        f.close()
        if cybeSystemsMainSettings['Main']['OpenGeneratedEmail'] == True:
            os.startfile(self.filename)


if __name__ == '__main__':
    writeEmailText = MailToDisk()
    writeEmailText.writeEmlFile()
    sys.exit()