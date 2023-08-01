import os
import shutil
from subprocess import check_output

def run_cmd(cmd_str):
  print(cmd_str)
  print(check_output(cmd_str + " ", shell=True).decode())
  return

table_file = '"PT01 (VPX 2023) 0x3e.vpx"'
table_vbs = '"PT01 (VPX 2023) 0x3e.vbs"'
there_table_file = 'C:\\"Visual Pinball"\Tables\\' + table_file
here_table_file = os.getcwd() + '\\' + table_file

print('')
print(there_table_file)
#print(os.stat(there_table_file))
print(here_table_file)
#print(os.stat(here_table_file))
#shutil.copyfile(there_table_file, here_table_file)  

run_cmd('copy /y ' + there_table_file + ' ' + here_table_file)
run_cmd('del ' + table_vbs)

run_cmd('cd C:\Visual Pinball &&  VPinballX.exe -ExtractVBS ' + here_table_file)

run_cmd('certutil -hashfile ' + here_table_file + ' MD5 > checksums.txt')

run_cmd('certutil -hashfile ' + here_table_file + ' SHA256 >> checksums.txt')
