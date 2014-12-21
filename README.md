HomeBankImportScript
====================

Script converting data exported from bank websites to HomeBank.
Supported as of 2014/12/21:
* Boursorama banque
* ING DiBa

Work in progress:
* linxo.fr

usage: conv2homebank.py [-h] [-i INPUT]
                        [-t {INGDiba_csv,Boursorama_qif,Linxo_csv}]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        input file
  -t {INGDiba_csv,Boursorama_qif,Linxo_csv}, --type {INGDiba_csv, Boursorama_qif, Linxo_csv}
                        Type of the file

If no argument is used, the script will try to process the content of the directory "In" and will output the results in "Out" (the directories have to exist).
