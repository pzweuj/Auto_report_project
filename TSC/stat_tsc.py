import pandas as pd
import numpy as np

db_infos = pd.read_excel("TSC1_Database_20170712.xls", sheetname=0)
#print db_infos.head()

var_infos = pd.read_csv("GS02189-alleles_IonXpress_095.xls", sep="\t", header=0)
#print var_infos.head()

db_fmt_infos = db_infos[['UniqueID', 'Chr',	'Start',	'End', 'Ref',	'Alt', 'dbSNP',  'Amplicon_ID', 'Insert_Start',	'Insert_Stop']]
#print db_fmt_infos.head()

#import pdb; pdb.set_trace()
db_fmt_infos.loc[db_fmt_infos['Insert_Start']=='-', 'Insert_Start']=10000000
db_fmt_infos.loc[db_fmt_infos['Insert_Stop']=='-', 'Insert_Stop']=10000000

db_fmt_infos['Down_distance'] = db_fmt_infos['Start'] - db_fmt_infos['Insert_Start']
db_fmt_infos['Up_distance'] = db_fmt_infos['End'] - db_fmt_infos['Insert_Start']
db_fmt_infos['Distance'] = db_fmt_infos['Insert_Stop'] - db_fmt_infos['Insert_Start'] + 1
##############################################

var_fmt_infos = var_infos[var_infos['Coverage']==0]
stat_res = db_fmt_infos[db_fmt_infos['UniqueID'].isin(var_fmt_infos['Allele Name'])]

print(stat_res.groupby(["Amplicon_ID"])["Amplicon_ID"].count())

ampli_ids = stat_res["Amplicon_ID"].unique()
print(ampli_ids)
print("Amplicon -----> {}".format(len(ampli_ids)))

var_fmt_infos = var_fmt_infos[var_fmt_infos['Allele Name'].isin(db_fmt_infos['UniqueID'])]
for amp in ampli_ids:
	coverage = np.sum(var_fmt_infos[var_fmt_infos["Region Name"]==amp]['Coverage'])
	print "Amplicon:{}   ------>   Coverage {}".format(amp, coverage)

stat_res.to_excel("Stat_Res.xlsx", index=False)