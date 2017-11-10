import pandas as pd
import numpy as np
import copy
'''
df1 = pd.read_csv("TSC1-alleles_IonXpress_068.xls", sep="\t", header=0)
df2 = pd.read_csv("TSC2-alleles_IonXpress_067.xls", sep="\t", header=0)

#pieces={"T1": df1,"T2": df2}
var=pd.concat([df1, df2])

var.index = range(len(var))
print(var.head())

drop_index = []

flag_len = len(var)

start = 0
#import pdb;pdb.set_trace()
while start < flag_len:
	name = var.loc[start, 'Allele Name']
	tmp_var = var[var['Allele Name'] == name]
	if len(tmp_var)  == 1:
		continue
	tmp_tmp_var = tmp_var[tmp_var['Allele Call'] != 'No Call' ]

	if len(tmp_tmp_var) > 0:
		#import pdb; pdb.set_trace()
		indexes = tmp_tmp_var.index	
		coverage = tmp_tmp_var['Coverage'].values.tolist()	
	else:
		indexes = tmp_var.index
		coverage = tmp_var['Coverage'].values.tolist()	

	ind = coverage.index(max(coverage))
	if indexes[ind] not in drop_index:
		drop_index.append(indexes[ind])
	start += 1

print(drop_index)
var_fmt = var.drop(set(range(flag_len)) - set(drop_index))

var_fmt.to_excel('out.xlsx', index=False)
'''
#---------------------------------------------------------
df3=pd.read_excel('out.xlsx', sheetname=0)
db =pd.read_excel('TSC1_Database_20170712.xls', sheetname=0)

#tmp=db[db["UniqueID"].isin(df3["Allele Name"])]

for index, row in db.iterrows():
	tmpDf=df3[df3["Allele Name"]==row["UniqueID"]]
	if not len(tmpDf): continue
	db.loc[index, 'Allele Call'] = tmpDf['Allele Call'].values.tolist()[0]
	db.loc[index, 'Coverage'] = tmpDf['Coverage'].values.tolist()[0]

#	print db["1000frequency"]
db.to_excel("out111.xlsx", index=False)
#import pdb;pdb.set_trace()
def splitAAChange(x):
	s1=x.split(",")[0]	
	s2=s1.split(':')
	transcNum = '-'
	baseAltC = '-'
	amiAlt = '-'
	if len(s2)>1:
		transcNum = s2[1]
	if len(s2)>3:	
		baseAltC = s2[3]
	if len(s2)>4:
		amiAlt = s2[4]
	return transcNum,baseAltC,amiAlt
db["transcNum"],db["baseAltC"],db["amiAlt"]= zip(*db['AAChange'].map(splitAAChange))

db.to_excel("out222.xlsx", index=False)

tmp1=db[["UniqueID","Ref","Alt","Gene", "cytoBand","dbSNP","transcNum","baseAltC","amiAlt","1000g2015aug_all","Clinvar","Allele Call","Coverage","CLNDBN"]]
tmp1.to_excel("out333.xlsx", index=False)
