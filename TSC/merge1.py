import pandas as pd
import numpy as np
import copy

df1 = pd.read_csv("alleles_IonXpress_095.xls", sep="\t", header=0)
df2 = pd.read_csv("alleles_IonXpress_096.xls", sep="\t", header=0)

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