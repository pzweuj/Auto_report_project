#!/usr/bin/python2.7
# encoding=utf-8
# pzw
# 20180102

import pandas as pd
import sys
import config.word_writer as writer
import docx
import yaml

reload(sys)
sys.setdefaultencoding('utf8')
templateFilePath = './config/report_TSC_v2_gs.docx'
dirinfo = yaml.load(open('openDir.yaml', 'rb'))
ad = dirinfo['director']
conf = yaml.load(open(ad + '/' + 'TSCInfo.yaml', 'rb'))
saveFilePath = ad + '/' + 'TSC_report.docx'
exwriter = pd.ExcelWriter(ad + '/' + 'results.xlsx')

if dirinfo['amount'] == 2:
	file1 = dirinfo['file1']
	file2 = dirinfo['file2']
	df1 = pd.read_csv(ad + '/' + file1, sep="\t", header=0)
	df2 = pd.read_csv(ad + '/' + file2, sep="\t", header=0)

	variant = pd.concat([df1, df2])
	variant.index = range(len(variant))

	drop_index = []
	flag_len = len(variant)
	start = 0
	while start < flag_len:
	    name = variant.loc[start, 'Allele Name']
	    tmp_var = variant[variant['Allele Name'] == name]
	    if len(tmp_var) == 1:
	        continue
	    tmp_tmp_var = tmp_var[tmp_var['Allele Call'] != 'No Call']

	    if len(tmp_tmp_var) > 0:
	        indexes = tmp_tmp_var.index
	        coverage = tmp_tmp_var['Coverage'].values.tolist()
	    else:
	        indexes = tmp_var.index
	        coverage = tmp_var['Coverage'].values.tolist()

	    ind = coverage.index(max(coverage))
	    if indexes[ind] not in drop_index:
	        drop_index.append(indexes[ind])
	    start += 1

	var_fmt = variant.drop(set(range(flag_len)) - set(drop_index))
elif dirinfo['amount'] == 1:
	file0 = dirinfo['file0']
	var_fmt = pd.read_table(ad + '/' + file0, header = 0, sep = '\t')
else:
	print 'Waring: Please check the origin file amount'

# var_fmt.to_excel(ad + '/' + 'out.xlsx', index=False)

# df3 = pd.read_excel(ad + '/' + 'out.xlsx', sheetname=0)
db = pd.read_excel('config/TSC1_Database_20170712.xls', sheetname=0)

for index, row in db.iterrows():
    tmpDf = var_fmt[var_fmt["Allele Name"] == row["UniqueID"]]
    if not len(tmpDf): continue
    db.loc[index, 'Allele Call'] = tmpDf['Allele Call'].values.tolist()[0]
    db.loc[index, 'Coverage'] = tmpDf['Coverage'].values.tolist()[0]


# db.to_excel("out111.xlsx", index=False)

def splitAAChange(x):
    s1 = x.split(",")[0]
    s2 = s1.split(':')
    transcNum = '-'
    baseAltC = '-'
    amiAlt = '-'
    if len(s2) > 1:
        transcNum = s2[1]
    if len(s2) > 3:
        baseAltC = s2[3]
    if len(s2) > 4:
        amiAlt = s2[4]
    return transcNum, baseAltC, amiAlt


db["transcNum"], db["baseAltC"], db["amiAlt"] = zip(*db['AAChange'].map(splitAAChange))

# db.to_excel("out222.xlsx", index=False)

tmp1 = db[["UniqueID", "Ref", "Alt", "Gene", "cytoBand", "dbSNP", "transcNum", "baseAltC", "amiAlt", "1000g2015aug_all",
           "Clinvar", "Allele Call", "Coverage", "CLNDBN"]]
# tmp1.to_excel(ad + '/' + "result.xlsx", index=False)

cleanTable = tmp1[(tmp1['Allele Call'] == 'Absent') | (tmp1['Allele Call'] == 'Heterozygous') | (tmp1['Allele Call'] == 'Homozygous')]
sumAll = len(cleanTable.index)

cleanTable.loc[cleanTable['Clinvar'].str.contains('Pathogenic'), 'Clinvar'] = 'Pathogenic'
cleanTable.loc[cleanTable['Clinvar'].str.contains('pathogenic'), 'Clinvar'] = 'Likely pathogenic'
cleanTable.loc[cleanTable['Allele Call'].str.contains('Absent'), 'Allele Call'] = 'Wt'
cleanTable.loc[cleanTable['Allele Call'].str.contains('Heterozygous'), 'Allele Call'] = 'Het'
cleanTable.loc[cleanTable['Allele Call'].str.contains('Homozygous'), 'Allele Call'] = 'Hom'

# cleanTable.to_excel(ad + '/' + "result.xlsx", index=False)
del cleanTable['UniqueID']
del cleanTable['Ref']
del cleanTable['Alt']
del cleanTable['CLNDBN']

cleanTable.to_excel(exwriter, 'cleanTable', index=False)

sumChange = sum(cleanTable['Allele Call'] != 'Wt')
sumPathogenic = sum(cleanTable['Clinvar'] == 'Pathogenic')
sumLikelyPathogenic = sum(cleanTable['Clinvar'] == 'Likely pathogenic')
sumOther = sumAll - sumPathogenic - sumLikelyPathogenic
sumChangePathogenic = sum((cleanTable['Allele Call'] != 'Wt') & (cleanTable['Clinvar'] == 'Pathogenic'))
sumChangeLikelyPathogenic = sum((cleanTable['Allele Call'] != 'Wt') & (cleanTable['Clinvar'] == 'Likely pathogenic'))
sumChangeOther = sumChange - sumChangePathogenic - sumChangeLikelyPathogenic

changeTable = cleanTable[cleanTable['Allele Call'] != 'Wt']
otherTable = cleanTable[((cleanTable['Clinvar'] == 'Pathogenic') | (cleanTable['Clinvar'] == 'Likely pathogenic')) & (cleanTable['Allele Call'] == 'Wt')]
changeTable.to_excel(exwriter, 'changeTable', index=False)
otherTable.to_excel(exwriter, 'otherTable', index=False)
exwriter.save()
changeTable.to_csv(ad + '/' + 'changeTable.txt', index=False, sep='\t', header=None)
otherTable.to_csv(ad + '/' + 'otherTable.txt', index=False, sep='\t', header=None)

report = docx.Document(unicode(templateFilePath, 'utf-8'))

resultMap = {}
resultMap['#[name]#'] = conf['personalinfo']['name']
resultMap['#[gender]#'] = conf['personalinfo']['gender']
resultMap['#[date_of_birth]#'] = conf['personalinfo']['date_of_birth']
resultMap['#[phone]#'] = conf['personalinfo']['phone']
resultMap['#[ID]#'] = conf['personalinfo']['ID']
resultMap['#[GSID]#'] = conf['personalinfo']['GSID']
resultMap['#[project]#'] = conf['personalinfo']['project']
resultMap['#[diagnostic]#'] = conf['personalinfo']['diagnostic']
resultMap['#[treatment]#'] = conf['personalinfo']['treatment']
resultMap['#[family]#'] = conf['personalinfo']['family']
resultMap['#[type]#'] = conf['otherinfo']['type']
resultMap['#[amount]#'] = conf['otherinfo']['amount']
resultMap['#[doctor]#'] = conf['otherinfo']['doctor']
resultMap['#[sampling_date]#'] = conf['otherinfo']['sampling_date']
resultMap['#[collection_date]#'] = conf['otherinfo']['collection_date']
resultMap['#[report_date]#'] = conf['otherinfo']['report_date']
resultMap['#[Inspection_agencies]#'] = conf['otherinfo']['Inspection_agencies']
resultMap['#[pathogenic1]#'] = sumPathogenic
resultMap['#[lp1]#'] = sumLikelyPathogenic
resultMap['#[other1]#'] = sumOther
resultMap['#[pathogenic2]#'] = sumChangePathogenic
resultMap['#[lp2]#'] = sumChangeLikelyPathogenic
resultMap['#[other2]#'] = sumChangeOther
resultMap['#[FILLTABLE-change]#'] = file(ad + '/' + 'changeTable.txt').read()
# resultMap['#[FILLTABLE-DC]#'] = file(ad + '/' + 'otherTable.txt').read()

report = writer.fillAnalyseResultMap(resultMap, report)
# report = writer.deleteEmptyTable(report,'#[FILLTABLE-')
report.save(saveFilePath)

print 'task done'