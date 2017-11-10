#! python2
# encoding=utf-8
# pzw
# 20171110

import pandas as pd
import sys
import config.word_writer as writer
import docx
import yaml

reload(sys)
sys.setdefaultencoding('utf8')
templateFilePath = './config/report_TSC_v1_gs.docx'
ad = yaml.load(open('openDir.yaml', 'rb'))['director']
conf = yaml.load(open(ad + '/' + 'TSCInfo', 'rb'))
saveFilePath = ad + '/' + 'TSC_report.docx'
df1 = pd.read_csv(ad + '/' + "alleles_IonXpress_095.xls", sep="\t", header=0)
df2 = pd.read_csv(ad + '/' + "alleles_IonXpress_096.xls", sep="\t", header=0)

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

cleanTable.to_excel(ad + '/' + "result.xlsx", index=False)
del cleanTable['UniqueID']
del cleanTable['Ref']
del cleanTable['Alt']
del cleanTable['CLNDBN']

changeTable = cleanTable[(cleanTable['Allele Call'] == 'Homozygous') | (cleanTable['Allele Call'] == 'Heterozygous')]
sumChange = len(changeTable.index)

pathTable = cleanTable[cleanTable['Clinvar'].str.contains('Pathogenic') | cleanTable['Clinvar'].str.contains('Likely pathogenic')]
sumPath = len(pathTable.index)

pathogenicTable = pathTable[pathTable['Clinvar'].str.contains('Pathogenic')]
sumPathogenic = len(pathogenicTable.index)

cpathTable = changeTable[changeTable['Clinvar'].str.contains('Pathogenic') | changeTable['Clinvar'].str.contains('Likely pathogenic')]
sumcPath = len(cpathTable.index)

cpathogenicTable = changeTable[changeTable['Clinvar'].str.contains('Pathogenic')]
sumcPathogenic = len(cpathogenicTable.index)

pathTableDC = pathTable[pathTable['Allele Call'].str.contains('Wt')]

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
resultMap['#[lp1]#'] = sumPath - sumPathogenic
resultMap['#[other1]#'] = sumAll - sumPath
resultMap['#[pathogenic2]#'] = sumcPathogenic
resultMap['#[lp2]#'] = sumcPath - sumcPathogenic
resultMap['#[other2]#'] = sumChange - sumcPath
resultMap['#[FILLTABLE-change]#'] = cpathTable
resultMap['#[FILLTABLE-DC]#'] = pathTableDC


report = writer.fillAnalyseResultMap(resultMap, report)
# report = writer.deleteEmptyTable(report,'#[FILLTABLE-')
report.save(saveFilePath)
print 'task done'