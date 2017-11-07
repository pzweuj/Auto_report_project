# encoding=utf-8
# pzw
# 2017/11/07
import sys
import pandas as pd
import yaml
import docx
import word_writer as writer


reload(sys)
sys.setdefaultencoding('utf8')
f = open('cancerDrugNoTarget.yaml')
conf = yaml.load(f)
templateFilePath = 'report_CD_v1_gs.docx'
saveFilePath = 'CancerDrug_report.docx'
df = pd.read_excel(conf['director'] + '\\' + 'temp.xlsx', sheetname = 0, header = 0)
med_stan = pd.read_excel(r'config\DataBase\med.xlsx', sheetname = 0, header = 0)
# exwriter = pd.ExcelWriter(conf['director'] + '\\' + 'out_all.xlsx')
out_des = open(conf['director'] + '\\' + 'out_des.txt', 'w')

df.columns = ['med', 'gene', 'rsid', 'genetype', 'sen', 'pmid', 'level', 'cancer']
var_detected = pd.DataFrame(columns=['gene', 'rsid', 'alle', 'genetype'])
var_detected['gene'] = df['gene']
var_detected['rsid'] = df['rsid']
var_detected['genetype'] = df['genetype']

# 等位基因
for i in range(var_detected['genetype'].count()):
    ele = var_detected['genetype'][i].split(';')
    if ele[0] == ele[1]:
        var_detected['alle'][i] = '纯合子'
    else:
        var_detected['alle'][i] = '杂合子'

var_selected = var_detected.sort_values(by='gene').drop_duplicates()

# 输出等位基因
# var_selected.to_excel(writer, 'rsfound', index=False)

l = []
for j in var_selected['gene'].drop_duplicates().index:
    l.append(var_selected['gene'][j])

# 阳性基因
gene_detected = u'、'.join(l)

# 填写敏感性与风险的高低
df['sense'] = '正常'
df['risk'] = '正常'
for k in range(df['sen'].count()):
    if df['sen'][k].__contains__(u'生存期'):
        if df['sen'][k].__contains__(u'增加'):
            df['sense'][k] = '高'
        if df['sen'][k].__contains__(u'减少'):
            df['sense'][k] = '低'

    if df['sen'][k].__contains__(u'药效'):
        if df['sen'][k].__contains__(u'好'):
            df['sense'][k] = '高'
        if df['sen'][k].__contains__(u'差'):
            df['sense'][k] = '低'

    if df['sen'][k].__contains__(u'耐药性'):
        if df['sen'][k].__contains__(u'减少') or df['sen'][k].__contains__(u'低'):
            df['sense'][k] = '高'
        if df['sen'][k].__contains__(u'增加'):
            df['sense'][k] = '低'

    if df['sen'][k].__contains__(u'吸收'):
        if df['sen'][k].__contains__(u'减少') or df['sen'][k].__contains__(u'降低'):
            df['sense'][k] = '高'
        if df['sen'][k].__contains__(u'增加'):
            df['sense'][k] = '低'

    if df['sen'][k].__contains__(u'曲线'):
        if df['sen'][k].__contains__(u'增加'):
            df['sense'][k] = '高'
        if df['sen'][k].__contains__(u'减少'):
            df['sense'][k] = '低'
			
    if df['sen'][k].__contains__(u'血药浓度'):
        if df['sen'][k].__contains__(u'升高'):
            df['sense'][k] = '高'
        if df['sen'][k].__contains__(u'血药浓度下降') or df['sen'][k].__contains__(u'血药浓度降低'):
            df['sense'][k] = '低'		

    if df['sen'][k].__contains__(u'风险较高'):
        df['risk'][k] = '高'

    if df['sen'][k].__contains__(u'风险较低'):
        df['risk'][k] = '低'

    if df['sen'][k].__contains__(u'风险增加'):
        df['risk'][k] = '高'

    if df['sen'][k].__contains__(u'风险减少'):
        df['risk'][k] = '低'

# med_detected = pd.DataFrame(columns=['med', 'sen', 'sense', 'risk'])
med_detected = pd.DataFrame(columns=['med', 'sense', 'risk'])
med_detected['med'] = df['med']
# med_detected['sen'] = df['sen']
med_detected['sense'] = df['sense']
med_detected['risk'] = df['risk']

# 输出检测出的药物
# med_detected.to_excel(writer, 'medd', index=False)

# 同种药物的概括
sense = {}
risk = {}
for ele in range(med_stan['med'].count()):
    sense[med_stan['med'][ele]] = '正常'
    risk[med_stan['med'][ele]] = '正常'

for des in range(med_detected['med'].count()):
    if med_detected['sense'][des] != '正常':
        sense[med_detected['med'][des]] = med_detected['sense'][des]

    if med_detected['risk'][des] != '正常':
        risk[med_detected['med'][des]] = med_detected['risk'][des]

for ele2 in range(med_stan['med'].count()):
    for key in sense.keys():
        if med_stan['med'][ele2] == key:
            med_stan['sense'][ele2] = sense[key]

for ele3 in range(med_stan['med'].count()):
    for key in risk.keys():
        if med_stan['med'][ele3] == key:
            med_stan['risk'][ele3] = risk[key]

# 删除正常的药物联用
for x in range(len(med_stan.index)):
    if x <= 21:
        if med_stan['sense'][x] == '正常' and med_stan['risk'][x] == '正常':
            med_stan.drop(x, inplace=True)
        else:
            continue
    else:
        break

# 概括分类
senseHigh = med_stan[med_stan.sense == '高']
senseLow = med_stan[med_stan.sense == '低']
riskHigh = med_stan[med_stan.risk == '高']
riskLow = med_stan[med_stan.risk == '低']

senseHigh_l = []
for l in senseHigh.index:
    senseHigh_l.append(senseHigh['med'][l])
senseHigh_des =  u'、'.join(senseHigh_l)

senseLow_l = []
for l in senseLow.index:
    senseLow_l.append(senseLow['med'][l])
senseLow_des =  u'、'.join(senseLow_l)

riskHigh_l = []
for l in riskHigh.index:
    riskHigh_l.append(riskHigh['med'][l])
riskHigh_des =  u'、'.join(riskHigh_l)

riskLow_l = []
for l in riskLow.index:
    riskLow_l.append(riskLow['med'][l])
riskLow_des =  u'、'.join(riskLow_l)

out_des.write(senseHigh_des.encode('utf-8') + '\n')
out_des.write(senseLow_des.encode('utf-8') + '\n')
out_des.write(riskHigh_des.encode('utf-8') + '\n')
out_des.write(riskLow_des.encode('utf-8') + '\n')
out_des.write(gene_detected.encode('utf-8') + '\n')
del df['risk']
del df['cancer']
del df['sense']

# 输出!
# med_stan.to_excel(exwriter, 'medfound', index=False)
# df.to_excel(exwriter, 'med_des', index=False)
# var_selected.to_excel(exwriter, 'rsfound', index=False)
# exwriter.save()
med_stan.to_csv('med_stan.txt', index=False, sep='\t', header=None)
df.to_csv('df.txt', index=False, sep='\t', header=None)
var_selected.to_csv('var_selected.txt', index=False, sep='\t', header=None)
out_des.close()
f.close()

# 自动写入word
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
resultMap['#[senseHigh_des]#'] = senseHigh_des
resultMap['#[senseLow_des]#'] = senseLow_des
resultMap['#[riskHigh_des]#'] = riskHigh_des
resultMap['#[riskLow_des]#'] = riskLow_des
resultMap['#[FILLTABLE-med_stan(0)]#'] = file('med_stan.txt').read()

if resultMap['#[gender]#'] == '男':
    resultMap['#[genderD]#'] = '先生'
elif resultMap['#[gender]#'] == '女':
    resultMap['#[genderD]#'] = '女士'
else:
    resultMap['#[genderD]#'] = ''

report = writer.fillAnalyseResultMap(resultMap, report)
report = writer.deleteEmptyTable(report,'#[FILLTABLE-')
report.save(saveFilePath)