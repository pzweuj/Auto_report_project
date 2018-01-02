# 20180102
# pzw

1，在TSC目录下创建文件夹，如GS00000
2，在Torrent Server中下载相关的数据，如alleles_IonXpress_095.xls，
放入GS00000文件夹中。
3，复制TSC文件夹中的TSCINfo.yaml到GS00000，并填写好相关的信息。
4，修改TSC文件夹中的openDir.yaml文件。
5，cmd进入TSC文件夹，输入命令python report_TSC_v2.py运行

```{python}
python report_TSC_v2.py
```

6，进入GS00000，
修改TSC_report.docx，另存为pdf，检查，发送。