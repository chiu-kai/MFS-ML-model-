# MFS-ML-model-

MFS_ML_model_heart.ipynb  :  
  為建立機器學習模型。
  
155 fivegroup 為建立機器學習模型的data。

MFS mutation.py  :  
  是將醫師所提供的基因突變位置資料，判斷出每一個突變位置位於第幾個Exon/Intron上、為何種突變類型 (Mutation Type)、原本的胺基酸變成哪種胺基   酸、突變的胺基酸位於fibrillin-1蛋白哪一個Domain上，以及突變的胺基酸是否位於Key residues上，並將這些分類資料區分為五種突變種類(分別為     Intronic mutations，Non-missense mutations，Not on cbEGF-like domains mutations，Not on the key residues mutations，on the key   residues mutations)作為後續預測模型裡的類別變量。
  
MFS new list 改.xlsx  :  醫師提供的Single nucleotide variants資料。

exon.csv  :  FBN1 Exon的位置 

domain.csv  :  FBN1 Protein Domain的位置 

FBN1 mRNA CDS(8616)(2).txt  :  FBN1 mRNA 的coding sequence 序列。
