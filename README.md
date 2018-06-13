# Filmtrust_ISA
----------------------------------------数据预处理--------------------
Step1：	
	目的：划分训练集测试集（Success）
	注意：excel无抬头 评分*2（0~10）
	工具：Data_pre.py
	引用：FilmTrust_ratings.txt
	结果：train\\Filmtrust_array_train_65%d.xlsx
		FilmTrust_ratings_test_241.txt
	用时：62s
Step2：
	目的：先划分测试集数据
	工具：ISA_test_method.py	import Test_split(t)	
	引用：train\\FilmTrust_ratings_test_65%d.txt		
	结果：eva\Filmtrust_isa_%d_%d_nmf_evaluate.xlsx
	用时：s
------------------------------------ISA聚类------------------------------
Step3：
	目的：ISA
	工具：Test_mtba_biclust.m (Matlab)
	引用：Filmtrust_array_train_242.xlsx
	结果：result_isa.m

Step4：	
	目的：主函数
	工具：ISA_for.py	
Step4：
	目的：存储聚类矩阵总体信息及每个聚类的评分信息
	工具：ISA_method.py	import ISA_save
	注意：总矩阵无抬头，方便直接读入Matlab中，具体群组信息需要两文件配合使用
	引用：result_242_800.mat	(t,k)训练集编号，分组数
		Filmtrust_array_train_242.xlsx
	结果：Filmtrust_isa_242_780_sta.xlsx
		Filmtrust_isa_242_780_all.xlsx
	用时：254s
评分数据共  			35497条
分组覆盖结果：242	2000	780
分组覆盖结果：242	2000	800	181882	30545	4949
分组覆盖结果：242	2000	803	3346	23294	12200

-----------------------------------------------群组融合 均值-------------------------
Step1：
	目的：计算群组评分均值（评分人数少于0.7，设为0）（或者可以设方差）
	工具：ISA_ave_method.py		import ISA_group_ave
	引用：Filmtrust_isa_242_803_all.xlsx
		Filmtrust_isa_242_803_sta.xlsx
	结果：Filmtrust_isa_242_803_ave.xlsx
	用时：160s
Step2：
	目的：群组评分均值融合详细（包括用户计数、均值、方差等）
	工具：ISA_ave_det_method.py	import ISA_group_ave_det
	引用：Filmtrust_isa_242_803_all.xlsx
	结果：Filmtrust_isa_242_803_ave_det.xlsx
	用时：s

Step4：
	目的：构建群组评分网络
	工具：ISA_net_method.py	import ISA_group_ave_net(t,k)
	引用：sta\\Filmtrust_isa_%d_%d_ave.xlsx
		sta\\Filmtrust_isa_%d_%d_sta.xlsx		
	结果：sta\\Filmtrust_isa_%d_%d_ave_net.xlsx
	用时：58s
Step3：
	目的：群组倒排档，计算并存储用户与每个群组的均值Virtual User Pearson相似度
	工具：ISA_inverted_method.py	import ISA_group_inverted	
	引用：sta\\Filmtrust_isa_%d_%d_ave_net.xlsx
		nmf\\Filmtrust_isa_%d_%d_all_nmf.xlsx
		sta\\Filmtrust_isa_%d_%d_sta.xlsx
		sta\\Filmtrust_isa_%d_%d_ave.xlsx
	结果：sta\Filmtrust_isa_%d_%d_inverted.xlsx
	用时：s
Step3：
	目的：协同推荐
	工具：ISA_forecase_method.py 	import ISA_group_forecast
	引用：train\\Filmtrust_test_split_65%d.xlsx
		sta\\Filmtrust_isa_%d_%d_ave.xlsx
		sta\\Filmtrust_isa_%d_%d_ave_net.xlsx
		sta\Filmtrust_isa_%d_%d_inverted.xlsx
	结果：eva\\Filmtrust_isa_%d_%d_ave_forecast.xlsx
	用时：s
Step4：
	目的：Top-N推荐 nDCG评估
	工具：ISA_evaluation_method.py 	import ISA_group_evaluation	
	引用：train\\Filmtrust_test_split_%d.xlsx
		eva\\Filmtrust_isa_%d_%d_ave_forecast.xlsx
	结果：eva\\Filmtrust_isa_%d_%d_ave_nDCG.xlsx
	用时：15s

Step5:
	目的：最后统计所有nDCG结果，汇成一个表格
	工具：nDCG_all.py	def  nDCG_all
	引用：r'eva\\Filmtrust_isa_%d_%d_ave_nDCG.xlsx'%(t,k)
	结果：r'eva\\Filmtrust_isa_ave_nDCG_all.xlsx'
	用时：s
