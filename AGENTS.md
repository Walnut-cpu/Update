身份：知识提取师
身份背景：你是一个医学专家（眼科专业） 
任务：1. 读取Markdown文件夹中自今日之后新建的markdown文件，识别其中的疾病实体和关系，实体类型包括Disease, Synonym, Staging and typing, Anatomical location, Examination, Symptom, Physical sign, OCT sign, Gene, Differential diagnosis, Complication, Etiology, Related disease, Treatment (General, Drug, Drug usage, Sugery, Indications and Contraindications), Age of onset, High risk population and Medical history；关系类型包括Contain, Same as, Classified as, Located in, Requires examination, Has OCT sign, Has symptom, Has physical sign, Related gene, Needs distinguished from, May case, Caused by, Related to, Onset during, Affects population and Related history
2. 在文件夹中的FundiGraph.xlsx中找到对应的疾病（如果没找到就用这个疾病的同义词找），而后对比所提取的知识与原表格中对应栏目的数据的内容是否有重叠，如果有完全相同或者相似度超过80%的，就不管，如果没有，就新建一行空白行（如果原有行列是空白的，也可以不新建直接填充），填充到相应的位置，并且把前面的疾病全部填满（按照原来表格的名称）
输出格式：1. 输出为markdown格式： (: Entity label {name: Disease entity})-[: relationship]-> (: Entity label {name: Related entity})，存储在新建的markdown文件中，然后提醒进行人工核验，每一条都提供确认按钮，如果有问题，则在选择错误按钮之后进行人工修正
2. 待人工确认和修改后，用核验后的数据覆盖原有的FundiGraph.xlsx文件，并且高亮修改部分

现在页面挺好的了，但是在excel更新这一块有点东西我想修改：
1. 如果搜索的疾病在excel中对应二级疾病：那么就在对应二级疾病最末尾一行后面再新建新的行进行信息的补充，同时填充前面的二级疾病和一级疾病，以确保前面没有信息遗漏
2. 如果提取的信息属于三级疾病（下属无四级疾病和分期）：那么就在三级疾病的最末尾一行后面再新建新的行进行信息的补充，同时填充前面的三级疾病、二级疾病和一级疾病，以确保前面没有信息遗漏
3. 如果提取的信息属于三级疾病（下属有分期而无四级疾病）：那么就在三级疾病的对应分期的行数最末尾一行后面再新建新的行进行信息的补充，同时填充前面的分类信息、三级疾病、二级疾病和一级疾病，以确保前面没有信息遗漏
4. 如果提取的信息属于三级疾病（下属有四级疾病而无分期）：那么就在四级疾病的最末尾一行后面再新建新的行进行信息的补充，同时填充前面的四级疾病、三级疾病、二级疾病和一级疾病，以确保前面没有信息遗漏
5. 如果提取的信息属于三级疾病（下属有四级疾病和分期）：那么就在四级疾病的最末尾一行后面再新建新的行进行信息的补充，同时填充前面的分类四级疾病、三级疾病、二级疾病和一级疾病，以确保前面没有信息遗漏

6. 在进行人工验证前，先由gemini的LLM（这个你选价格最划算的模型）进行核验，然后给出初步的correct和error的评判，然后再给人工核验
请你帮我完成以上任务要求，使得我在使用更新代码的时候他会自动同步，你可以输出几个例子给我看一下


身份：你是一个医学专家，也是个文献检索专家
任务：读取文件夹中Disease_list.xlsx这个文件中的疾病名称，并在开放网络资源中寻找自2026年6月5日之后的所有数据，包括但不局限于英文论文、书籍、专家共识、指南等，均获取Open Access的摘要，将获取的DOI或可识别编码存储在文件夹中的DOI_list.xlsx，将摘要下载保存为markdown格式文件存在在文件夹的Markdown文件夹中

在文献检索这一块，我希望模型能够自动识别disease名称，而后使用大模型自动生成文献检索式之后再去检索相关文献，已经生成的检索式储存在本地，之后再次使用的时候直接用本地的检索式，然后如果有新的疾病实体的话，再生成新的检索式，然后也储存到本地以便于后面使用，这个是循环的