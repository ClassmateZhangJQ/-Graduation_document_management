import jieba
from collections import defaultdict
from gensim import corpora
from gensim import models, similarities
from PIL import Image, ImageTk
class Title_sims():
    def __init__(self,data,score,epoch,stopwords):
        self.data = data
        self.score = score
        self.epoch = epoch
        self.stopwords = stopwords
        self.texts = self.process_data(self.data,self.stopwords)
        self.time = 0
        self.res = [] # 存放run 后的结果
        self.ret = [] # 存储要返回去的结果
        self.run()
        self.deal_with_list_with_tuple(self.res,self.ret)

    def run(self):
        for i in range(self.epoch):
            self.time = i
            self.get_result()


    @staticmethod
    def process_data(data,stopwords):
        document = []
        for line in data:
            temp = []
            seq = jieba.cut(line.strip())
            for word in seq:
                if word not in stopwords and word != '\u3000' and word != ' ':
                    temp.append(word)
            document.append(temp)

        texts = document

        frequency = defaultdict(int)
        for text in texts:
            for token in text:
                frequency[token] += 1
        return texts
    def get_result(self):

        dictionary = corpora.Dictionary(self.texts)
        # dictionary.save('sample.json')
        corpus = [dictionary.doc2bow(text) for text in self.texts]

        tf_idf = models.TfidfModel(corpus)

        num_features = len(dictionary.token2id.keys())

        index = similarities.MatrixSimilarity(tf_idf[corpus],num_features = num_features)

        test_word = self.texts
        # for i in range(len(test_word)):
        new_vec = dictionary.doc2bow(test_word[self.time])

        sims = index[tf_idf[new_vec]]
        sims = sorted(enumerate(sims),key = lambda item:-item[1])[1::]
        result = []
        for index , score in sims:
            if score > self.score:
                result.append((self.data[index].replace('\n',''),score))
            else:
                break

        if len(result):
                temp = []
                temp.append(self.data[self.time].replace('\n',''))
                # print('与 <<{}>> 相似的有->'.format(self.data[self.time].replace('\n','')),result)
                for key in result:
                    temp.append(key[0])
                self.res.append(temp)

        # 对应每一个标题相似的结果
    @staticmethod
    def deal_with_list_with_tuple(_list,_rest):
        flag = {}
        rest = _rest
        for i in range(len(_list)):
            flag[i] = 0
        for i in range(len(_list)):  # 遍历a的所有类别
            if flag[i] == 1:  # 标记改类别是否被处理过
                continue
            start = _list[i]  # 比较对象
            j = i + 1
            while (j < len(_list)):  # 比较样本类别循环
                per_list = _list[j]  # 获取单次比较样本
                for temp in per_list:  # 循环单次类别内容
                    if (start.count(temp) and flag[j] == 0):
                        for i in per_list:
                            if start.count(i):
                                continue
                            else:
                                start.append(i)
                        flag[j] = 1
                        break
                j += 1
            rest.append(start)

# if __name__ == '__main__':
#     image = Image.open('logo/artistic_body01.png')
#     print(image)
