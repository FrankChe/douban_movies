#coding=utf-8
__author__ = 'chexiaoyu'

import urllib
import urllib2
import requests
import json
import xlwt

class Douban_movie:

    def __init__(self):
        self.rank = []  #电影排名
        self.title = [] #存放电影名称
        self.id = []    #电影id
        self.types = [] #电影类型
        self.regions = []   #电影地区
        self.actors = []    #演员
        self.release_date = []
        self.actor_count = []
        self.vote_count = []
        self.rate = []
        self.type = None
        self.wb = xlwt.Workbook()

        self.first = 1  #判断是否第一次进入
        self.start = 0  #判断从xls文件中第几行开始写

    def spider(self, num, type):
        url = 'http://movie.douban.com/j/chart/top_list'
        #电影的类型为1到31
        self.type = type
        parameter = {
            'type' : self.type,
            'interval_id' : '100:90',
            'start' : num,
            'limit' : 100
        }

        r = requests.get(url,params=parameter)
        data = json.loads(r.text)
        if data == []:
            return False
        #print data

        #print data
        for item in data:
            self.data_handle(item)
        self.start = self.start + num
        self.write_xls(data)

        print self.rank
        return True
            #print 'Top',i['title']
    def data_handle(self, data):
        self.rank.append(data['rank'])
        self.title.append(data['title'])
        self.id.append(data['id'])
        types = data['types']
        self.types.append(types)
        regions = data['regions']
        self.regions.append(regions)
        actors = data['actors']
        self.actors.append(actors)
        self.release_date.append(data['release_date'])
        self.actor_count.append(data['actor_count'])
        self.vote_count.append(data['vote_count'])
        self.rate.append(data['rating'][0])

    def write_xls(self, data):
        if self.first == 1:
            print self.type
            ws = self.wb.add_sheet("%s" % str(self.type).decode('utf-8'), cell_overwrite_ok=True)
            ws.write(0,0,'排名'.decode('utf-8'))
            ws.write(0,1,'评分'.decode('utf-8'))
            ws.write(0,2,'电影名'.decode('utf-8'))
            ws.write(0,3,'电影ID'.decode('utf-8'))
            ws.write(0,4,'类型'.decode('utf-8'))
            ws.write(0,5,'地区'.decode('utf-8'))
            ws.write(0,6,'演员'.decode('utf-8'))
            ws.write(0,7,'上映时间'.decode('utf-8'))
            ws.write(0,8,'列出演员数'.decode('utf-8'))
            ws.write(0,9,'投票数'.decode('utf-8'))

            for i in range(len(self.rank)):
                # ws = self.wb.get_sheet(self.type - 1)
                ws.write(i+1, 0, self.rank[i])
                ws.write(i+1, 1, self.rate[i])
                ws.write(i+1, 2, self.title[i])
                ws.write(i+1, 3, self.id[i])
                ws.write(i+1, 4, self.types[i])
                ws.write(i+1, 5, self.regions[i])
                ws.write(i+1, 6, self.actors[i])
                ws.write(i+1, 7, self.release_date[i])
                ws.write(i+1, 8, self.actor_count[i])
                ws.write(i+1, 9, self.vote_count[i])
        else:
            print self.type
            #self.clear()
            print 'start:',self.start
            ws = self.wb.get_sheet(self.type-1)
            for i in range(self.start + 1, len(self.rank)):
                # ws.write(i+1, 0, self.rank[i])
                # ws.write(i+1, 1, self.rate[i])
                # ws.write(i+1, 2, self.title[i])
                # ws.write(i+1, 3, self.id[i])
                # ws.write(i+1, 4, self.types[i])
                # ws.write(i+1, 5, self.regions[i])
                # ws.write(i+1, 6, self.actors[i])
                # ws.write(i+1, 7, self.release_date[i])
                # ws.write(i+1, 8, self.actor_count[i])
                # ws.write(i+1, 9, self.vote_count[i])
                ws.write(i+self.start, 0, self.rank[i])
                ws.write(i+self.start, 1, self.rate[i])
                ws.write(i+self.start, 2, self.title[i])
                ws.write(i+self.start, 3, self.id[i])
                ws.write(i+self.start, 4, self.types[i])
                ws.write(i+self.start, 5, self.regions[i])
                ws.write(i+self.start, 6, self.actors[i])
                ws.write(i+self.start, 7, self.release_date[i])
                ws.write(i+self.start, 8, self.actor_count[i])
                ws.write(i+self.start, 9, self.vote_count[i])


        #self.wb.save('douban_movies.xls')

    def clear(self):
        self.rank = []  #电影排名
        self.title = [] #存放电影名称
        self.id = []    #电影id
        self.types = [] #电影类型
        self.regions = []   #电影地区
        self.actors = []    #演员
        self.release_date = []
        self.actor_count = []
        self.vote_count = []
        self.rate = []
        self.start = 0
        self.first = 1


    def begin(self, type):
        self.clear()
        num = 0
        while self.spider(num, type) == True:
            self.first = 0
            num += 100

    def __del__(self):
        self.wb.save('douban_movies.xls')

douban = Douban_movie()
# for type in range(32):
#     douban.begin(type)

douban.begin(5)



