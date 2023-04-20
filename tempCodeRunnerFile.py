# 随机抽人参加活动
# 时间:2022-04-20
# 作者:金文明

import random
import time
import pandas as pd


class RandomDraw:
    # 初始化
    def __init__(self):
        self.df = pd.read_excel('名单.xlsx')
        self.list = self.df['姓名 Passport Name'].tolist()

    # 将名单写入临时文件
    def first_write_temp(self):
        with open('temp.txt', 'w') as f:
            for i in self.list:
                f.write(i + '\n')
        
    # 随机抽取人员
    def random_draw(self, list, num):
        random_draw_list = random.sample(list, num)
        return random_draw_list


    # 写入历史记录
    def write_history(self, activity_name, num, time, random_draw_list):
        with open('history.txt', 'a') as f:
            f.write('活动名称:' + activity_name + '\n')
            f.write('抽取人数:' + str(num) + '\n')
            f.write('抽取时间:' + time + '\n')
            f.write('抽取结果:' + str([i.strip('\n') for i in random_draw_list]) + '\n')
            f.write('\n')
        

if __name__ == '__main__':

    # 是否需要重新写入临时文件
    rd = RandomDraw()
    write_temp = input('是否需要重新写入临时文件(y/n):')
    if write_temp == 'y':
        # 将名单写入临时文件
        rd.first_write_temp()
    # 获取活动名称
    activity_name = input('请输入活动名称:')
    # 获取需要抽取的人数
    num = int(input('请输入需要抽取的人数:'))
    # 获取时间
    time = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    # 读取临时文件
    with open('temp.txt', 'r') as f:
        list = f.readlines()
    if len(list) < num:
        list_copy = list
        rd.first_write_temp()
        with open('temp.txt', 'r') as f:
            list = f.readlines()
        temp_list = rd.random_draw(list, num-len(list_copy))
        while set(temp_list).intersection(set(list_copy)):
            temp_list = rd.random_draw(list, num-len(list_copy))
            random_draw_list = list_copy + temp_list
        # 将抽取的人员从临时文件中删除
        with open('temp.txt', 'r') as f:
            lines = f.readlines()
        with open('temp.txt', 'w') as f_w:
            for line in lines:
                if line not in temp_list:
                    f_w.write(line)
    else:
        random_draw_list = rd.random_draw(list, num)
         # 将抽取的人员从临时文件中删除
        with open('temp.txt', 'r') as f:
            lines = f.readlines()
        with open('temp.txt', 'w') as f_w:
            for line in lines:
                if line not in random_draw_list:
                    f_w.write(line)
    
    # 输出结果
    print('活动名称:', activity_name)
    print('抽取人数:', num)
    print('抽取时间:', time)
    print('抽取结果:', [i.strip('\n') for i in random_draw_list])
    
    # 写入历史记录
    RandomDraw().write_history(activity_name, num, time, random_draw_list)
