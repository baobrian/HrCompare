# coding=utf-8
import pandas as pd
from glob import glob
import os
from pathlib import Path

def find_files(directory, ext='xlsx'):
    return sorted(glob(directory + f'/**/*.{ext}', recursive=True))

def cmptime(t1,t2):
    return t1.timestamp()-t2.timestamp()

class CompareHrCheakin:
    def __init__(self,rawdata_path,becompare_path):
        if not rawdata_path and not becompare_path:
            raise TypeError('请检查输入的路径参数是否有误')
        self.rawdata_path=rawdata_path
        self.becompare_path=becompare_path
        assert len(self.cheakpath(rawdata_path,ext='xls'))==1, '请检查路径{}下的文档'%{self.rawdata_path}
        assert len(self.cheakpath(becompare_path))>0,'请检查路径{}下的文档'%{self.becompare_path}
    @staticmethod
    def cheakpath(path,ext='xlsx'):
        return find_files(path,ext=ext)
    def collect_cheakindata(self):
        # # # 处理打卡数据excel
        # # # 处理类型为钉钉打卡数据
        # rawfile=find_files(directory=self.rawdata_path)
        # raw_cheakin_df = pd.read_excel(rawfile[0],header=2)
        # # raw_cheakin_df['partment']=raw_cheakin_df['部门'].str[0:5]
        # mask=(raw_cheakin_df['部门'].str[0:5]!='信息科技部')
        # raw_cheakin_df=raw_cheakin_df[mask]
        # raw_cheakin_df['deadline']=pd.to_datetime('20'+raw_cheakin_df['考勤日期'].str[0:8]+' 21:30:00')
        # raw_cheakin_df['justtime'] = pd.to_datetime(raw_cheakin_df['打卡时间'])
        # raw_cheakin_df['over_time']=raw_cheakin_df.apply(lambda row :cmptime(row['justtime'],row['deadline']),axis=1)
        # raw_cheakin_df=raw_cheakin_df[(raw_cheakin_df['over_time']>=0)&(raw_cheakin_df['打卡结果']=='正常')]
        # result_df=raw_cheakin_df.groupby(['姓名'])['姓名'].size()
        # golden_cheakin_df=pd.DataFrame({'姓名':result_df.index,'晚归天数':result_df.values})
        # golden_cheakin_df = golden_cheakin_df.set_index(keys=['姓名'])
        # ########################################
        # 处理打卡数据为方舟数据
        fangzhou_file=find_files(directory=self.rawdata_path,ext='xls')
        fangzhou_df=pd.read_excel(fangzhou_file[0],header=0)
        raw_cheakin_df = fangzhou_df[((fangzhou_df['一级部门'] == '信息科技部')\
                                     |(fangzhou_df['机构'] == '信息科技部'))\
                                     &(fangzhou_df['签到时间'].notnull())\
                                     &(fangzhou_df['签退时间'].notnull())] # 8月方舟的数据只有科技使用
        raw_cheakin_df['deadline'] = pd.to_datetime(raw_cheakin_df['签到时间'].str[0:10] + ' 21:30:00')
        raw_cheakin_df['leavetime'] = pd.to_datetime(raw_cheakin_df['签退时间'])
        raw_cheakin_df['over_time'] = raw_cheakin_df.apply(lambda row: cmptime(row['leavetime'], row['deadline']),axis=1)
        raw_cheakin_df = raw_cheakin_df[(raw_cheakin_df['over_time'] >= 0) & (raw_cheakin_df['签退状态'] == '正常')]
        result_df=raw_cheakin_df.groupby(['姓名'])['姓名'].size()
        golden_cheakin_df=pd.DataFrame({'姓名':result_df.index,'晚归天数':result_df.values})
        golden_cheakin_df = golden_cheakin_df.set_index(keys=['姓名'])
        print('考勤数据汇总完毕')
        #######
        #处理部门汇总数据
        compare_sum=None
        for i,be_comparedfile in enumerate(find_files(directory=self.becompare_path)):
            filename=Path(be_comparedfile).stem
            compare_df = pd.read_excel(be_comparedfile,header=3)
            pk=[]
            for i in [2,3,16]:
                s=compare_df.columns[i]
                pk.append(s)
            df=compare_df[pk]
            df.columns=['工号','姓名','晚归天数']
            df=df[df['工号'].notnull()]
            df.fillna(0,inplace=True)
            if compare_sum is None:
                compare_sum=df
            else:
                compare_sum=compare_sum.append(df,ignore_index=True)
            final_df=compare_sum.join(golden_cheakin_df,how='left',lsuffix="部门汇总",rsuffix="考勤汇总",on='姓名')
            final_df.fillna(0,inplace=True)
            final_df['晚归数据是否一致']=((final_df['晚归天数考勤汇总'])==(final_df['晚归天数部门汇总'])).map(lambda x : '核对数据一致' if x else '核对数据不一致' )
            writer=pd.ExcelWriter(str(filename)[5::]+'___晚归核对.xlsx')
            final_df.to_excel(excel_writer=writer,sheet_name='sheet1',encoding='utf-8')
            os.chdir(r'E:\Data_temp\20200902\result\\')
            writer.save()
        print('数据核对完毕！')


if __name__ == '__main__':
    raw_cheakin= r'E:\Data_temp\20200902\fangzhou'
    be_compared= r'E:\Data_temp\20200902\becompared'
    chc=CompareHrCheakin(rawdata_path=raw_cheakin,becompare_path=be_compared)
    chc.collect_cheakindata()
    pass

