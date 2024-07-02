# -*- coding: utf-8 -*-
import pandas as pd
import os 
from datetime import datetime
import matplotlib as mpl
import matplotlib.pyplot as plt
import japanize_matplotlib
import openpyxl
import warnings
from file_read import message_read_complete
from file_read import questionnaire_xlsx_read
from file_read import message_save_complete
warnings.simplefilter("ignore")

base_path = os.path.dirname(os.path.abspath(__file__)) #本番 (__file__)
# input_path = os.path.dirname(base_path)
# now_path = os.path.normpath(os.path.join(base_path, "test"))
current_month = str(datetime.now().month)
current_year = str(datetime.now().year)
excel_new_path = os.path.normpath(os.path.join(base_path, current_year + "_" + current_month + '_アンケート集計_【ファイル名】.xlsx'))

def readFile():
    #エクセル取り込み
    # questionnaire_file = os.path.normpath(os.path.join(base_path, "契約_回答.xlsx"))
    df = questionnaire_xlsx_read(base_path)
    #エクセルDF化
    # df = pd.read_excel(questionnaire_file, engine='openpyxl',index_col=0,header=2)
    message_read_complete()
    return df

#推奨度抽出
def recommendData(file):
    copy_df = file.copy()
    index_reset = copy_df.reset_index().set_index("送信日")
    def extractRecommendData(data,agentNum):
        recommend_column = data.loc[:,"推奨度"]
        recommend_mean = recommend_column.resample("ME").mean("推奨度").rename("平均値_" + agentNum).round(2)
        recommend_count = recommend_column.resample("ME").count().rename("個数_" + agentNum)
        recommend_df = pd.concat([recommend_mean,recommend_count], axis=1)
        return recommend_df
    def agentExtractRecommendData(data,agent_list):
        _df = pd.DataFrame()
        for i in agent_list:
            agent_detaset = data[data.loc[:,"代理店コード"] == i]
            agent_df = extractRecommendData(agent_detaset,str(i))
            _df = pd.concat([_df,agent_df], axis=1)
        return _df
    recommend_df_all = extractRecommendData(index_reset,"全体")
    agent_num_list = index_reset.iloc[:,5].unique()
    recommend_df_agent = agentExtractRecommendData(index_reset,agent_num_list)
    recommend_df = pd.concat([recommend_df_all,recommend_df_agent],axis=1)
    fillna_0 = recommend_df.fillna(0)
    reset_ind = fillna_0.reset_index()
    reset_ind['送信日'] = reset_ind['送信日'].dt.strftime('%Y/%m')
    set_ind = reset_ind.set_index('送信日')
    return set_ind
# 今月のコメント
def commentData(file):
    copy_df = file.copy()
    index_reset = copy_df.reset_index().set_index("送信日")
    comment_column = index_reset.loc[:,["推奨度の理由","フリーコメント"]]
    # '\n\n' を削除
    comment_column = comment_column.apply(lambda x: x.str.replace('\n\n', ''))
    comment_column = comment_column.apply(lambda x: x.str.replace('\n', ''))
    comment_column = comment_column.dropna(subset=["推奨度の理由", "フリーコメント"], how='all')
    comment_column = comment_column.fillna("-")
    # 送信日が今月のものだけ抽出
    current_month = datetime.now().month
    current_year = datetime.now().year
    comment_column = comment_column[comment_column.index.to_series().apply(lambda x: x.month == current_month and x.year == current_year)]
    return comment_column
# 過去のコメント
def commentPastData(file):
    copy_df = file.copy()
    index_reset = copy_df.reset_index().set_index("送信日")
    comment_column = index_reset.loc[:,["推奨度の理由","フリーコメント"]]
    # '\n\n' を削除
    comment_column = comment_column.apply(lambda x: x.str.replace('\n\n', ''))
    comment_column = comment_column.apply(lambda x: x.str.replace('\n', ''))
    comment_column = comment_column.dropna(subset=["推奨度の理由", "フリーコメント"], how='all')
    comment_column = comment_column.fillna("-")
    return comment_column
#グラフシート作成
def add_gpaph_recommend(data):
    _df = data.iloc[:,[0,1]]
    _df.plot(
        subplots=True,
        title="推奨度_グラフ",
        grid=True,
        colormap='Set1',
        legend=True,
        alpha=1.0,
        figsize=(8, 6),
        sharex=False,
        marker="o"
    )
    image_file_path = base_path + "/tmp.png"
    plt.savefig(image_file_path)
    img = openpyxl.drawing.image.Image(image_file_path)
    plt.close('all')
    # エクセル出力
    wb = openpyxl.load_workbook(excel_new_path)
    sh = wb.create_sheet("推奨度_グラフ")
    sh.add_image(img, "B2")
    wb.save(excel_new_path)
    wb.close()
    os.remove(image_file_path)
    return

def startExcelCreate():
    df = readFile()
    recommend_mean = recommendData(df)
    month_comment = commentData(df)
    past_comment = commentPastData(df)
    #エクセル出力
    sheet_name = ["推奨度","今月のコメント","過去のコメント"]
    with pd.ExcelWriter(excel_new_path, engine='openpyxl') as writer:
        recommend_mean.to_excel(writer, sheet_name=sheet_name[0])
    with pd.ExcelWriter(excel_new_path, mode='a', engine='openpyxl') as writer:
        month_comment.to_excel(writer, sheet_name=sheet_name[1])
    with pd.ExcelWriter(excel_new_path, mode='a', engine='openpyxl') as writer:
        past_comment.to_excel(writer, sheet_name=sheet_name[2])
    add_gpaph_recommend(recommend_mean)
    message_save_complete()

if __name__ == '__main__':
    startExcelCreate()