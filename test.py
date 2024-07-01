# -*- coding: utf-8 -*-
import pandas as pd
import os 
from datetime import datetime
import matplotlib as mpl
import matplotlib.pyplot as plt
import japanize_matplotlib
import openpyxl
import warnings

warnings.simplefilter("ignore")

base_path = os.path.dirname(os.path.abspath(__file__)) #本番 (__file__)
# now_path = os.path.normpath(os.path.join(base_path, "test"))

#エクセル取り込み
questionnaire_contract = os.path.normpath(os.path.join(base_path, "契約_回答.xlsx"))
questionnaire_claim = os.path.normpath(os.path.join(base_path, "損害_回答.xlsx"))

#エクセルDF化
contract_file = pd.read_excel(questionnaire_contract, engine='openpyxl',index_col=0,header=2)
claim_file = pd.read_excel(questionnaire_claim, engine='openpyxl',index_col=0,header=2)

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
    recommend_df_all = extractRecommendData(index_reset,"0702")
    agent_num_list = index_reset.iloc[:,5].unique()
    recommend_df_agent = agentExtractRecommendData(index_reset,agent_num_list)
    recommend_df = pd.concat([recommend_df_all,recommend_df_agent],axis=1)
    fillna_0 = recommend_df.fillna(0)
    return fillna_0

claim_recommend = recommendData(claim_file)
contract_recommend = recommendData(contract_file)

# 今月のコメント
# 契約＝推奨度の理由、フリーコメント
# 損害＝推奨度の理由、フリーコメント
def commentData(file):
    copy_df = file.copy()
    index_reset = copy_df.reset_index().set_index("送信日")

    # def extractCommentData(data,agentNum):
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
claim_month_comment = commentData(claim_file)
contract_month_comment = commentData(contract_file)

#エクセル出力
current_month = str(datetime.now().month)
current_year = str(datetime.now().year)
excel_new_path = os.path.normpath(os.path.join(base_path, current_year + "_" + current_month + '_アンケート集計.xlsx'))
sheet_name = ["推奨度_契約","推奨度_損害","今月のコメント_契約","今月のコメント_損害"]
with pd.ExcelWriter(excel_new_path, engine='openpyxl') as writer:
    contract_recommend.to_excel(writer, sheet_name=sheet_name[0])
with pd.ExcelWriter(excel_new_path, mode='a', engine='openpyxl') as writer:
    claim_recommend.to_excel(writer, sheet_name=sheet_name[1])
with pd.ExcelWriter(excel_new_path, mode='a', engine='openpyxl') as writer:
    contract_month_comment.to_excel(writer, sheet_name=sheet_name[2])
with pd.ExcelWriter(excel_new_path, mode='a', engine='openpyxl') as writer:
    claim_month_comment.to_excel(writer, sheet_name=sheet_name[3])

#グラフシート作成
data_name_list = ["グラフ_契約","グラフ_損害"]
def add_gpaph_recommend(data,data_name):
    _df = data.iloc[:,[0,1]]
    if data_name == data_name_list[0]:
        list_number = 0
    else:
        list_number = 1
    graph_name = ["推奨度_契約","推奨度_損害"]

    _df.plot(
        subplots=True,
        title=graph_name[list_number],
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
    sh = wb.create_sheet(data_name)
    sh.add_image(img, "B2")
    wb.save(excel_new_path)
    wb.close()
    # os.remove(base_path + "/tmp.png")
    os.remove(image_file_path)

add_gpaph_recommend(contract_recommend,data_name_list[0])
add_gpaph_recommend(claim_recommend,data_name_list[1])
