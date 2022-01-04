import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib
import seaborn as sns
import openpyxl
from openpyxl import Workbook



title = ["上永谷","港南台","平塚駅北口","秦野駅前","小田原駅前","二俣川","戸塚駅前","海老名駅前","保土ヶ谷駅前","大和駅前","逗子","相模原駅前","愛甲石田","溝ノ口駅前","緑園都市駅前","鴨宮駅前","藤沢駅南口","伊勢原駅前","鶴ヶ峰駅前","小田急相模原駅前","さがみ野駅前","綱島駅前","大倉山駅前","茅ヶ崎駅北口","上大岡駅西口"]
# titles=["上永谷","港南台","秦野","戸塚","保土ヶ谷","逗子","溝の口","上大岡"]

acts = ["向上得点分布可視化", "t"]
grades = ["全体","高3","高2","高1"]

# st.balloons()
act = st.sidebar.selectbox("Pick One", acts)

if acts.index(act) == 0:

    st.title("向上得点分布")

    # excelファイルを出力する方法を考える
    st.write("------------------------------------")
    st.write("この中はまだ作成中...")
    if st.checkbox('フォーマットを取得'):
        # wb = Workbook() #Bookインスタンス生成
        # ws = wb.active  #シートは1つ生成された状態になっている．
        # for i in title:
        #     wb.create_sheet(title=i) #新しいシートの生成
        # wb.remove(ws) #最初のwsの削除
        # wb.save("sample.xlsx") #保存；2つのシート (Sheet1, Sample Sheet)がある


        with open("sample.xlsx", "rb") as sample:
            st.download_button(label="sample", data=sample, file_name="sample.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.write("------------------------------------")



    uploaded_file=st.file_uploader("ファイルアップロード", type="xlsx")


    # グラフの表示サイズを変更する

    if uploaded_file:
        if st.checkbox('校舎ごとのグラフを表示'):
            df = pd.read_excel(uploaded_file, sheet_name=None)
            titles = list(df.keys())
            selected_prefectures = st.multiselect('校舎を選択',titles)
            for i in selected_prefectures:
                df_sheet = df[i]
                if df_sheet.empty:
                    st.write(i + "校はデータが無いよん")
                else:
                    df_score = df_sheet.iloc[2:, 6]
                    # figsize=(10, 8)

                    col1, col2 = st.columns([1,1])
                    keys = df_sheet.keys()
                    # 番号,生徒番号,生　徒　名,Unnamed: 3,現学年,生徒区分,向上得点

                    with col1:
                        st.subheader("内訳(" + i + ")")
                        st.write("人数")
                        for grade in grades[1:]:
                            st.write(grade +" : " + str(sum(df_sheet["現学年 "]==grade)))
                        st.write("最高点：" + str(df_score.max()))
                        st.write("平均点：" + str(round(df_score.mean(),2)))

                    with col2:
                        st.subheader("分布図")
                        # st.selectbox("select", grades)
                        # 学年別でグラフが出るようにしたい
                        fig = sns.displot(df_score, bins=10, rug=True, kde=True)
                        plt.title(i)
                        st.pyplot(fig,use_column_width=True)