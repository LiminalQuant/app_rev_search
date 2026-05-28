from app_store_scraper import AppStore
from google_play_scraper import app, Sort, reviews_all
import uuid
import pandas as pd
import numpy as np
import plotly.express as px
import base64
from io import StringIO, BytesIO

def generate_excel_download_link(df_2):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df_selection.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)


g_reviews = reviews_all(
        "ru.mts.smartmed",
        sleep_milliseconds=0, # defaults to 0
        lang='ru', # defaults to 'en'
        country='us', # defaults to 'us'
        sort=Sort.NEWEST, # defaults to Sort.MOST_RELEVANT
    )
g_df = pd.DataFrame(np.array(g_reviews),columns=['review'])
g_df2 = g_df.join(pd.DataFrame(g_df.pop('review').tolist()))

g_df2.drop(columns={'userImage', 'reviewCreatedVersion'},inplace = True)
g_df2.rename(columns= {'score': 'rating','userName': 'user_name', 'reviewId': 'review_id', 'content': 'review_description', 'at': 'review_date', 'replyContent': 'developer_response', 'repliedAt': 'developer_response_date', 'thumbsUpCount': 'thumbs_up'},inplace = True)
g_df2.insert(loc=0, column='source', value='Google Play')
g_df2.insert(loc=3, column='review_title', value=None)


a_reviews = AppStore('us', 'storemed', '6736902949')
a_reviews.review()


a_df = pd.DataFrame(np.array(a_reviews.reviews), columns=['review'])
a_df2_ = a_df.join(pd.DataFrame(a_df.pop('review').tolist()))

#a_df2_.drop(columns={'isEdited'}, inplace = True)
a_df2_.insert(loc=0, column='source', value='App Store')
a_df2_['developer_response_date'] = None
a_df2_['thumbs_up'] = None
a_df2_.insert(loc=1, column='review_id', value=[uuid.uuid4() for _ in range(len(a_df2_.index))])
a_df2_.rename(columns= {'review': 'review_description','userName': 'user_name', 'date': 'review_date','title': 'review_title', 'developerResponse': 'developer_response'},inplace = True)
a_df2_ = a_df2_.where(pd.notnull(a_df2_), None)

a_df2 = pd.concat([g_df2,a_df2_])


a_df2['review_date'] = a_df2['review_date'].dt.strftime('%m/%d/%Y')
a_df2['date'] = pd.to_datetime(a_df2['review_date']).dt.floor('d')

a_df2['month'] = a_df2['date'].dt.month
a_df2['year'] = a_df2['date'].dt.year

a_df2.sort_values(by='review_date', inplace=True)

a_df2.loc[a_df2['rating'] < 4, 'рейтинг'] = 'Отрицательный'
a_df2.loc[a_df2['rating'] >= 4, 'рейтинг'] = 'Положительный'
a_df2.loc[a_df2['rating'] > 0, 'value'] = 1



import streamlit as st



st.set_page_config(page_title='Отзывы_SM', layout='wide')


year_options = a_df2['year'].unique().tolist()
month_options = a_df2['month'].unique().tolist()

st.sidebar.header('Фильтры:')

year_ = st.sidebar.multiselect("Год", options=a_df2['year'].unique(), default=2026)

# year2 = st.slider("ГОД", max_value=max(a_df2['year'].unique().tolist()), min_value=min(a_df2['year'].unique().tolist()), value=(max(a_df2['year'].unique().tolist()), min(a_df2['year'].unique().tolist())))
# year_ = st.selectbox('ВЫБИРИТЕ ГОД', year_options, 0)
market = st.sidebar.multiselect("Ресурс", options=a_df2['source'].unique(), default=a_df2['source'].unique())
# month_ = st.slider("Month", max_value=max(a_df2['month'].unique().tolist()), min_value=min(a_df2['month'].unique().tolist()), value=(max(a_df2['month'].unique().tolist()), min(a_df2['month'].unique().tolist())))
month_ = st.sidebar.multiselect("Месяц", options=a_df2['month'].unique(), default=a_df2['month'].unique())
raiting = st.sidebar.multiselect("Рейтинг", options=a_df2['рейтинг'].unique(), default=a_df2['рейтинг'].unique())



df_selection = a_df2.query("рейтинг == @raiting & year == @year_ & month == @month_ & source == @market")

st.title(":bar_chart: Основные показатели")
st.markdown('##')

average_raiting = round(df_selection['rating'].mean(), 1)
star_raiting = ":star:" * int(round(average_raiting, 0))


st.subheader("Средний Рейтинг:")
st.subheader(f"{average_raiting} {star_raiting}")

st.markdown("---")


# rating_by_value = (df_selection.groupby(by=['рейтинг']).sum()[['value']].sort_values(by="value"))
fig_rating = px.bar(df_selection,
                    x="value",
                    y='source',
                    orientation="h",
                    title="<b>Оценки</b>",
                    color_discrete_map={
                "Отрицательный": "red",
                "Положительный": "green",},
                    color='рейтинг',
                    )

st.plotly_chart(fig_rating)


fig_month = px.bar(df_selection,
                  x="month",
                  y="value",
                  orientation="v",
                  title="<b>Рейтинг за месяц</b>",
                  color_discrete_map={
                "Отрицательный": "red",
                "Положительный": "green",},
                  color="рейтинг",
                  )


st.plotly_chart(fig_month)

st.markdown("---")

# date_by_value = (df_selection.groupby(by=['date']).sum()[['value']].sort_values(by='date'))
fig_date = px.bar(df_selection,
                  x="date",
                  y="value",
                  orientation="v",
                  title="<b>Комментарии_(количество) по датам</b>",
                  color_discrete_map={
                "Отрицательный": "red",
                "Положительный": "green",},
                  color="рейтинг",
                  # template="plotly_white", 
                  )


st.plotly_chart(fig_date)




# fig_new = px.bar(df_selection,
#                  x="рейтинг",
#                  y="value",
#                  title="<b>Динамика по дням</b>",
#                  color="рейтинг",
#                  animation_frame="review_date",
#                  animation_group="рейтинг",
#                 )
# fig_new.update_layout(width=800)
# st.write(fig_new)

if st.checkbox('Сформировать файл для скачивания'):

    df_2 = st.dataframe(df_selection)
if st.checkbox('Скачать файл Excel'):

    st.subheader('ССЫЛКА ДЛЯ СКАЧИВАНИЯ')
    generate_excel_download_link(df_2)
